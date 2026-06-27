const express = require('express');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');
const zlib = require('zlib');
const readline = require('readline');

const app = express();
const PORT = process.env.PORT || 3000;
const API_KEY = process.env.API_KEY || 'med-pro-services-2026';
const DB_PATH = path.join(__dirname, 'medicines.db');
const TSV_PATH = path.join(__dirname, 'medicines.tsv.gz');

const limiter = rateLimit({
  windowMs: 60 * 1000,
  max: 120,
  message: { error: 'Too many requests. Please try again later.' }
});
app.use('/api/', limiter);
app.use(cors());

app.use('/api/', (req, res, next) => {
  const key = req.query.key || req.headers['x-api-key'];
  if (key !== API_KEY) {
    return res.status(401).json({ error: 'Invalid or missing API key' });
  }
  next();
});

const columns = ['name','mrp','manufacturer','type','pack','comp1','comp2','available','therapeutic','usage','substitutes','sideEffects','chemicalClass','habitForming','actionClass','schedule','dosage','strength','contraindications','interactions','pregCategory','storage','license','unitPrice','hsn','gst','marketYear','imageUrl','alcohol','driving','majorCat','therapeuticCat','jaAlt','jaMRP','jaSavings'];

function esc(s) {
  if (!s) return '';
  return String(s).replace(/\\n/g, '\n').replace(/\\t/g, '\t').replace(/\\\\/g, '\\');
}

async function initDB() {
  const needsBuild = !fs.existsSync(DB_PATH) || fs.statSync(DB_PATH).size === 0;
  const db = new Database(DB_PATH);
  db.exec('PRAGMA journal_mode=WAL');
  db.exec('PRAGMA synchronous=OFF');

  const colNames = columns.map(c => '"' + c + '"').join(', ');
  const placeholders = columns.map(() => '?').join(', ');
  db.exec(`CREATE TABLE IF NOT EXISTS medicines (id INTEGER PRIMARY KEY, ${colNames})`);

  const rowCount = db.prepare('SELECT COUNT(*) as c FROM medicines').get().c;
  if (rowCount > 0) return db;

  db.exec('DROP INDEX IF EXISTS idx_name');
  db.exec('DROP INDEX IF EXISTS idx_comp1');
  db.exec('DROP INDEX IF EXISTS idx_manufacturer');
  db.exec('DROP INDEX IF EXISTS idx_usage');
  db.exec('DROP INDEX IF EXISTS idx_therapeutic');

  const insert = db.prepare(`INSERT INTO medicines (${colNames}) VALUES (${placeholders})`);
  let count = 0;

  const insertBatch = db.transaction((rows) => {
    for (const vals of rows) {
      insert.run(vals);
      count++;
    }
  });

  const stream = fs.createReadStream(TSV_PATH).pipe(zlib.createGunzip());
  const rl = readline.createInterface({ input: stream, crlfDelay: Infinity });
  let header = true;
  let batch = [];

  console.log('Loading database from TSV...');
  const t0 = Date.now();

  for await (const line of rl) {
    if (header) { header = false; continue; }
    const parts = line.split('\t');
    const vals = columns.map((c, i) => i < parts.length ? esc(parts[i]) : '');
    batch.push(vals);
    if (batch.length >= 5000) {
      insertBatch(batch);
      batch = [];
    }
  }
  if (batch.length) insertBatch(batch);

  db.exec('CREATE INDEX IF NOT EXISTS idx_name ON medicines(name)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_comp1 ON medicines(comp1)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_usage ON medicines(usage)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_manufacturer ON medicines(manufacturer)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_therapeutic ON medicines(therapeutic)');
  db.exec('ANALYZE');

  console.log(`Loaded ${count} records in ${Date.now() - t0}ms`);
  return db;
}

let db;
initDB().then(d => {
  db = d;
  app.listen(PORT, () => {
    const count = db.prepare('SELECT COUNT(*) as c FROM medicines').get().c;
    console.log(`Medicine API running on port ${PORT}`);
    console.log(`Records: ${count}`);
  });
}).catch(e => {
  console.error('Failed to initialize:', e);
  process.exit(1);
});

function lookupName(name) {
  if (!name) return null;
  const row = db.prepare('SELECT mrp, manufacturer FROM medicines WHERE name = ? LIMIT 1').get(name);
  if (row) return { mrp: row.mrp || '', mfr: row.manufacturer || '' };
  const row2 = db.prepare(`SELECT mrp, manufacturer FROM medicines WHERE name LIKE ? LIMIT 1`).get('%' + name + '%');
  if (row2) return { mrp: row2.mrp || '', mfr: row2.manufacturer || '' };
  return null;
}

function attachSubsPrices(results) {
  const getSubs = db.prepare(`SELECT name, mrp, manufacturer FROM medicines WHERE name = ? LIMIT 1`);
  for (const r of results) {
    if (r.substitutes) {
      const subs = r.substitutes.split(', ').filter(s => s.trim());
      r.subsWithPrices = [];
      for (const s of subs) {
        let info = getSubs.get(s);
        if (!info) {
          info = getSubs.get(s.replace(/\s+\d+.*$/, ''));
        }
        r.subsWithPrices.push({
          name: s,
          mrp: info && info.mrp ? info.mrp : '',
          mfr: info && info.manufacturer ? info.manufacturer : ''
        });
      }
    }
  }
  return results;
}

app.use('/api/', (req, res, next) => {
  if (!db) return res.status(503).json({ error: 'Database is loading, try again in a moment' });
  next();
});

app.get('/', (req, res) => {
  res.redirect('/api/health?key=' + encodeURIComponent(API_KEY));
});

app.get('/api/search', (req, res) => {
  const q = (req.query.q || '').toLowerCase().trim();
  const fieldsParam = req.query.fields || 'name';
  const limit = Math.min(parseInt(req.query.limit) || 200, 500);

  if (!q || !db) {
    return res.json({ query: q, total: 0, results: [], time: '0ms' });
  }

  const searchFields = fieldsParam.split(',').map(f => f.trim()).filter(Boolean);
  const t0 = Date.now();

  const likeQ = q.replace(/'/g, "''").replace(/\\/g, '\\\\');
  const conditions = searchFields.map(f => `"${f}" LIKE '%' || ? || '%'`).join(' OR ');

  const params = searchFields.map(() => likeQ);
  const all = db.prepare(`SELECT * FROM medicines WHERE ${conditions} LIMIT 500`).all(...params);

  const exact = [], starts = [], contains = [];
  for (const d of all) {
    let matchedVal = '';
    for (const f of searchFields) {
      const val = (d[f] || '').toLowerCase();
      if (val.indexOf(q) !== -1) {
        matchedVal = val;
        break;
      }
    }
    if (matchedVal === q || matchedVal.startsWith(q + ' ')) exact.push(d);
    else if (matchedVal.startsWith(q)) starts.push(d);
    else contains.push(d);
  }

  const results = attachSubsPrices(exact.concat(starts, contains).slice(0, limit));
  const time = Date.now() - t0;

  res.json({
    query: q,
    total: results.length,
    results,
    time: time + 'ms',
    fields: searchFields
  });
});

app.get('/api/lookup', (req, res) => {
  const names = (req.query.names || '').split(',').map(s => s.trim()).filter(Boolean);
  const t0 = Date.now();
  const result = {};
  const stmt = db.prepare('SELECT mrp, manufacturer FROM medicines WHERE name = ? LIMIT 1');
  for (const name of names) {
    const row = stmt.get(name);
    if (row) result[name] = { mrp: row.mrp || '', mfr: row.manufacturer || '' };
  }
  res.json({ names: result, time: (Date.now() - t0) + 'ms' });
});

app.get('/api/health', (req, res) => {
  if (!db) return res.json({ status: 'loading', records: 0, version: '2.0.0' });
  const count = db.prepare('SELECT COUNT(*) as c FROM medicines').get().c;
  res.json({ status: 'ok', records: count, version: '2.0.0' });
});
