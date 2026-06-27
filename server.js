const express = require('express');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const API_KEY = process.env.API_KEY || 'med-pro-services-2026';
const EXCEL_PATH = process.env.EXCEL_PATH || path.join(__dirname, 'Complete_Indian_Medicine_List_June_2026.xlsx');
const DB_PATH = path.join(__dirname, 'medicines.db');

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

const sampleKeys = [];
const COLUMNS = {};
const columnsList = [];
const allData = [];

const HSN_GST_MAP = {
  '3003': 12, '3004': 12, '3005': 12, '3006': 12,
  '2936': 12, '2937': 12, '2938': 12, '2939': 12,
  '2941': 12, '2942': 12, '4015': 12, '3926': 12,
  '9018': 12, '9019': 12, '9020': 12, '9021': 12, '9022': 12,
  '9027': 12, '9028': 12, '2842': 18, '2918': 18, '2915': 18,
  '2106': 12, '2108': 12, '3301': 18, '3302': 18, '3303': 18,
  '3304': 18, '3305': 18, '3306': 18, '3307': 18,
  '3401': 12, '3402': 12, '3822': 12, '2202': 12, '1211': 5
};

function gstFromHSN(hsn) {
  if (!hsn) return '';
  const clean = hsn.replace(/[\s\-]/g, '');
  for (let len = 4; len >= 2; len--) {
    const prefix = clean.slice(0, len);
    if (HSN_GST_MAP[prefix]) return String(HSN_GST_MAP[prefix]);
  }
  return '';
}

function initDB() {
  console.log('Loading Excel file...');
  const workbook = XLSX.readFile(EXCEL_PATH, {
    cellFormula: false,
    cellHTML: false,
    cellNF: false,
    cellText: false
  });
  const sheetName = workbook.SheetNames.find(s => s.includes('All Medicines')) || workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const totalRows = range.e.r - range.s.r;

  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: range.s.r, c });
    headers[c] = String(ws[addr] ? ws[addr].v : '');
  }
  Object.assign(sampleKeys, headers.filter(Boolean));

  function pick(...names) {
    for (const n of names) {
      const match = sampleKeys.find(k => k.toLowerCase().replace(/[\s\-/]/g, '') === n.toLowerCase().replace(/[\s\-/]/g, ''));
      if (match) return match;
      const partial = sampleKeys.find(k => k.toLowerCase().includes(n.toLowerCase().split(' ')[0]));
      if (partial) return partial;
    }
    return names[0];
  }

  const colMap = {
    name: pick('Medicine Name', 'name', 'Medicine'),
    mrp: pick('MRP (Rs.)', 'MRP', 'mrp'),
    manufacturer: pick('Manufacturer', 'manufacturer'),
    type: pick('Type'),
    pack: pick('Pack Size', 'pack_size'),
    comp1: pick('Composition 1', 'Composition', 'comp1'),
    comp2: pick('Composition 2', 'comp2'),
    available: pick('Available'),
    therapeutic: pick('Therapeutic Class', 'therapeutic_class'),
    usage: pick('Usage / Indication', 'Usage', 'usage'),
    substitutes: pick('Substitutes'),
    sideEffects: pick('Side Effects', 'sideEffects'),
    chemicalClass: pick('Chemical Class'),
    habitForming: pick('Habit Forming'),
    actionClass: pick('Action Class'),
    schedule: pick('Schedule', 'Drug Schedule'),
    dosage: pick('Dosage Form', 'dosage_form', 'dosage'),
    strength: pick('Strength', 'strength'),
    contraindications: pick('Contraindications', 'contraindications', 'contra'),
    interactions: pick('Drug Interactions', 'interactions', 'drug_interactions'),
    pregCategory: pick('Pregnancy Category', 'pregnancy_category', 'preg'),
    storage: pick('Storage Conditions', 'storage_conditions', 'storage'),
    license: pick('License No.', 'Manufacturer License No.', 'license_no', 'license'),
    unitPrice: pick('Price per Unit', 'price_per_unit', 'unit_price'),
    hsn: pick('HSN Code', 'HSN', 'hsn'),
    gst: pick('GST (%)', 'GST', 'gst'),
    marketYear: pick('Market Year', 'Launch Year', 'launch_year'),
    imageUrl: pick('Image URL', 'image_url', 'Image'),
    alcohol: pick('Alcohol Interaction', 'alcohol_interaction', 'Alcohol'),
    driving: pick('Driving Warning', 'driving_warning', 'Driving'),
    majorCat: pick('Major Category', 'major_category', 'Category'),
    therapeuticCat: pick('Therapeutic_Category', 'therapeutic_category'),
    jaAlt: pick('Jan Aushadhi Alternative'),
    jaMRP: pick('Jan Aushadhi MRP'),
    jaSavings: pick('Jan Aushadhi Savings')
  };

  Object.assign(COLUMNS, colMap);
  columnsList.push(...Object.keys(colMap));

  const headerIndex = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    headerIndex[headers[c]] = c;
  }

  const colIndices = {};
  for (const [key, header] of Object.entries(colMap)) {
    colIndices[key] = headerIndex[header];
  }

  const db = new Database(DB_PATH);
  db.exec('PRAGMA journal_mode=WAL');
  db.exec('PRAGMA synchronous=OFF');
  const colNames = columnsList.map(c => '"' + c + '"').join(', ');
  const placeholders = columnsList.map(() => '?').join(', ');
  db.exec(`CREATE TABLE IF NOT EXISTS medicines (id INTEGER PRIMARY KEY, ${colNames})`);
  db.exec('CREATE INDEX IF NOT EXISTS idx_name ON medicines(name)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_comp1 ON medicines(comp1)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_usage ON medicines(usage)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_manufacturer ON medicines(manufacturer)');
  db.exec('CREATE INDEX IF NOT EXISTS idx_therapeutic ON medicines(therapeutic)');

  const insert = db.prepare(`INSERT INTO medicines (${colNames}) VALUES (${placeholders})`);
  let count = 0;

  const insertBatch = db.transaction(() => {
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      const vals = columnsList.map(c => {
        const ci = colIndices[c];
        if (ci === undefined) return '';
        const addr = XLSX.utils.encode_cell({ r, c: ci });
        const cell = ws[addr];
        let v = cell ? String(cell.v) : '';
        if (c === 'gst' && !v) {
          const hsnIdx = colIndices['hsn'];
          if (hsnIdx !== undefined) {
            const hsnAddr = XLSX.utils.encode_cell({ r, c: hsnIdx });
            const hsnCell = ws[hsnAddr];
            const hsn = hsnCell ? String(hsnCell.v) : '';
            v = gstFromHSN(hsn);
          }
        }
        return v;
      });
      insert.run(vals);
      count++;
    }
  });
  insertBatch();

  console.log(`Inserted ${count} records into SQLite`);
  return db;
}

let db;
try {
  db = initDB();
} catch (e) {
  console.error('Failed to initialize database:', e.message);
  process.exit(1);
}

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

app.get('/', (req, res) => {
  res.redirect('/api/health?key=' + encodeURIComponent(API_KEY));
});

app.get('/api/search', (req, res) => {
  const q = (req.query.q || '').toLowerCase().trim();
  const fieldsParam = req.query.fields || 'name';
  const limit = Math.min(parseInt(req.query.limit) || 200, 500);

  if (!q) {
    return res.json({ query: q, total: 0, results: [], time: '0ms' });
  }

  const searchFields = fieldsParam.split(',').map(f => f.trim()).filter(Boolean);
  const t0 = Date.now();

  const likeQ = q.replace(/'/g, "''");
  const conditions = searchFields.map(f => `"${f}" LIKE '%' || ? || '%'`).join(' OR ');
  if (!conditions) {
    return res.json({ query: q, total: 0, results: [], time: (Date.now() - t0) + 'ms' });
  }

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
  const count = db.prepare('SELECT COUNT(*) as c FROM medicines').get().c;
  res.json({ status: 'ok', records: count, version: '2.0.0' });
});

app.listen(PORT, () => {
  console.log(`Medicine API running on port ${PORT}`);
  console.log(`API Key: ${API_KEY}`);
  const count = db.prepare('SELECT COUNT(*) as c FROM medicines').get().c;
  console.log(`Records: ${count}`);
});
