const express = require('express');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const API_KEY = process.env.API_KEY || 'med-pro-services-2026';
const EXCEL_PATH = process.env.EXCEL_PATH || path.join(__dirname, 'Complete_Indian_Medicine_List_June_2026.xlsx');

// Rate limiting
const limiter = rateLimit({
  windowMs: 60 * 1000,
  max: 120,
  message: { error: 'Too many requests. Please try again later.' }
});
app.use('/api/', limiter);
app.use(cors());

// API key middleware
app.use('/api/', (req, res, next) => {
  const key = req.query.key || req.headers['x-api-key'];
  if (key !== API_KEY) {
    return res.status(401).json({ error: 'Invalid or missing API key' });
  }
  next();
});

console.log('Loading Excel file...');
const workbook = XLSX.readFile(EXCEL_PATH);
const sheetName = workbook.SheetNames.find(s => s.includes('All Medicines')) || workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const rawData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
console.log(`Loaded ${rawData.length} records`);

const sampleKeys = rawData.length ? Object.keys(rawData[0]) : [];
function pick(...names) {
  for (const n of names) {
    const match = sampleKeys.find(k => k.toLowerCase().replace(/[\s\-/]/g,'') === n.toLowerCase().replace(/[\s\-/]/g,''));
    if (match) return match;
    const partial = sampleKeys.find(k => k.toLowerCase().includes(n.toLowerCase().split(' ')[0]));
    if (partial) return partial;
  }
  return names[0];
}

const COLUMNS = {
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

// Map all data into flat objects
const allData = rawData.map((row, idx) => {
  const d = { id: idx };
  for (const [key, col] of Object.entries(COLUMNS)) {
    d[key] = String(row[col] || '');
  }
  return d;
});

// Auto-fill GST from HSN
const HSN_GST_MAP = {
  '3003': 12, '3004': 12, '3005': 12, '3006': 12,
  '2936': 12, '2937': 12, '2938': 12, '2939': 12,
  '2941': 12, '2942': 12, '4015': 12, '3926': 12,
  '9018': 12, '9019': 12, '9020': 12, '9021': 12, '9022': 12,
  '9027': 12, '9028': 12, '2842': 18, '2918': 18, '2915': 18,
  '2106': 12, '2108': 12,
  '3301': 18, '3302': 18, '3303': 18, '3304': 18, '3305': 18, '3306': 18, '3307': 18,
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
for (const d of allData) {
  if (!d.gst && d.hsn) {
    const rate = gstFromHSN(d.hsn);
    if (rate) d.gst = rate;
  }
}

// Build name index for substitute price lookups
const nameIndex = {};
for (const d of allData) {
  const key = d.name.toLowerCase().trim();
  if (!nameIndex[key]) nameIndex[key] = {};
  if (d.mrp) nameIndex[key].mrp = d.mrp;
  if (d.manufacturer) nameIndex[key].mfr = d.manufacturer;
}

function lookupName(name) {
  if (!name) return null;
  const key = name.toLowerCase().trim();
  if (nameIndex[key]) return nameIndex[key];
  for (const [k, v] of Object.entries(nameIndex)) {
    if (k.includes(key) || key.includes(k)) return v;
  }
  return null;
}

// Pre-attach substitute prices to results
function attachSubsPrices(results) {
  for (const r of results) {
    if (r.substitutes) {
      const subs = r.substitutes.split(', ').filter(s => s.trim());
      r.subsWithPrices = [];
      for (const s of subs) {
        const info = lookupName(s);
        r.subsWithPrices.push({
          name: s,
          mrp: info && info.mrp ? info.mrp : '',
          mfr: info && info.mfr ? info.mfr : ''
        });
      }
    }
  }
  return results;
}

// Search endpoint
app.get('/api/search', (req, res) => {
  const q = (req.query.q || '').toLowerCase().trim();
  const fieldsParam = req.query.fields || 'name';
  const limit = Math.min(parseInt(req.query.limit) || 200, 500);

  if (!q) {
    return res.json({ query: q, total: allData.length, results: [], time: '0ms' });
  }

  const searchFields = fieldsParam.split(',').map(f => f.trim()).filter(Boolean);
  const t0 = Date.now();

  const exact = [], starts = [], contains = [];
  for (let i = 0; i < allData.length; i++) {
    const d = allData[i];
    let matched = false, matchedVal = '';
    for (let s = 0; s < searchFields.length; s++) {
      const val = d[searchFields[s]];
      if (val && val.toLowerCase().indexOf(q) !== -1) {
        matched = true;
        matchedVal = val.toLowerCase();
        break;
      }
    }
    if (!matched) continue;

    if (matchedVal === q || matchedVal.startsWith(q + ' ')) exact.push(d);
    else if (matchedVal.startsWith(q)) starts.push(d);
    else contains.push(d);

    if (exact.length + starts.length + contains.length > 500) break;
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

// Substitute lookup endpoint
app.get('/api/lookup', (req, res) => {
  const names = (req.query.names || '').split(',').map(s => s.trim()).filter(Boolean);
  const t0 = Date.now();
  const result = {};
  for (const name of names) {
    const info = lookupName(name);
    if (info) result[name] = info;
  }
  res.json({ names: result, time: (Date.now() - t0) + 'ms' });
});

// Root redirect
app.get('/', (req, res) => {
  res.redirect('/api/health?key=' + encodeURIComponent(API_KEY));
});

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', records: allData.length, version: '1.0.0' });
});

app.listen(PORT, () => {
  console.log(`Medicine API running on port ${PORT}`);
  console.log(`API Key: ${API_KEY}`);
  console.log(`Records: ${allData.length}`);
});
