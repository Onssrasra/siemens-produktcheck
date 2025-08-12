const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');

const { parseDimensionsToLBH } = require('./utils');
const { SiemensProductScraper } = require('./scraper');

const app = express();
const PORT = process.env.PORT || 3000;
const SCRAPE_CONCURRENCY = Number(process.env.SCRAPE_CONCURRENCY || 6);

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// ---- Coloring ----
const FILL = {
  GREEN:  { type:'pattern', pattern:'solid', fgColor:{ argb:'FFD5F4E6' } }, // green
  RED:    { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFDEAEA' } }, // red
  ORANGE: { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF4CC' } }, // orange
};

function setFill(cell, name) {
  if (!cell) return;
  const f = FILL[name];
  if (f) cell.fill = f;
}

// ---- Helpers ----
const HEADER_MAIN_ROW = 3;
const HEADER_SUB_ROW  = 4;
const DATA_START_ROW  = 5;

function normMainHeader(v) {
  const s = String(v || '').trim();
  return s.replace(/\s*-\s*/g, '-'); // normalise hyphen spacing
}

function findPairs(ws) {
  const wanted = [
    'Materialkurztext', 'Material-Kurztext',
    'Her.-Artikelnummer',
    'Fert./Prüfhinweis',
    'Werkstoff',
    'Nettogewicht',
    'Länge','Breite','Höhe'
  ];
  const pairs = {};
  const row3 = ws.getRow(HEADER_MAIN_ROW);
  const row4 = ws.getRow(HEADER_SUB_ROW);
  for (let c = 1; c <= ws.columnCount; c++) {
    const main = normMainHeader(row3.getCell(c).value);
    if (!wanted.includes(main)) continue;
    const left = String(row4.getCell(c).value || '').trim();
    const right = String(row4.getCell(c+1).value || '').trim();
    if (left === 'DB-Wert' && right === 'Web-Wert') {
      // prefer canonical names
      const key = main === 'Material-Kurztext' ? 'Materialkurztext' : main;
      pairs[key] = { db: c, web: c + 1 };
    }
  }
  return pairs;
}

// Find the longest A2Vxxxx... token in a row (any column)
function findA2VInRow(ws, rowIndex) {
  const row = ws.getRow(rowIndex);
  let best = null;
  for (let c = 1; c <= ws.columnCount; c++) {
    const v = row.getCell(c).value;
    if (!v) continue;
    const m = String(v).toUpperCase().match(/A2V\d+/g);
    if (m && m.length) {
      const longest = m.reduce((a,b)=> (a.length>=b.length?a:b));
      if (!best || longest.length > best.length) best = longest;
    }
  }
  return best;
}

// Parsing numbers (no tolerance)
function toNumStrict(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  const s = String(v).trim().replace(',', '.');
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

// Parse weight text to kg (number) or null. No tolerance in comparison.
function parseWeightKg(text) {
  if (text == null) return null;
  const s = String(text).trim();
  if (s === '') return null;
  // If pure number: assume kg
  if (/^-?\d+(?:[.,]\d+)?$/.test(s)) {
    return Number(s.replace(',', '.'));
  }
  const kg = s.match(/(-?\d+(?:[.,]\d+)?)\s*kg\b/i);
  if (kg) return Number(kg[1].replace(',', '.'));
  const g = s.match(/(-?\d+(?:[.,]\d+)?)\s*g\b/i);
  if (g) return Number(g[1].replace(',', '.'))/1000;
  const t = s.match(/(-?\d+(?:[.,]\d+)?)\s*t\b/i);
  if (t) return Number(t[1].replace(',', '.'))*1000;
  return null;
}

function eqTextStrict(dbVal, webVal) {
  // exact match after simple trim only (case-sensitive)
  const a = dbVal == null ? '' : String(dbVal).trim();
  const b = webVal == null ? '' : String(webVal).trim();
  return a === b;
}
function eqNumberStrict(dbVal, webVal) {
  if (dbVal == null || webVal == null) return false;
  const a = toNumStrict(dbVal);
  const b = toNumStrict(webVal);
  if (a == null || b == null) return false;
  return a === b;
}

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/api/health', (req, res) => res.json({ ok: true, time: new Date().toISOString() }));

app.post('/api/process-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    // Collect rows with A2V and build tasks
    const rowsPerSheet = new Map();
    const pairsPerSheet = new Map();
    const allIds = [];

    for (const ws of wb.worksheets) {
      const pairs = findPairs(ws);
      pairsPerSheet.set(ws, pairs);
      const indices = [];
      const last = ws.lastRow?.number || 0;
      for (let r = DATA_START_ROW; r <= last; r++) {
        const id = findA2VInRow(ws, r);
        if (id) { indices.push(r); allIds.push(id); }
      }
      rowsPerSheet.set(ws, indices);
    }

    const scraper = new SiemensProductScraper();
    const resultsMap = await scraper.scrapeMany(allIds, SCRAPE_CONCURRENCY);

    // Fill in-place into Web columns, then color (green/red/orange)
    for (const ws of wb.worksheets) {
      const rows = rowsPerSheet.get(ws) || [];
      const pairs = pairsPerSheet.get(ws) || {};
      for (const r of rows) {
        const id = findA2VInRow(ws, r);
        const web = (id && resultsMap.get(id)) || {};

        // --- Materialkurztext ---
        if (pairs['Materialkurztext']) {
          const dbCol = pairs['Materialkurztext'].db;
          const webCol = pairs['Materialkurztext'].web;
          const dbVal = ws.getRow(r).getCell(dbCol).value;
          const webVal = web.Produkttitel || '';
          const c = ws.getRow(r).getCell(webCol);
          c.value = webVal;
          if (webVal === '' || dbVal == null || dbVal === '') setFill(c, 'ORANGE');
          else setFill(c, eqTextStrict(dbVal, webVal) ? 'GREEN' : 'RED');
        }

        // --- Her.-Artikelnummer (Web: "Weitere Artikelnummer") ---
        if (pairs['Her.-Artikelnummer']) {
          const dbCol = pairs['Her.-Artikelnummer'].db;
          const webCol = pairs['Her.-Artikelnummer'].web;
          const dbVal = ws.getRow(r).getCell(dbCol).value;
          const webVal = web['Weitere Artikelnummer'] || '';
          const c = ws.getRow(r).getCell(webCol);
          c.value = webVal;
          if (webVal === '' || dbVal == null || dbVal === '') setFill(c, 'ORANGE');
          else setFill(c, eqTextStrict(dbVal, webVal) ? 'GREEN' : 'RED');
        }

        // --- Fert./Prüfhinweis ---
        if (pairs['Fert./Prüfhinweis']) {
          const dbCol = pairs['Fert./Prüfhinweis'].db;
          const webCol = pairs['Fert./Prüfhinweis'].web;
          const dbVal = ws.getRow(r).getCell(dbCol).value;
          const webVal = web['Fert./Prüfhinweis'] || '';
          const c = ws.getRow(r).getCell(webCol);
          c.value = webVal;
          if (webVal === '' || dbVal == null || dbVal === '') setFill(c, 'ORANGE');
          else setFill(c, eqTextStrict(dbVal, webVal) ? 'GREEN' : 'RED');
        }

        // --- Werkstoff ---
        if (pairs['Werkstoff']) {
          const dbCol = pairs['Werkstoff'].db;
          const webCol = pairs['Werkstoff'].web;
          const dbVal = ws.getRow(r).getCell(dbCol).value;
          const webVal = web['Werkstoff'] || '';
          const c = ws.getRow(r).getCell(webCol);
          c.value = webVal;
          if (webVal === '' || dbVal == null || dbVal === '') setFill(c, 'ORANGE');
          else setFill(c, eqTextStrict(dbVal, webVal) ? 'GREEN' : 'RED');
        }

        // --- Nettogewicht (kg) ---
        if (pairs['Nettogewicht']) {
          const dbCol = pairs['Nettogewicht'].db;
          const webCol = pairs['Nettogewicht'].web;
          const dbCell = ws.getRow(r).getCell(dbCol);
          const webCell = ws.getRow(r).getCell(webCol);

          const dbKg = parseWeightKg(dbCell.value);
          const webKg = parseWeightKg(web['Gewicht']);

          webCell.value = (webKg != null ? Number(webKg) : null);
          webCell.numFmt = '0.00';

          if (webKg == null || dbKg == null) setFill(webCell, 'ORANGE');
          else setFill(webCell, (dbKg === webKg) ? 'GREEN' : 'RED');
        }

        // --- Abmessungen (L/B/H in mm) ---
        const dims = web['Abmessung'] ? parseDimensionsToLBH(web['Abmessung']) : {L:null,B:null,H:null};

        if (pairs['Länge']) {
          const dbCol = pairs['Länge'].db;
          const webCol = pairs['Länge'].web;
          const dbVal = toNumStrict(ws.getRow(r).getCell(dbCol).value);
          const webVal = (dims.L != null ? dims.L : null);
          const c = ws.getRow(r).getCell(webCol);
          c.value = webVal;
          c.numFmt = '0';
          if (webVal == null || dbVal == null) setFill(c, 'ORANGE');
          else setFill(c, (dbVal === webVal) ? 'GREEN' : 'RED');
        }
        if (pairs['Breite']) {
          const dbCol = pairs['Breite'].db;
          const webCol = pairs['Breite'].web;
          const dbVal = toNumStrict(ws.getRow(r).getCell(dbCol).value);
          const webVal = (dims.B != null ? dims.B : null);
          const c = ws.getRow(r).getCell(webCol);
          c.value = webVal;
          c.numFmt = '0';
          if (webVal == null || dbVal == null) setFill(c, 'ORANGE');
          else setFill(c, (dbVal === webVal) ? 'GREEN' : 'RED');
        }
        if (pairs['Höhe']) {
          const dbCol = pairs['Höhe'].db;
          const webCol = pairs['Höhe'].web;
          const dbVal = toNumStrict(ws.getRow(r).getCell(dbCol).value);
          const webVal = (dims.H != null ? dims.H : null);
          const c = ws.getRow(r).getCell(webCol);
          c.value = webVal;
          c.numFmt = '0';
          if (webVal == null || dbVal == null) setFill(c, 'ORANGE');
          else setFill(c, (dbVal === webVal) ? 'GREEN' : 'RED');
        }
      }
    }

    const out = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="DB_Produktvergleich_verarbeitet.xlsx"');
    res.send(Buffer.from(out));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || 'Internal error' });
  }
});

app.listen(PORT, () => console.log(`Server running at http://0.0.0.0:${PORT}`));