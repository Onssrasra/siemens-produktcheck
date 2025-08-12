const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
const multer = require('multer');
const ExcelJS = require('exceljs');

const {
  toNumber,
  parseWeight,
  weightToKg,
  parseDimensionsToLBH,
  normPartNo,
  mapMaterialClassificationToExcel,
  normalizeNCode
} = require('./utils');
const { SiemensProductScraper, a2vUrl } = require('./scraper');

const app = express();
const PORT = process.env.PORT || 3000;
const SCRAPE_CONCURRENCY = Number(process.env.SCRAPE_CONCURRENCY || 6);
const WEIGHT_TOL_PCT = Number(process.env.WEIGHT_TOL_PCT || 0); // 0 = strikt
const COLS = { Z:'Z', E:'E', C:'C', S:'S', T:'T', U:'U', V:'V', W:'W', P:'P', N:'N' };
const HEADER_ROW = 3;
const FIRST_DATA_ROW = 5;

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const scraper = new SiemensProductScraper();

function fillColor(ws, addr, color) {
  if (!color) return;
  const map = { green: 'FFD5F4E6', red: 'FFFDEAEA', orange: 'FFFFF4CC' };
  ws.getCell(addr).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: map[color] || map.green } };
}

// Gleichheitstests (strikt, aber mit Normalisierung)
function eqText(a,b) {
  if (a == null || b == null) return false;
  const A = String(a).trim().toLowerCase().replace(/\s+/g,' ');
  const B = String(b).trim().toLowerCase().replace(/\s+/g,' ');
  return A === B;
}
function eqPart(a,b) { return normPartNo(a) === normPartNo(b); }
function eqN(a,b) { return normalizeNCode(a) === normalizeNCode(b); }
function eqDim(exU, exV, exW, webTxt) {
  const L = toNumber(exU), B = toNumber(exV), H = toNumber(exW);
  const w = parseDimensionsToLBH(webTxt);
  if (L==null || B==null || H==null || w.L==null || w.B==null || w.H==null) return false;
  return L === w.L && B === w.B && H === w.H;
}
function eqWeight(exS, exT, webVal) {
  const { value: wv, unit: wu } = parseWeight(webVal);
  if (wv == null) return false;
  const exNum = toNumber(exS);
  const exUnit = (exT||'').toString().trim().toLowerCase();
  if (exNum == null) return false;
  const a = weightToKg(exNum, exUnit);
  const b = weightToKg(wv, wu || exUnit || 'kg');
  if (a == null || b == null) return false;
  // strikt
  return Math.abs(a - b) < 1e-9;
}

// Routes
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/api/health', (req, res) => res.json({ ok: true, time: new Date().toISOString() }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.post('/api/process-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    // 1) A2V-Nummern aus Spalte Z ab Zeile 4 einsammeln
    
const tasks = [];
const rowsPerSheet = new Map(); // ws -> [rowIndex,...]
const pairsPerSheet = new Map(); // ws -> pairs

for (const ws of wb.worksheets) {
  const pairs = findPairs(ws);
  pairsPerSheet.set(ws, pairs);
  const indices = [];
  const last = ws.lastRow?.number || 0;
  for (let r = FIRST_DATA_ROW; r <= last; r++) {
    const a2v = findA2VInRow(ws, r);
    if (a2v) { indices.push(r); tasks.push(a2v); }
  }
  rowsPerSheet.set(ws, indices);
}


    // 2) Scrapen
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);

    
// 3) Web-Werte in die "Web-Wert"-Spalten derselben Zeile schreiben (keine Zeileneinfügung)
for (const ws of wb.worksheets) {
  const prodRows = rowsPerSheet.get(ws) || [];
  const pairs = pairsPerSheet.get(ws) || {};
  for (const r of prodRows) {
    const a2v = findA2VInRow(ws, r);
    const web = (a2v && resultsMap.get(a2v)) || {};

    // --- Werte schreiben ---
    // Materialkurztext
    if (pairs['Materialkurztext']) {
      const c = ws.getRow(r).getCell(pairs['Materialkurztext'].web);
      c.value = web.Produkttitel || '';
      compareTextAndColor(ws, r, pairs['Materialkurztext'].db, pairs['Materialkurztext'].web);
    }

    // Her.-Artikelnummer  (aus Web: "Weitere Artikelnummer")
    if (pairs['Her.-Artikelnummer']) {
      const c = ws.getRow(r).getCell(pairs['Her.-Artikelnummer'].web);
      c.value = web['Weitere Artikelnummer'] || '';
      compareTextAndColor(ws, r, pairs['Her.-Artikelnummer'].db, pairs['Her.-Artikelnummer'].web);
    }

    // Fert./Prüfhinweis (falls vorhanden, sonst bleibt leer -> orange)
    if (pairs['Fert./Prüfhinweis']) {
      const c = ws.getRow(r).getCell(pairs['Fert./Prüfhinweis'].web);
      c.value = web['Fert./Prüfhinweis'] || '';
      compareTextAndColor(ws, r, pairs['Fert./Prüfhinweis'].db, pairs['Fert./Prüfhinweis'].web);
    }

    // Werkstoff
    if (pairs['Werkstoff']) {
      const c = ws.getRow(r).getCell(pairs['Werkstoff'].web);
      c.value = web['Werkstoff'] || '';
      compareTextAndColor(ws, r, pairs['Werkstoff'].db, pairs['Werkstoff'].web);
    }

    // Nettogewicht (kg als Zahl)
    if (pairs['Nettogewicht']) {
      const dbCol = pairs['Nettogewicht'].db;
      const webCol = pairs['Nettogewicht'].web;
      const kgCell = ws.getRow(r).getCell(webCol);
      // parse web weight string to kg
      let kg = null;
      if (web['Gewicht']) {
        const mKg = String(web['Gewicht']).match(/(-?\d+[.,]?\d*)\s*kg/i);
        const mG  = String(web['Gewicht']).match(/(-?\d+[.,]?\d*)\s*g\b/i);
        if (mKg) kg = Number(mKg[1].replace(',','.'));
        else if (mG) kg = Number(mG[1].replace(',','.'))/1000;
      }
      kgCell.value = kg;
      kgCell.numFmt = '0.00';
      compareNumberAndColor(ws, r, dbCol, webCol, { tolAbs: 0.02, tolRel: 0.01 });
    }

    // Abmessungen -> L/B/H (mm)
    const dims = web['Abmessung'] ? parseDimensionsToLBH(web['Abmessung']) : {L:null,B:null,H:null};
    if (pairs['Länge']) {
      const c = ws.getRow(r).getCell(pairs['Länge'].web);
      c.value = (dims.L != null ? dims.L : null);
      c.numFmt = '0';
      compareIntAndColor(ws, r, pairs['Länge'].db, pairs['Länge'].web, 1);
    }
    if (pairs['Breite']) {
      const c = ws.getRow(r).getCell(pairs['Breite'].web);
      c.value = (dims.B != null ? dims.B : null);
      c.numFmt = '0';
      compareIntAndColor(ws, r, pairs['Breite'].db, pairs['Breite'].web, 1);
    }
    if (pairs['Höhe']) {
      const c = ws.getRow(r).getCell(pairs['Höhe'].web);
      c.value = (dims.H != null ? dims.H : null);
      c.numFmt = '0';
      compareIntAndColor(ws, r, pairs['Höhe'].db, pairs['Höhe'].web, 1);
    }
  }
}

    const out = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="DB_Produktvergleich_verarbeitet.xlsx"');
    res.send(Buffer.from(out));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Server running at http://0.0.0.0:${PORT}`));