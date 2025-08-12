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
const SCRAPE_CONCURRENCY = Number(process.env.SCRAPE_CONCURRENCY || 4);
const WEIGHT_TOL_PCT = Number(process.env.WEIGHT_TOL_PCT || 0); // 0 = strikt
const COLS = { Z:'Z', E:'E', C:'C', S:'S', T:'T', U:'U', V:'V', W:'W', P:'P', N:'N' };
const HEADER_ROW = 3;
const FIRST_DATA_ROW = 4;

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const scraper = new SiemensProductScraper();

function fillColor(ws, addr, color) {
  if (!color) return;
  const map = {
    green: 'FFD5F4E6',
    red:   'FFFDEAEA'
  };
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
    for (const ws of wb.worksheets) {
      const indices = [];
      const last = ws.lastRow?.number || 0;
      for (let r = FIRST_DATA_ROW; r <= last; r++) {
        const a2v = (ws.getCell(`${COLS.Z}${r}`).value || '').toString().trim().toUpperCase();
        if (a2v.startsWith('A2V')) {
          indices.push(r);
          tasks.push(a2v);
        }
      }
      rowsPerSheet.set(ws, indices);
    }

    // 2) Scrapen
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);

    // 3) Nach JEDEM Produkt genau EINE Web-Daten-Zeile einfügen
    for (const ws of wb.worksheets) {
      const prodRows = rowsPerSheet.get(ws) || [];
      for (let i = prodRows.length - 1; i >= 0; i--) {
        const r = prodRows[i];
        const a2v = (ws.getCell(`${COLS.Z}${r}`).value || '').toString().trim().toUpperCase();
        const web = resultsMap.get(a2v) || {};

        // Einfügen einer Zeile
        const insertAt = r + 1;
        ws.spliceRows(insertAt, 0, [null]);
        const webRow = insertAt;

        // Excel-Werte zum Vergleichen
        const exE = ws.getCell(`${COLS.E}${r}`).value;
        const exC = ws.getCell(`${COLS.C}${r}`).value;
        const exS = ws.getCell(`${COLS.S}${r}`).value;
        const exT = ws.getCell(`${COLS.T}${r}`).value;
        const exU = ws.getCell(`${COLS.U}${r}`).value;
        const exV = ws.getCell(`${COLS.V}${r}`).value;
        const exW = ws.getCell(`${COLS.W}${r}`).value;
        const exP = ws.getCell(`${COLS.P}${r}`).value;
        const exN = ws.getCell(`${COLS.N}${r}`).value;

        // 3.1 Z (immer A2V)
        ws.getCell(`${COLS.Z}${webRow}`).value = a2v;

        // 3.2 E – Weitere Artikelnummer (nur setzen, wenn gefunden)
        if (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden') {
          const val = web['Weitere Artikelnummer'];
          ws.getCell(`${COLS.E}${webRow}`).value = val;
          fillColor(ws, `${COLS.E}${webRow}`, eqPart(exE || '', val) ? 'green' : 'red');
        }

        // 3.3 C – Produkttitel
        if (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') {
          const val = web.Produkttitel;
          ws.getCell(`${COLS.C}${webRow}`).value = val;
          fillColor(ws, `${COLS.C}${webRow}`, eqText(exC || '', val) ? 'green' : 'red');
        }

        // 3.4 S/T – Gewicht (Wert + Einheit)
        if (web.Gewicht && web.Gewicht !== 'Nicht gefunden') {
          const { value, unit } = parseWeight(web.Gewicht);
          if (value != null) {
            ws.getCell(`${COLS.S}${webRow}`).value = value;
            if (unit) ws.getCell(`${COLS.T}${webRow}`).value = unit;
            const ok = eqWeight(exS, exT, web.Gewicht);
            fillColor(ws, `${COLS.S}${webRow}`, ok ? 'green' : 'red');
            if (unit) fillColor(ws, `${COLS.T}${webRow}`, ok ? 'green' : 'red');
          }
        }

        // 3.5 U/V/W – Abmessungen
        if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
          const d = parseDimensionsToLBH(web.Abmessung);
          if (d.L != null) ws.getCell(`${COLS.U}${webRow}`).value = d.L;
          if (d.B != null) ws.getCell(`${COLS.V}${webRow}`).value = d.B;
          if (d.H != null) ws.getCell(`${COLS.W}${webRow}`).value = d.H;
          const ok = eqDim(exU, exV, exW, web.Abmessung);
          if (d.L != null) fillColor(ws, `${COLS.U}${webRow}`, ok ? 'green' : 'red');
          if (d.B != null) fillColor(ws, `${COLS.V}${webRow}`, ok ? 'green' : 'red');
          if (d.H != null) fillColor(ws, `${COLS.W}${webRow}`, ok ? 'green' : 'red');
        }

        // 3.6 P – Werkstoff (nur setzen, wenn gefunden)
        if (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') {
          const val = web.Werkstoff;
          ws.getCell(`${COLS.P}${webRow}`).value = val;
          fillColor(ws, `${COLS.P}${webRow}`, eqText(exP || '', val) ? 'green' : 'red');
        }

        // 3.7 N – Klassifizierung → Excel-Code
        if (web.Materialklassifizierung && web.Materialklassifizierung !== 'Nicht gefunden') {
          const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung));
          if (code) {
            ws.getCell(`${COLS.N}${webRow}`).value = code;
            const ok = eqN(exN || '', code);
            fillColor(ws, `${COLS.N}${webRow}`, ok ? 'green' : 'red');
          }
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