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
const { SiemensProductScraper } = require('./scraper');

const app = express();
const PORT = process.env.PORT || 3000;
const SCRAPE_CONCURRENCY = Number(process.env.SCRAPE_CONCURRENCY || 4);

// Layout-Konstanten der Vorlage
const HEADER_ROW = 3;       // Spaltennamen
const SUBHEADER_ROW = 4;    // "DB-Wert" / "Web-Wert"
const FIRST_DATA_ROW = 5;   // erste Datenzeile

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

const scraper = new SiemensProductScraper();

// =============================
// Hilfsfunktionen Excel
// =============================
function colNumberToLetter(num) {
  let s = ''; let n = num;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function addr(col, row) { return `${colNumberToLetter(col)}${row}`; }

function getCellValueAsString(cell) {
  const v = cell?.value;
  if (v == null) return '';
  if (typeof v === 'object' && v.text != null) return String(v.text);
  return String(v);
}

function fillColor(ws, col, row, color) {
  if (!color) return;
  const map = { green: 'FFD5F4E6', red: 'FFFDEAEA', orange: 'FFFFF3CD' };
  ws.getCell(addr(col, row)).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: map[color] || map.green } };
}

function eqText(a, b) {
  if (a == null || b == null) return false;
  const A = String(a).trim().toLowerCase().replace(/\s+/g, ' ');
  const B = String(b).trim().toLowerCase().replace(/\s+/g, ' ');
  return A === B;
}

function eqPart(a, b) { return normPartNo(a) === normPartNo(b); }
function eqN(a, b)     { return normalizeNCode(a) === normalizeNCode(b); }

function eqWeight(exVal, exUnit, webVal) {
  const { value: wv, unit: wu } = parseWeight(webVal);
  if (wv == null) return false;
  const exNum = toNumber(exVal);
  const exU   = (exUnit || '').toString().trim().toLowerCase();
  if (exNum == null) return false;
  const a = weightToKg(exNum, exU);
  const b = weightToKg(wv, wu || exU || 'kg');
  if (a == null || b == null) return false;
  return Math.abs(a - b) < 1e-9;
}

function findColumnByIncludes(ws, headerRow, ...needles) {
  const row = ws.getRow(headerRow);
  const lastCol = Math.max(ws.actualColumnCount || 0, row.actualCellCount || 0, row.cellCount || 0, 50);
  const nls = needles.map(n => n.toLowerCase());
  for (let c = 1; c <= lastCol; c++) {
    const name = getCellValueAsString(row.getCell(c)).toLowerCase();
    if (!name) continue;
    const ok = nls.every(n => name.includes(n));
    if (ok) return c;
  }
  return null;
}

function ensureSubheaders(ws, pairs) {
  for (const { dbCol, webCol } of pairs) {
    if (!dbCol || !webCol) continue;
    const dbCell  = ws.getRow(SUBHEADER_ROW).getCell(dbCol);
    const webCell = ws.getRow(SUBHEADER_ROW).getCell(webCol);
    if (!dbCell.value)  dbCell.value  = 'DB-Wert';
    if (!webCell.value) webCell.value = 'Web-Wert';
  }
}

// =============================
// Routes
// =============================
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/api/health', (req, res) => res.json({ ok: true, time: new Date().toISOString() }));

app.post('/api/process-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    for (const ws of wb.worksheets) {
      const col = {};
      col.materialkurz = findColumnByIncludes(ws, HEADER_ROW, 'material', 'kurz');
      col.herstArtNr   = findColumnByIncludes(ws, HEADER_ROW, 'hersteller', 'artikel');
      col.fertPruef    = findColumnByIncludes(ws, HEADER_ROW, 'fertigung', 'prüf') || findColumnByIncludes(ws, HEADER_ROW, 'fertigung', 'pruef');
      col.werkstoff    = findColumnByIncludes(ws, HEADER_ROW, 'werkstoff');
      col.nettogew     = findColumnByIncludes(ws, HEADER_ROW, 'netto');
      col.nGewEinheit  = findColumnByIncludes(ws, HEADER_ROW, 'gewicht', 'einheit') || null;
      col.laenge       = findColumnByIncludes(ws, HEADER_ROW, 'länge') || findColumnByIncludes(ws, HEADER_ROW, 'laenge');
      col.breite       = findColumnByIncludes(ws, HEADER_ROW, 'breite');
      col.hoehe        = findColumnByIncludes(ws, HEADER_ROW, 'höhe') || findColumnByIncludes(ws, HEADER_ROW, 'hoehe');

      col.a2v = findColumnByIncludes(ws, HEADER_ROW, 'a2v')
             || findColumnByIncludes(ws, HEADER_ROW, 'siemens', 'materialnummer')
             || findColumnByIncludes(ws, HEADER_ROW, 'materialnummer')
             || 34; // Fallback AH

      const pairs = [];
      function addPair(dbCol) { if (dbCol) pairs.push({ dbCol, webCol: dbCol + 1 }); }
      addPair(col.materialkurz);
      addPair(col.herstArtNr);
      addPair(col.fertPruef);
      addPair(col.werkstoff);
      addPair(col.nettogew);
      addPair(col.laenge);
      addPair(col.breite);
      addPair(col.hoehe);

      ensureSubheaders(ws, pairs);

      const last = ws.lastRow ? ws.lastRow.number : ws.rowCount || ws.actualRowCount || 0;
      const rowsToProcess = [];
      const lookups = [];
      for (let r = FIRST_DATA_ROW; r <= last; r++) {
        const a2vVal = getCellValueAsString(ws.getRow(r).getCell(col.a2v)).trim().toUpperCase();
        if (a2vVal && a2vVal.startsWith('A2V')) { rowsToProcess.push(r); lookups.push(a2vVal); }
      }

      const resultsMap = await scraper.scrapeMany(lookups, SCRAPE_CONCURRENCY);

      for (const r of rowsToProcess) {
        const a2v = getCellValueAsString(ws.getRow(r).getCell(col.a2v)).trim().toUpperCase();
        const web = resultsMap.get(a2v) || {};

        const db = {
          materialkurz: ws.getRow(r).getCell(col.materialkurz)?.value,
          herstArtNr:   ws.getRow(r).getCell(col.herstArtNr)?.value,
          fertPruef:    ws.getRow(r).getCell(col.fertPruef)?.value,
          werkstoff:    ws.getRow(r).getCell(col.werkstoff)?.value,
          nettogew:     ws.getRow(r).getCell(col.nettogew)?.value,
          nGewEinheit:  col.nGewEinheit ? ws.getRow(r).getCell(col.nGewEinheit)?.value : '',
          laenge:       ws.getRow(r).getCell(col.laenge)?.value,
          breite:       ws.getRow(r).getCell(col.breite)?.value,
          hoehe:        ws.getRow(r).getCell(col.hoehe)?.value
        };

        // Materialkurztext
        if (col.materialkurz) {
          const webVal = (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') ? web.Produkttitel : '';
          const c = col.materialkurz + 1;
          if (webVal) ws.getRow(r).getCell(c).value = webVal;
          if (!webVal) fillColor(ws, c, r, 'orange');
          else fillColor(ws, c, r, eqText(db.materialkurz || '', webVal) ? 'green' : 'red');
        }

        // Hersteller-Artikelnummer
        if (col.herstArtNr) {
          let webVal = (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden') ? web['Weitere Artikelnummer'] : '';
          if (!webVal && a2v) webVal = a2v;
          const c = col.herstArtNr + 1;
          if (webVal) ws.getRow(r).getCell(c).value = webVal;
          if (!webVal) fillColor(ws, c, r, 'orange');
          else {
            const excelVal = db.herstArtNr || a2v;
            fillColor(ws, c, r, eqPart(excelVal, webVal) ? 'green' : 'red');
          }
        }

        // Fertigung/Prüf-Nachweis
        if (col.fertPruef) {
          const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung || ''));
          const c = col.fertPruef + 1;
          if (code) ws.getRow(r).getCell(c).value = code;
          if (!code) fillColor(ws, c, r, 'orange');
          else fillColor(ws, c, r, eqN(db.fertPruef || '', code) ? 'green' : 'red');
        }

        // Werkstoff
        if (col.werkstoff) {
          const webVal = (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') ? web.Werkstoff : '';
          const c = col.werkstoff + 1;
          if (webVal) ws.getRow(r).getCell(c).value = webVal;
          if (!webVal) fillColor(ws, c, r, 'orange');
          else fillColor(ws, c, r, eqText(db.werkstoff || '', webVal) ? 'green' : 'red');
        }

        // Nettogewicht
        if (col.nettogew) {
          const c = col.nettogew + 1;
          const webVal = (web.Gewicht && web.Gewicht !== 'Nicht gefunden') ? web.Gewicht : '';
          if (webVal) {
            const { value, unit } = parseWeight(webVal);
            if (value != null) ws.getRow(r).getCell(c).value = value; else ws.getRow(r).getCell(c).value = webVal;
            if (unit && col.nGewEinheit) {
              const uCell = ws.getRow(r).getCell(col.nGewEinheit);
              if (!uCell.value) uCell.value = unit;
            }
            const ok = eqWeight(db.nettogew, db.nGewEinheit, webVal);
            fillColor(ws, c, r, ok ? 'green' : 'red');
          } else {
            fillColor(ws, c, r, 'orange');
          }
        }

        // Abmessungen
        const dims = (web.Abmessung && web.Abmessung !== 'Nicht gefunden') ? parseDimensionsToLBH(web.Abmessung) : { L:null, B:null, H:null };
        if (col.laenge) {
          const c = col.laenge + 1;
          if (dims.L != null) ws.getRow(r).getCell(c).value = dims.L;
          if (toNumber(db.laenge) != null && dims.L != null) fillColor(ws, c, r, toNumber(db.laenge) === dims.L ? 'green' : 'red');
          else if (dims.L == null) fillColor(ws, c, r, 'orange');
        }
        if (col.breite) {
          const c = col.breite + 1;
          if (dims.B != null) ws.getRow(r).getCell(c).value = dims.B;
          if (toNumber(db.breite) != null && dims.B != null) fillColor(ws, c, r, toNumber(db.breite) === dims.B ? 'green' : 'red');
          else if (dims.B == null) fillColor(ws, c, r, 'orange');
        }
        if (col.hoehe) {
          const c = col.hoehe + 1;
          if (dims.H != null) ws.getRow(r).getCell(c).value = dims.H;
          if (toNumber(db.hoehe) != null && dims.H != null) fillColor(ws, c, r, toNumber(db.hoehe) === dims.H ? 'green' : 'red');
          else if (dims.H == null) fillColor(ws, c, r, 'orange');
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
