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
  const map = {
    green:  'FFD5F4E6', // hellgrün
    red:    'FFFDEAEA', // hellrot
    orange: 'FFFFF3CD'  // hellorange
  };
  const a = addr(col, row);
  ws.getCell(a).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: map[color] || map.green } };
}

// Textgleichheit (robust, streng nach Normalisierung)
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
  return Math.abs(a - b) < 1e-9; // strikt
}

// =========
// Header-Lookup: wir finden die Spalten dynamisch anhand der Vorlagen-Header (Zeile 3)
// =========
function findColumnByIncludes(ws, headerRow, ...needles) {
  const lastCol = ws.columnCount || (ws.getRow(headerRow).cellCount);
  const nls = needles.map(n => n.toLowerCase());
  for (let c = 1; c <= lastCol; c++) {
    const name = getCellValueAsString(ws.getCell(headerRow, c)).toLowerCase();
    if (!name) continue;
    const ok = nls.every(n => name.includes(n));
    if (ok) return c;
  }
  return null;
}

function ensureSubheaders(ws, pairs) {
  // Nur die Paare markieren: links DB-Wert, rechts Web-Wert; Rest leer lassen
  for (const { dbCol, webCol } of pairs) {
    if (!dbCol || !webCol) continue;
    const dbAddr  = addr(dbCol, SUBHEADER_ROW);
    const webAddr = addr(webCol, SUBHEADER_ROW);
    if (!ws.getCell(dbAddr).value)  ws.getCell(dbAddr).value  = 'DB-Wert';
    if (!ws.getCell(webAddr).value) ws.getCell(webAddr).value = 'Web-Wert';
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

    // Wir verarbeiten alle Tabellenblätter
    for (const ws of wb.worksheets) {
      // 1) Spalten dynamisch anhand Kopfzeile (Zeile 3) finden
      const col = {};
      col.materialkurz = findColumnByIncludes(ws, HEADER_ROW, 'material', 'kurz');        // C erwartet
      col.herstArtNr   = findColumnByIncludes(ws, HEADER_ROW, 'hersteller', 'artikel');   // E erwartet
      col.fertPruef    = findColumnByIncludes(ws, HEADER_ROW, 'fertigung', 'prüf');       // N erwartet
      if (!col.fertPruef) col.fertPruef = findColumnByIncludes(ws, HEADER_ROW, 'fertigung', 'pruef');
      col.werkstoff    = findColumnByIncludes(ws, HEADER_ROW, 'werkstoff');               // P erwartet
      col.nettogew     = findColumnByIncludes(ws, HEADER_ROW, 'netto');                   // S erwartet
      col.nGewEinheit  = findColumnByIncludes(ws, HEADER_ROW, 'gewicht', 'einheit') || null; // optional
      col.laenge       = findColumnByIncludes(ws, HEADER_ROW, 'länge') || findColumnByIncludes(ws, HEADER_ROW, 'laenge'); // U erwartet
      col.breite       = findColumnByIncludes(ws, HEADER_ROW, 'breite');                  // V erwartet
      col.hoehe        = findColumnByIncludes(ws, HEADER_ROW, 'höhe') || findColumnByIncludes(ws, HEADER_ROW, 'hoehe');   // W erwartet

      // A2V-Spalte (in der Vorlage meist "Siemens Mobility Materialnummer (A2V)")
      col.a2v = findColumnByIncludes(ws, HEADER_ROW, 'a2v')
             || findColumnByIncludes(ws, HEADER_ROW, 'siemens', 'materialnummer')
             || findColumnByIncludes(ws, HEADER_ROW, 'materialnummer')
             || null;
      // Fallback: bekannte Positionen (AH, Z)
      if (!col.a2v) {
        // AH ~ 34
        col.a2v = 34; // AH
      }

      // Für jede der relevanten DB-Spalten die direkte Web-Nachbarspalte
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

      // 2) Subheader "DB-Wert" / "Web-Wert" in Zeile 4 setzen (nur für obige Paare)
      ensureSubheaders(ws, pairs);

      // 3) A2V-Werte einsammeln (ab Zeile 5)
      const last = ws.lastRow?.number || 0;
      const rowsToProcess = [];
      const lookups = [];
      for (let r = FIRST_DATA_ROW; r <= last; r++) {
        const a2vCell = ws.getCell(r, col.a2v);
        const a2vVal = getCellValueAsString(a2vCell).trim().toUpperCase();
        if (a2vVal.startsWith('A2V')) {
          rowsToProcess.push(r);
          lookups.push(a2vVal);
        }
      }

      // 4) Scrapen (gebündelt)
      const resultsMap = await scraper.scrapeMany(lookups, SCRAPE_CONCURRENCY);

      // 5) Pro Zeile Web-Daten in die Nachbarspalte schreiben (kein Zeileneinfügen!)
      for (const r of rowsToProcess) {
        const a2v = getCellValueAsString(ws.getCell(r, col.a2v)).trim().toUpperCase();
        const web = resultsMap.get(a2v) || {};

        // Excel-DB-Werte einsammeln
        const db = {
          materialkurz: ws.getCell(r, col.materialkurz)?.value,
          herstArtNr:   ws.getCell(r, col.herstArtNr)?.value,
          fertPruef:    ws.getCell(r, col.fertPruef)?.value,
          werkstoff:    ws.getCell(r, col.werkstoff)?.value,
          nettogew:     ws.getCell(r, col.nettogew)?.value,
          nGewEinheit:  col.nGewEinheit ? ws.getCell(r, col.nGewEinheit)?.value : '',
          laenge:       ws.getCell(r, col.laenge)?.value,
          breite:       ws.getCell(r, col.breite)?.value,
          hoehe:        ws.getCell(r, col.hoehe)?.value
        };

        // 5.1 Materialkurztext (Produkttitel als Proxy)
        if (col.materialkurz) {
          const webVal = (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') ? web.Produkttitel : '';
          const c = col.materialkurz + 1;
          if (webVal) ws.getCell(r, c).value = webVal;
          if (!webVal) fillColor(ws, c, r, 'orange');
          else fillColor(ws, c, r, eqText(db.materialkurz || '', webVal) ? 'green' : 'red');
        }

        // 5.2 Hersteller-Artikelnummer ("Weitere Artikelnummer" oder A2V als Fallback)
        if (col.herstArtNr) {
          let webVal = (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden') ? web['Weitere Artikelnummer'] : '';
          if (!webVal && a2v) webVal = a2v;
          const c = col.herstArtNr + 1;
          if (webVal) ws.getCell(r, c).value = webVal;
          if (!webVal) fillColor(ws, c, r, 'orange');
          else {
            const excelVal = db.herstArtNr || a2v;
            fillColor(ws, c, r, eqPart(excelVal, webVal) ? 'green' : 'red');
          }
        }

        // 5.3 Fertigung/Prüf-Nachweis (aus Materialklassifizierung gemappt)
        if (col.fertPruef) {
          const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung || ''));
          const c = col.fertPruef + 1;
          if (code) ws.getCell(r, c).value = code;
          if (!code) fillColor(ws, c, r, 'orange');
          else fillColor(ws, c, r, eqN(db.fertPruef || '', code) ? 'green' : 'red');
        }

        // 5.4 Werkstoff
        if (col.werkstoff) {
          const webVal = (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') ? web.Werkstoff : '';
          const c = col.werkstoff + 1;
          if (webVal) ws.getCell(r, c).value = webVal;
          if (!webVal) fillColor(ws, c, r, 'orange');
          else fillColor(ws, c, r, eqText(db.werkstoff || '', webVal) ? 'green' : 'red');
        }

        // 5.5 Nettogewicht (Wert + Einheit im Vergleich; Web-Wert in die Nachbarspalte schreiben)
        if (col.nettogew) {
          const c = col.nettogew + 1;
          const webVal = (web.Gewicht && web.Gewicht !== 'Nicht gefunden') ? web.Gewicht : '';
          if (webVal) {
            const { value, unit } = parseWeight(webVal);
            if (value != null) ws.getCell(r, c).value = value; else ws.getCell(r, c).value = webVal; // falls keine Zahl parsebar
            // Einheit in DB-Einheitsspalte belassen/ergänzen, falls vorhanden
            if (unit && col.nGewEinheit && !ws.getCell(r, col.nGewEinheit).value) ws.getCell(r, col.nGewEinheit).value = unit;
            const ok = eqWeight(db.nettogew, db.nGewEinheit, webVal);
            fillColor(ws, c, r, ok ? 'green' : 'red');
          } else {
            fillColor(ws, c, r, 'orange');
          }
        }

        // 5.6 Abmessungen L/B/H (Web → Nachbarspalte, Vergleich gegen DB)
        const dims = (web.Abmessung && web.Abmessung !== 'Nicht gefunden') ? parseDimensionsToLBH(web.Abmessung) : { L:null, B:null, H:null };
        // Länge
        if (col.laenge) {
          const c = col.laenge + 1;
          if (dims.L != null) ws.getCell(r, c).value = dims.L;
          if (toNumber(db.laenge) != null && dims.L != null) fillColor(ws, c, r, toNumber(db.laenge) === dims.L ? 'green' : 'red');
          else if (dims.L == null) fillColor(ws, c, r, 'orange');
        }
        // Breite
        if (col.breite) {
          const c = col.breite + 1;
          if (dims.B != null) ws.getCell(r, c).value = dims.B;
          if (toNumber(db.breite) != null && dims.B != null) fillColor(ws, c, r, toNumber(db.breite) === dims.B ? 'green' : 'red');
          else if (dims.B == null) fillColor(ws, c, r, 'orange');
        }
        // Höhe
        if (col.hoehe) {
          const c = col.hoehe + 1;
          if (dims.H != null) ws.getCell(r, c).value = dims.H;
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
