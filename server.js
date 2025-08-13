// server.js — kompletter Rewrite
// Ziel: DB/Web nebeneinander (keine zusätzliche Zeile), Zeile 3 = Header, Zeile 4 = Subheader (DB-Wert/Web-Wert), Daten ab Zeile 5
// Farblogik: gleich=grün, ungleich=rot, fehlt/nicht gefunden=orange

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

// ===== Excel Layout Konstanten (gemäß Vorlage)
const HEADER_ROW = 3;       // Spaltennamen (z.B. "Materialkurztext")
const SUBHEADER_ROW = 4;    // Unterkopf: "DB-Wert" / "Web-Wert"
const FIRST_DATA_ROW = 5;   // Daten beginnen in Zeile 5

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// =============================
// Excel-Helfer
// =============================
function colNumberToLetter(num) {
  let s = ''; let n = num;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
const addr = (col, row) => `${colNumberToLetter(col)}${row}`;

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

// Vergleiche
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

// Spalte anhand von Begriffen in Header-Zeile 3 finden
function findColumnByIncludes(ws, headerRow, ...needles) {
  const row = ws.getRow(headerRow);
  const lastCol = Math.max(ws.actualColumnCount || 0, row.actualCellCount || 0, row.cellCount || 0, 60);
  const nls = needles.map(n => n.toLowerCase());
  for (let c = 1; c <= lastCol; c++) {
    const name = getCellValueAsString(row.getCell(c)).toLowerCase();
    if (!name) continue;
    const ok = nls.every(n => name.includes(n));
    if (ok) return c;
  }
  return null;
}

// DB/Web-Subheader in Zeile 4 setzen
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
// HTTP-Routen
// =============================
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/api/health', (req, res) => res.json({ ok: true, time: new Date().toISOString() }));

app.post('/api/process-excel', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    const scraper = new SiemensProductScraper();

    for (const ws of wb.worksheets) {
      // 1) Spalten in Zeile 3 finden
      const col = {};
      col.material      = findColumnByIncludes(ws, HEADER_ROW, 'material'); // B (nur Anzeige, wird nicht verglichen)
      col.materialkurz  = findColumnByIncludes(ws, HEADER_ROW, 'material', 'kurz'); // C
      col.herstArtNr    = findColumnByIncludes(ws, HEADER_ROW, 'artikel');         // F ("Her.-Artikelnummer")
      col.fertPruef     = findColumnByIncludes(ws, HEADER_ROW, 'fert', 'prüf') || findColumnByIncludes(ws, HEADER_ROW, 'fert', 'pruef'); // P
      col.werkstoff     = findColumnByIncludes(ws, HEADER_ROW, 'werkstoff');       // S
      col.nettogew      = findColumnByIncludes(ws, HEADER_ROW, 'netto');           // W
      col.gewEinheit    = findColumnByIncludes(ws, HEADER_ROW, 'gewichtseinheit') || findColumnByIncludes(ws, HEADER_ROW, 'gewicht', 'einheit'); // Y
      col.laenge        = findColumnByIncludes(ws, HEADER_ROW, 'länge') || findColumnByIncludes(ws, HEADER_ROW, 'laenge'); // Z
      col.breite        = findColumnByIncludes(ws, HEADER_ROW, 'breite');          // AB
      col.hoehe         = findColumnByIncludes(ws, HEADER_ROW, 'höhe') || findColumnByIncludes(ws, HEADER_ROW, 'hoehe'); // AD
      col.dimEinheit    = findColumnByIncludes(ws, HEADER_ROW, 'einheit', 'abmaße') || findColumnByIncludes(ws, HEADER_ROW, 'einheit', 'abmasse'); // AF

      // A2V (Siemens Mobility Materialnummer)
      col.a2v = findColumnByIncludes(ws, HEADER_ROW, 'a2v')
             || findColumnByIncludes(ws, HEADER_ROW, 'siemens', 'materialnummer')
             || findColumnByIncludes(ws, HEADER_ROW, 'materialnummer')
             || 34; // Fallback AH

      // 2) Nachbarspalten (Web) definieren und Subheader setzen
      const pairs = [];
      const addPair = (dbCol) => { if (dbCol) pairs.push({ dbCol, webCol: dbCol + 1 }); };
      addPair(col.materialkurz);
      addPair(col.herstArtNr);
      addPair(col.fertPruef);
      addPair(col.werkstoff);
      addPair(col.nettogew);
      addPair(col.laenge);
      addPair(col.breite);
      addPair(col.hoehe);
      ensureSubheaders(ws, pairs);

      // 3) A2V-Liste sammeln (nur Zeilen mit A2V)
      const last = ws.lastRow ? ws.lastRow.number : Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
      const rows = [];
      const ids = [];
      for (let r = FIRST_DATA_ROW; r <= last; r++) {
        const v = getCellValueAsString(ws.getRow(r).getCell(col.a2v)).trim().toUpperCase();
        if (v && v.startsWith('A2V')) { rows.push(r); ids.push(v); }
      }
      if (!ids.length) continue; // nichts zu tun auf diesem Sheet

      // 4) Scrapen
      const results = await scraper.scrapeMany(ids, SCRAPE_CONCURRENCY); // Map<A2V, {...}>

      // 5) Schreiben & Färben je Zeile (kein Zeileneinfügen)
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const a2v = ids[i];
        const web = results.get(a2v) || {};

        // DB-Werte lesen
        const db = {
          materialkurz: ws.getRow(r).getCell(col.materialkurz || 0)?.value,
          herstArtNr:   ws.getRow(r).getCell(col.herstArtNr   || 0)?.value,
          fertPruef:    ws.getRow(r).getCell(col.fertPruef    || 0)?.value,
          werkstoff:    ws.getRow(r).getCell(col.werkstoff    || 0)?.value,
          nettogew:     ws.getRow(r).getCell(col.nettogew     || 0)?.value,
          gewEinheit:   col.gewEinheit ? ws.getRow(r).getCell(col.gewEinheit)?.value : '',
          laenge:       ws.getRow(r).getCell(col.laenge       || 0)?.value,
          breite:       ws.getRow(r).getCell(col.breite       || 0)?.value,
          hoehe:        ws.getRow(r).getCell(col.hoehe        || 0)?.value
        };

        // 5.1 Materialkurztext (Produkttitel)
        if (col.materialkurz) {
          const c = col.materialkurz + 1;
          const webVal = (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') ? web.Produkttitel : '';
          if (webVal) ws.getRow(r).getCell(c).value = webVal;
          fillColor(ws, c, r, webVal ? (eqText(db.materialkurz || '', webVal) ? 'green' : 'red') : 'orange');
        }

        // 5.2 Hersteller-Artikelnummer (Weitere Artikelnummer; Fallback A2V)
        if (col.herstArtNr) {
          const c = col.herstArtNr + 1;
          let webVal = (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden') ? web['Weitere Artikelnummer'] : '';
          if (!webVal && a2v) webVal = a2v; // Fallback
          if (webVal) ws.getRow(r).getCell(c).value = webVal;
          const excelVal = db.herstArtNr || a2v;
          fillColor(ws, c, r, webVal ? (eqPart(excelVal, webVal) ? 'green' : 'red') : 'orange');
        }

        // 5.3 Fert./Prüfhinweis (aus Materialklassifizierung gemappt)
        if (col.fertPruef) {
          const c = col.fertPruef + 1;
          const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung || ''));
          if (code) ws.getRow(r).getCell(c).value = code;
          fillColor(ws, c, r, code ? (eqN(db.fertPruef || '', code) ? 'green' : 'red') : 'orange');
        }

        // 5.4 Werkstoff
        if (col.werkstoff) {
          const c = col.werkstoff + 1;
          const webVal = (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') ? web.Werkstoff : '';
          if (webVal) ws.getRow(r).getCell(c).value = webVal;
          fillColor(ws, c, r, webVal ? (eqText(db.werkstoff || '', webVal) ? 'green' : 'red') : 'orange');
        }

        // 5.5 Nettogewicht (mit Einheit)
        if (col.nettogew) {
          const c = col.nettogew + 1;
          const webVal = (web.Gewicht && web.Gewicht !== 'Nicht gefunden') ? web.Gewicht : '';
          if (webVal) {
            const { value, unit } = parseWeight(webVal);
            ws.getRow(r).getCell(c).value = (value != null) ? value : webVal;
            if (unit && col.gewEinheit) {
              const uCell = ws.getRow(r).getCell(col.gewEinheit);
              if (!uCell.value) uCell.value = unit;
            }
            fillColor(ws, c, r, eqWeight(db.nettogew, db.gewEinheit, webVal) ? 'green' : 'red');
          } else {
            fillColor(ws, c, r, 'orange');
          }
        }

        // 5.6 Abmessungen L/B/H (aus "Abmessung")
        const dims = (web.Abmessung && web.Abmessung !== 'Nicht gefunden') ? parseDimensionsToLBH(web.Abmessung) : { L:null, B:null, H:null };
        if (col.laenge) {
          const c = col.laenge + 1;
          if (dims.L != null) ws.getRow(r).getCell(c).value = dims.L;
          if      (toNumber(db.laenge) != null && dims.L != null) fillColor(ws, c, r, toNumber(db.laenge) === dims.L ? 'green' : 'red');
          else if (dims.L == null) fillColor(ws, c, r, 'orange');
        }
        if (col.breite) {
          const c = col.breite + 1;
          if (dims.B != null) ws.getRow(r).getCell(c).value = dims.B;
          if      (toNumber(db.breite) != null && dims.B != null) fillColor(ws, c, r, toNumber(db.breite) === dims.B ? 'green' : 'red');
          else if (dims.B == null) fillColor(ws, c, r, 'orange');
        }
        if (col.hoehe) {
          const c = col.hoehe + 1;
          if (dims.H != null) ws.getRow(r).getCell(c).value = dims.H;
          if      (toNumber(db.hoehe) != null && dims.H != null) fillColor(ws, c, r, toNumber(db.hoehe) === dims.H ? 'green' : 'red');
          else if (dims.H == null) fillColor(ws, c, r, 'orange');
        }

        // Optional: Einheit Abmaße setzen, wenn leer
        if (col.dimEinheit) {
          const u = ws.getRow(r).getCell(col.dimEinheit);
          if (!u.value) u.value = 'MM';
        }
      }
    }

    // 6) Download
    const out = await (wb.xlsx.writeBuffer ? wb.xlsx.writeBuffer() : wb.xlsx.writeBuffer());
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="DB_Produktvergleich_verarbeitet.xlsx"');
    res.send(Buffer.from(out));
  } catch (err) {
    console.error('PROCESS ERROR:', err && err.stack || err);
    res.status(500).json({ error: String(err && err.message || err) });
  }
});

app.listen(PORT, () => console.log(`Server running at http://0.0.0.0:${PORT}`));
