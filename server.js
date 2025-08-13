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

// Neue Spaltenkonstanten für DB ↔ Web Paarung
const COLS = {
  // DB-Spalten (bleiben unverändert)
  B: 'B',  // Material
  C: 'C',  // Material-Kurztext (DB)
  E: 'E',  // Herstellername
  F: 'F',  // Hersteller-Artikelnummer (DB)
  N: 'N',  // Fertigung/Prüfhinweis (DB)
  P: 'P',  // Fertigung/Prüfhinweis (DB)
  S: 'S',  // Werkstoff (DB)
  W: 'W',  // Nettogewicht (DB)
  Z: 'Z',  // Länge (DB)
  AB: 'AB', // Breite (DB)
  AD: 'AD', // Höhe (DB)
  AH: 'AH', // Siemens Mobility Materialnummer (A2V)
  
  // Web-Spalten (neu hinzugefügt)
  D: 'D',   // Material-Kurztext (Web)
  G: 'G',   // Hersteller-Artikelnummer (Web)
  Q: 'Q',   // Fertigung/Prüfhinweis (Web)
  T: 'T',   // Werkstoff (Web)
  X: 'X',   // Nettogewicht (Web)
  AA: 'AA', // Länge (Web)
  AC: 'AC', // Breite (Web)
  AE: 'AE'  // Höhe (Web)
};

const HEADER_ROW = 3;
const SUBHEADER_ROW = 4;
const FIRST_DATA_ROW = 5;

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const scraper = new SiemensProductScraper();

function fillColor(ws, addr, color) {
  if (!color) return;
  const map = {
    green: 'FFD5F4E6',
    red:   'FFFDEAEA',
    orange: 'FFFFE6CC'
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
function eqPart(a,b) { 
  const normA = normPartNo(a);
  const normB = normPartNo(b);
  const result = normA === normB;
  console.log(`eqPart: "${a}" -> "${normA}", "${b}" -> "${normB}" -> ${result}`);
  return result;
}
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

    // 1) A2V-Nummern aus Spalte AH ab Zeile 5 einsammeln
    const tasks = [];
    const rowsPerSheet = new Map(); // ws -> [rowIndex,...]
    for (const ws of wb.worksheets) {
      const indices = [];
      const last = ws.lastRow?.number || 0;
      for (let r = FIRST_DATA_ROW; r <= last; r++) {
        const a2v = (ws.getCell(`${COLS.AH}${r}`).value || '').toString().trim().toUpperCase();
        if (a2v.startsWith('A2V')) {
          indices.push(r);
          tasks.push(a2v);
        }
      }
      rowsPerSheet.set(ws, indices);
    }

    // 2) Scrapen
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);

    // 3) Neue Spalten einfügen und Subheader setzen
    for (const ws of wb.worksheets) {
      const prodRows = rowsPerSheet.get(ws) || [];
      if (prodRows.length === 0) continue;

      // 3.1 Neue Web-Spalten einfügen (von rechts nach links, um Spaltenindizes nicht zu verschieben)
      const newColumns = [
        { after: 'C', insert: 'D' },   // Material-Kurztext Web nach C
        { after: 'F', insert: 'G' },   // Hersteller-Artikelnummer Web nach F
        { after: 'P', insert: 'Q' },   // Fertigung/Prüfhinweis Web nach P
        { after: 'S', insert: 'T' },   // Werkstoff Web nach S
        { after: 'W', insert: 'X' },   // Nettogewicht Web nach W
        { after: 'Z', insert: 'AA' },  // Länge Web nach Z
        { after: 'AB', insert: 'AC' }, // Breite Web nach AB
        { after: 'AD', insert: 'AE' }  // Höhe Web nach AD
      ];

      // Spalten von rechts nach links einfügen
      for (let i = newColumns.length - 1; i >= 0; i--) {
        const col = newColumns[i];
        ws.spliceColumns(col.insert, 0, [null]);
      }

      // 3.2 Subheader in Zeile 4 setzen
      ws.getCell('D4').value = 'Web-Wert';
      ws.getCell('G4').value = 'Web-Wert';
      ws.getCell('Q4').value = 'Web-Wert';
      ws.getCell('T4').value = 'Web-Wert';
      ws.getCell('X4').value = 'Web-Wert';
      ws.getCell('AA4').value = 'Web-Wert';
      ws.getCell('AC4').value = 'Web-Wert';
      ws.getCell('AE4').value = 'Web-Wert';

      // 3.3 Web-Daten in die neuen Spalten eintragen
      for (const r of prodRows) {
        const a2v = (ws.getCell(`${COLS.AH}${r}`).value || '').toString().trim().toUpperCase();
        const web = resultsMap.get(a2v) || {};

        // Excel-Werte zum Vergleichen (DB-Werte)
        const exC = ws.getCell(`${COLS.C}${r}`).value;  // Material-Kurztext (DB)
        const exF = ws.getCell(`${COLS.F}${r}`).value;  // Hersteller-Artikelnummer (DB)
        const exP = ws.getCell(`${COLS.P}${r}`).value;  // Fertigung/Prüfhinweis (DB)
        const exS = ws.getCell(`${COLS.S}${r}`).value;  // Werkstoff (DB)
        const exW = ws.getCell(`${COLS.W}${r}`).value;  // Nettogewicht (DB)
        const exZ = ws.getCell(`${COLS.Z}${r}`).value;  // Länge (DB)
        const exAB = ws.getCell(`${COLS.AB}${r}`).value; // Breite (DB)
        const exAD = ws.getCell(`${COLS.AD}${r}`).value; // Höhe (DB)

        // 3.3.1 D – Material-Kurztext (Web)
        if (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') {
          const val = web.Produkttitel;
          ws.getCell(`D${r}`).value = val;
          const isEqual = eqText(exC || '', val);
          fillColor(ws, `D${r}`, isEqual ? 'green' : 'red');
        } else {
          fillColor(ws, `D${r}`, 'orange'); // Orange für fehlende Werte
        }

        // 3.3.2 G – Hersteller-Artikelnummer (Web)
        if (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden') {
          const val = web['Weitere Artikelnummer'];
          ws.getCell(`G${r}`).value = val;
          const excelVal = exF || a2v;
          const isEqual = eqPart(excelVal, val);
          fillColor(ws, `G${r}`, isEqual ? 'green' : 'red');
        } else {
          fillColor(ws, `G${r}`, 'orange');
        }

        // 3.3.3 Q – Fertigung/Prüfhinweis (Web)
        if (web.Materialklassifizierung && web.Materialklassifizierung !== 'Nicht gefunden') {
          const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung));
          if (code) {
            ws.getCell(`Q${r}`).value = code;
            const ok = eqN(exP || '', code);
            fillColor(ws, `Q${r}`, ok ? 'green' : 'red');
          } else {
            fillColor(ws, `Q${r}`, 'orange');
          }
        } else {
          fillColor(ws, `Q${r}`, 'orange');
        }

        // 3.3.4 T – Werkstoff (Web)
        if (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') {
          const val = web.Werkstoff;
          ws.getCell(`T${r}`).value = val;
          const isEqual = eqText(exS || '', val);
          fillColor(ws, `T${r}`, isEqual ? 'green' : 'red');
        } else {
          fillColor(ws, `T${r}`, 'orange');
        }

        // 3.3.5 X – Nettogewicht (Web)
        if (web.Gewicht && web.Gewicht !== 'Nicht gefunden') {
          const { value, unit } = parseWeight(web.Gewicht);
          if (value != null) {
            ws.getCell(`X${r}`).value = value;
            const ok = eqWeight(exW, null, web.Gewicht);
            fillColor(ws, `X${r}`, ok ? 'green' : 'red');
          } else {
            fillColor(ws, `X${r}`, 'orange');
          }
        } else {
          fillColor(ws, `X${r}`, 'orange');
        }

        // 3.3.6 AA – Länge (Web)
        if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
          const d = parseDimensionsToLBH(web.Abmessung);
          if (d.L != null) {
            ws.getCell(`AA${r}`).value = d.L;
            const excelL = toNumber(exZ);
            if (excelL != null) {
              const isEqual = excelL === d.L;
              fillColor(ws, `AA${r}`, isEqual ? 'green' : 'red');
            } else {
              fillColor(ws, `AA${r}`, 'green'); // Nur Web-Wert vorhanden
            }
          } else {
            fillColor(ws, `AA${r}`, 'orange');
          }
        } else {
          fillColor(ws, `AA${r}`, 'orange');
        }

        // 3.3.7 AC – Breite (Web)
        if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
          const d = parseDimensionsToLBH(web.Abmessung);
          if (d.B != null) {
            ws.getCell(`AC${r}`).value = d.B;
            const excelB = toNumber(exAB);
            if (excelB != null) {
              const isEqual = excelB === d.B;
              fillColor(ws, `AC${r}`, isEqual ? 'green' : 'red');
            } else {
              fillColor(ws, `AC${r}`, 'green'); // Nur Web-Wert vorhanden
            }
          } else {
            fillColor(ws, `AC${r}`, 'orange');
          }
        } else {
          fillColor(ws, `AC${r}`, 'orange');
        }

        // 3.3.8 AE – Höhe (Web)
        if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
          const d = parseDimensionsToLBH(web.Abmessung);
          if (d.H != null) {
            ws.getCell(`AE${r}`).value = d.H;
            const excelH = toNumber(exAD);
            if (excelH != null) {
              const isEqual = excelH === d.H;
              fillColor(ws, `AE${r}`, isEqual ? 'green' : 'red');
            } else {
              fillColor(ws, `AE${r}`, 'green'); // Nur Web-Wert vorhanden
            }
          } else {
            fillColor(ws, `AE${r}`, 'orange');
          }
        } else {
          fillColor(ws, `AE${r}`, 'orange');
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