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
    red: 'FFFDEAEA',
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

// Neue Funktion: Erstelle das gewünschte Layout
function createComparisonLayout(ws, headerRow, firstDataRow) {
  // Zeile 4: Neue Zeile mit "DB-Wert" und "Web-Wert" nur unter den relevanten Spalten
  // Diese Spalten bekommen eine neue Zeile 4 mit DB-Wert/Web-Wert:
  const comparisonColumns = {
    'C': 'D', // Materialkurztext: C=DB, D=Web
    'E': 'F', // Her.-Artikelnummer: E=DB, F=Web  
    'G': 'H', // Fert./Prüfhinweis: G=DB, H=Web
    'I': 'J', // Werkstoff: I=DB, J=Web
    'K': 'L', // Nettogewicht: K=DB, L=Web
    'M': 'N', // Länge: M=DB, N=Web
    'O': 'P', // Breite: O=DB, P=Web
    'Q': 'R'  // Höhe: Q=DB, R=Web
  };

  // Neue Spalten für Web-Werte einfügen (nach den bestehenden DB-Spalten)
  Object.entries(comparisonColumns).forEach(([dbCol, webCol]) => {
    // Web-Spalte einfügen (nach der DB-Spalte)
    ws.getColumn(webCol).insert(1);
    
    // Zeile 4: DB-Wert und Web-Wert setzen
    ws.getCell(`${dbCol}${headerRow + 1}`).value = 'DB-Wert';
    ws.getCell(`${webCol}${headerRow + 1}`).value = 'Web-Wert';
    
    // Formatierung der Unterüberschriften
    [dbCol, webCol].forEach(col => {
      ws.getCell(`${col}${headerRow + 1}`).alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getCell(`${col}${headerRow + 1}`).font = { bold: true, size: 12 };
    });
  });

  // Header-Zeilen einfrieren
  ws.views = [
    {
      state: 'frozen',
      xSplit: 0,
      ySplit: headerRow - 1
    }
  ];
}



// Neue Funktion: Setze Web-Daten und färbe entsprechend
function setWebDataAndColor(ws, targetRow, webData, dbData, headerRow, firstDataRow) {
  // Web-Daten in die entsprechenden Spalten setzen und einfärben
  
  // Materialkurztext (Web) - Spalte D
  if (webData.Produkttitel && webData.Produkttitel !== 'Nicht gefunden') {
    ws.getCell(`D${targetRow}`).value = webData.Produkttitel;
    const color = eqText(dbData.materialkurztext, webData.Produkttitel) ? 'green' : 'red';
    fillColor(ws, `D${targetRow}`, color);
  } else {
    fillColor(ws, `D${targetRow}`, 'orange');
  }

  // Her.-Artikelnummer (Web) - Spalte F
  if (webData['Weitere Artikelnummer'] && webData['Weitere Artikelnummer'] !== 'Nicht gefunden') {
    ws.getCell(`F${targetRow}`).value = webData['Weitere Artikelnummer'];
    const color = eqPart(dbData.artikelnummer, webData['Weitere Artikelnummer']) ? 'green' : 'red';
    fillColor(ws, `F${targetRow}`, color);
  } else {
    fillColor(ws, `F${targetRow}`, 'orange');
  }

  // Fert./Prüfhinweis (Web) - Spalte H
  if (webData.FertPruefhinweis && webData.FertPruefhinweis !== 'Nicht gefunden') {
    ws.getCell(`H${targetRow}`).value = webData.FertPruefhinweis;
    const color = eqText(dbData.fertPruefhinweis, webData.FertPruefhinweis) ? 'green' : 'red';
    fillColor(ws, `H${targetRow}`, color);
  } else {
    fillColor(ws, `H${targetRow}`, 'orange');
  }

  // Werkstoff (Web) - Spalte J
  if (webData.Werkstoff && webData.Werkstoff !== 'Nicht gefunden') {
    ws.getCell(`J${targetRow}`).value = webData.Werkstoff;
    const color = eqText(dbData.werkstoff, webData.Werkstoff) ? 'green' : 'red';
    fillColor(ws, `J${targetRow}`, color);
  } else {
    fillColor(ws, `J${targetRow}`, 'orange');
  }

  // Nettogewicht (Web) - Spalte L
  if (webData.Gewicht && webData.Gewicht !== 'Nicht gefunden') {
    const { value, unit } = parseWeight(webData.Gewicht);
    if (value != null) {
      ws.getCell(`L${targetRow}`).value = value;
      const color = eqWeight(dbData.nettogewicht, dbData.gewichtEinheit, webData.Gewicht) ? 'green' : 'red';
      fillColor(ws, `L${targetRow}`, color);
    } else {
      fillColor(ws, `L${targetRow}`, 'orange');
    }
  } else {
    fillColor(ws, `L${targetRow}`, 'orange');
  }

  // Abmessungen (Web)
  if (webData.Abmessung && webData.Abmessung !== 'Nicht gefunden') {
    const dims = parseDimensionsToLBH(webData.Abmessung);
    
    // Länge (Web) - Spalte N
    if (dims.L != null) {
      ws.getCell(`N${targetRow}`).value = dims.L;
      const color = (dbData.laenge != null && dbData.laenge === dims.L) ? 'green' : 'red';
      fillColor(ws, `N${targetRow}`, color);
    } else {
      fillColor(ws, `N${targetRow}`, 'orange');
    }
    
    // Breite (Web) - Spalte P
    if (dims.B != null) {
      ws.getCell(`P${targetRow}`).value = dims.B;
      const color = (dbData.breite != null && dbData.breite === dims.B) ? 'green' : 'red';
      fillColor(ws, `P${targetRow}`, color);
    } else {
      fillColor(ws, `P${targetRow}`, 'orange');
    }
    
    // Höhe (Web) - Spalte R
    if (dims.H != null) {
      ws.getCell(`R${targetRow}`).value = dims.H;
      const color = (dbData.hoehe != null && dbData.hoehe === dims.H) ? 'green' : 'red';
      fillColor(ws, `R${targetRow}`, color);
    } else {
      fillColor(ws, `R${targetRow}`, 'orange');
    }
  } else {
    // Alle Abmessungen orange markieren wenn nicht gefunden
    ['N', 'P', 'R'].forEach(col => {
      fillColor(ws, `${col}${targetRow}`, 'orange');
    });
  }

  // Debug-Logging
  console.log(`Row ${targetRow}: Web data processed for A2V ${webData.A2V || 'unknown'}`);
  console.log(`  - Produkttitel: ${webData.Produkttitel || 'Nicht gefunden'}`);
  console.log(`  - Weitere Artikelnummer: ${webData['Weitere Artikelnummer'] || 'Nicht gefunden'}`);
  console.log(`  - FertPruefhinweis: ${webData.FertPruefhinweis || 'Nicht gefunden'}`);
  console.log(`  - Werkstoff: ${webData.Werkstoff || 'Nicht gefunden'}`);
  console.log(`  - Gewicht: ${webData.Gewicht || 'Nicht gefunden'}`);
  console.log(`  - Abmessung: ${webData.Abmessung || 'Nicht gefunden'}`);
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

    // Das erste Arbeitsblatt verwenden (nicht neu erstellen)
    const ws = wb.getWorksheet(1);
    if (!ws) {
      return res.status(400).json({ error: 'Kein Arbeitsblatt in der Excel-Datei gefunden.' });
    }

    // Layout anpassen (neue Zeile 4 mit DB-Wert/Web-Wert hinzufügen)
    createComparisonLayout(ws, HEADER_ROW, FIRST_DATA_ROW);

    // 1) A2V-Nummern aus Spalte Z ab Zeile 4 einsammeln
    const tasks = [];
    const prodRows = [];
    const last = ws.lastRow?.number || 0;
    for (let r = FIRST_DATA_ROW; r <= last; r++) {
      const a2v = (ws.getCell(`${COLS.Z}${r}`).value || '').toString().trim().toUpperCase();
      if (a2v.startsWith('A2V')) {
        prodRows.push(r);
        tasks.push(a2v);
      }
    }

    // 2) Scrapen
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);

    // 3) Alle Datenzeilen nach unten verschieben (wegen der neuen Zeile 4)
    // Zuerst alle Zeilen ab Zeile 5 nach unten verschieben
    for (let r = last; r >= FIRST_DATA_ROW; r--) {
      // Alle Spalten von A bis Z kopieren
      for (let col = 'A'; col <= 'Z'; col++) {
        const sourceCell = ws.getCell(`${col}${r}`);
        if (sourceCell.value != null) {
          ws.getCell(`${col}${r + 1}`).value = sourceCell.value;
          // Ursprüngliche Zelle leeren
          ws.getCell(`${col}${r}`).value = null;
        }
      }
    }

    // 4) Web-Daten in die entsprechenden Web-Spalten eintragen
    for (const sourceRow of prodRows) {
      try {
        const a2v = (ws.getCell(`${COLS.Z}${sourceRow + 1}`).value || '').toString().trim().toUpperCase();
        const webData = resultsMap.get(a2v) || {};

        // DB-Daten aus der verschobenen Zeile extrahieren
        const dbData = {
          materialkurztext: ws.getCell(`${COLS.C}${sourceRow + 1}`).value,
          artikelnummer: ws.getCell(`${COLS.E}${sourceRow + 1}`).value,
          fertPruefhinweis: ws.getCell(`${COLS.N}${sourceRow + 1}`).value,
          werkstoff: ws.getCell(`${COLS.P}${sourceRow + 1}`).value,
          nettogewicht: ws.getCell(`${COLS.S}${sourceRow + 1}`).value,
          gewichtEinheit: ws.getCell(`${COLS.T}${sourceRow + 1}`).value,
          laenge: ws.getCell(`${COLS.U}${sourceRow + 1}`).value,
          breite: ws.getCell(`${COLS.V}${sourceRow + 1}`).value,
          hoehe: ws.getCell(`${COLS.W}${sourceRow + 1}`).value
        };

        console.log(`Processing row ${sourceRow + 1} with A2V: ${a2v}`);
        console.log(`DB data:`, dbData);

        // Web-Daten setzen und einfärben
        setWebDataAndColor(ws, sourceRow + 1, webData, dbData);

      } catch (error) {
        console.error(`Error processing row ${sourceRow + 1}:`, error);
        // Fehlerzeile überspringen und weitermachen
        continue;
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