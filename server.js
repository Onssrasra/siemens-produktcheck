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
  // Zeile 3: Hauptüberschriften als Blöcke
  const mainHeaders = [
    { name: 'Material', cols: ['A', 'A'] },
    { name: 'Herstellername', cols: ['B', 'B'] },
    { name: 'Materialkurztext', cols: ['C', 'D'] },
    { name: 'Her.-Artikelnummer', cols: ['E', 'F'] },
    { name: 'Fert./Prüfhinweis', cols: ['G', 'H'] },
    { name: 'Werkstoff', cols: ['I', 'J'] },
    { name: 'Nettogewicht', cols: ['K', 'L'] },
    { name: 'Länge', cols: ['M', 'N'] },
    { name: 'Breite', cols: ['O', 'P'] },
    { name: 'Höhe', cols: ['Q', 'R'] },
    { name: 'Produkt-ID', cols: ['S', 'S'] }
  ];

  // Hauptüberschriften setzen und Zellen zusammenfassen
  mainHeaders.forEach((header, index) => {
    const startCol = header.cols[0];
    const endCol = header.cols[1];
    ws.getCell(`${startCol}${headerRow}`).value = header.name;
    
    // Nur zusammenfassen wenn es zwei verschiedene Spalten sind
    if (startCol !== endCol) {
      ws.mergeCells(`${startCol}${headerRow}:${endCol}${headerRow}`);
    }
    
    // Zentrieren und Formatierung
    ws.getCell(`${startCol}${headerRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(`${startCol}${headerRow}`).font = { bold: true };
  });

  // Zeile 4: Unterüberschriften (DB-Wert und Web-Wert) nur für die Blöcke
  const blockHeaders = mainHeaders.filter(h => h.cols[0] !== h.cols[1]);
  blockHeaders.forEach((header) => {
    const dbCol = header.cols[0];
    const webCol = header.cols[1];
    
    ws.getCell(`${dbCol}${headerRow + 1}`).value = 'DB-Wert';
    ws.getCell(`${webCol}${headerRow + 1}`).value = 'Web-Wert';
    
    // Formatierung der Unterüberschriften
    [dbCol, webCol].forEach(col => {
      ws.getCell(`${col}${headerRow + 1}`).alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getCell(`${col}${headerRow + 1}`).font = { bold: true, size: 12 };
    });
  });

  // Spaltenbreiten anpassen
  const columnWidths = {
    'A': 15, // Material
    'B': 20, // Herstellername
    'C': 20, // Materialkurztext DB
    'D': 20, // Materialkurztext Web
    'E': 20, // Her.-Artikelnummer DB
    'F': 20, // Her.-Artikelnummer Web
    'G': 20, // Fert./Prüfhinweis DB
    'H': 20, // Fert./Prüfhinweis Web
    'I': 20, // Werkstoff DB
    'J': 20, // Werkstoff Web
    'K': 15, // Nettogewicht DB
    'L': 15, // Nettogewicht Web
    'M': 12, // Länge DB
    'N': 12, // Länge Web
    'O': 12, // Breite DB
    'P': 12, // Breite Web
    'Q': 12, // Höhe DB
    'R': 12, // Höhe Web
    'S': 20  // Produkt-ID
  };

  Object.entries(columnWidths).forEach(([col, width]) => {
    ws.getColumn(col).width = width;
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

// Neue Funktion: Kopiere DB-Daten in die neue Struktur
function copyDBDataToNewLayout(ws, sourceRow, targetRow, headerRow, firstDataRow) {
  // Mapping der alten Spalten zu den neuen DB-Spalten
  const columnMapping = {
    'A': 'A', // Material
    'B': 'B', // Herstellername
    'C': 'C', // Materialkurztext DB
    'E': 'E', // Her.-Artikelnummer DB
    'N': 'G', // Fert./Prüfhinweis DB
    'P': 'I', // Werkstoff DB
    'S': 'K', // Nettogewicht DB
    'U': 'M', // Länge DB
    'V': 'O', // Breite DB
    'W': 'Q', // Höhe DB
    'Z': 'S'  // Produkt-ID
  };

  // DB-Daten kopieren
  Object.entries(columnMapping).forEach(([sourceCol, targetCol]) => {
    const sourceCell = ws.getCell(`${sourceCol}${sourceRow}`);
    if (sourceCell.value != null) {
      ws.getCell(`${targetCol}${targetRow}`).value = sourceCell.value;
    }
  });

  // Debug-Logging
  console.log(`Copying DB data from row ${sourceRow} to row ${targetRow}`);
  console.log(`Column mapping:`, columnMapping);
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

    // Neues Arbeitsblatt für den Vergleich erstellen
    const newWs = wb.addWorksheet('DB_vs_Web_Vergleich');
    
    // Layout erstellen
    createComparisonLayout(newWs, HEADER_ROW, FIRST_DATA_ROW);

    // 1) A2V-Nummern aus Spalte Z ab Zeile 4 einsammeln
    const tasks = [];
    const rowsPerSheet = new Map(); // ws -> [rowIndex,...]
    for (const ws of wb.worksheets) {
      if (ws.name === 'DB_vs_Web_Vergleich') continue; // Neues Blatt überspringen
      
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

    // 3) Neue Tabelle mit DB vs Web Daten erstellen
    let newRowIndex = FIRST_DATA_ROW;
    
    for (const ws of wb.worksheets) {
      if (ws.name === 'DB_vs_Web_Vergleich') continue;
      
      const prodRows = rowsPerSheet.get(ws) || [];
      for (const sourceRow of prodRows) {
        try {
          const a2v = (ws.getCell(`${COLS.Z}${sourceRow}`).value || '').toString().trim().toUpperCase();
          const webData = resultsMap.get(a2v) || {};

          // DB-Daten aus der Quelltabelle extrahieren
          const dbData = {
            materialkurztext: ws.getCell(`${COLS.C}${sourceRow}`).value,
            artikelnummer: ws.getCell(`${COLS.E}${sourceRow}`).value,
            fertPruefhinweis: ws.getCell(`${COLS.N}${sourceRow}`).value,
            werkstoff: ws.getCell(`${COLS.P}${sourceRow}`).value,
            nettogewicht: ws.getCell(`${COLS.S}${sourceRow}`).value,
            gewichtEinheit: ws.getCell(`${COLS.T}${sourceRow}`).value,
            laenge: ws.getCell(`${COLS.U}${sourceRow}`).value,
            breite: ws.getCell(`${COLS.V}${sourceRow}`).value,
            hoehe: ws.getCell(`${COLS.W}${sourceRow}`).value
          };

          console.log(`Processing row ${sourceRow} with A2V: ${a2v}`);
          console.log(`DB data:`, dbData);

          // DB-Daten in die neue Struktur kopieren
          copyDBDataToNewLayout(newWs, sourceRow, newRowIndex, HEADER_ROW, FIRST_DATA_ROW);

          // Web-Daten setzen und einfärben
          setWebDataAndColor(newWs, newRowIndex, webData, dbData, HEADER_ROW, FIRST_DATA_ROW);

          newRowIndex++;
        } catch (error) {
          console.error(`Error processing row ${sourceRow}:`, error);
          // Fehlerzeile überspringen und weitermachen
          continue;
        }
      }
    }

    // Alle anderen Arbeitsblätter entfernen
    const sheetsToRemove = [];
    wb.worksheets.forEach(ws => {
      if (ws.name !== 'DB_vs_Web_Vergleich') {
        sheetsToRemove.push(ws);
      }
    });
    sheetsToRemove.forEach(ws => wb.removeWorksheet(ws.id));

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