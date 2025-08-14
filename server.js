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

// Ursprüngliche Spalten-Definition
const ORIGINAL_COLS = { Z:'Z', E:'E', C:'C', S:'S', T:'T', U:'U', V:'V', W:'W', P:'P', N:'N' };

// DB/Web-Spaltenpaare - hier definieren wir welche Spalten DB/Web-Paare bekommen
const DB_WEB_PAIRS = [
  { original: 'C', dbCol: null, webCol: null, label: 'Material-Kurztext' },
  { original: 'E', dbCol: null, webCol: null, label: 'Herstellartikelnummer' },
  { original: 'N', dbCol: null, webCol: null, label: 'Fert./Prüfhinweis' },
  { original: 'P', dbCol: null, webCol: null, label: 'Werkstoff' },
  { original: 'S', dbCol: null, webCol: null, label: 'Nettogewicht' },
  { original: 'U', dbCol: null, webCol: null, label: 'Länge' },
  { original: 'V', dbCol: null, webCol: null, label: 'Breite' },
  { original: 'W', dbCol: null, webCol: null, label: 'Höhe' }
];

const HEADER_ROW = 3;
const LABEL_ROW = 4; // Neue Zeile für "DB-Wert"/"Web-Wert" Labels
const FIRST_DATA_ROW = 5; // Daten beginnen jetzt ab Zeile 5

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const scraper = new SiemensProductScraper();

// Funktion um Spaltenbuchstaben zu berechnen (A, B, C, ... Z, AA, AB, ...)
function getColumnLetter(index) {
  let result = '';
  while (index > 0) {
    index--;
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26);
  }
  return result;
}

// Funktion um Spaltenindex aus Buchstaben zu berechnen
function getColumnIndex(letter) {
  let index = 0;
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64);
  }
  return index;
}

// Funktion um neue Spaltenstruktur zu berechnen
function calculateNewColumnStructure(ws) {
  const newStructure = {
    pairs: [],
    otherCols: new Map(),
    totalInsertedCols: 0
  };
  
  let insertedCols = 0;
  
  // Für jedes DB/Web-Paar
  for (const pair of DB_WEB_PAIRS) {
    const originalIndex = getColumnIndex(pair.original);
    const adjustedOriginalIndex = originalIndex + insertedCols;
    
    pair.dbCol = getColumnLetter(adjustedOriginalIndex);
    pair.webCol = getColumnLetter(adjustedOriginalIndex + 1);
    
    newStructure.pairs.push(pair);
    insertedCols++;
  }
  
  newStructure.totalInsertedCols = insertedCols;
  
  // Andere Spalten (die nicht in DB/Web-Paaren sind) anpassen
  const lastRow = ws.lastRow?.number || 0;
  const lastCol = ws.lastColumn?.number || 0;
  
  for (let colIndex = 1; colIndex <= lastCol; colIndex++) {
    const originalLetter = getColumnLetter(colIndex);
    
    // Prüfen ob diese Spalte ein DB/Web-Paar ist
    const isPairColumn = DB_WEB_PAIRS.some(pair => pair.original === originalLetter);
    
    if (!isPairColumn) {
      // Berechnen wie viele Spalten vor dieser Spalte eingefügt wurden
      let insertedBefore = 0;
      for (const pair of DB_WEB_PAIRS) {
        if (getColumnIndex(pair.original) < colIndex) {
          insertedBefore++;
        }
      }
      
      const newLetter = getColumnLetter(colIndex + insertedBefore);
      newStructure.otherCols.set(originalLetter, newLetter);
    }
  }
  
  return newStructure;
}

function fillColor(ws, addr, color) {
  if (!color) return;
  const map = {
    green: 'FFD5F4E6',  // Hellgrün für übereinstimmende Werte
    red:   'FFFDEAEA',  // Hellrot für unterschiedliche Werte
    orange: 'FFFFEAA7'  // Orange für fehlende Werte
  };
  ws.getCell(addr).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: map[color] || map.green } };
}

function copyColumnFormatting(ws, fromCol, toCol, rowStart, rowEnd) {
  for (let row = rowStart; row <= rowEnd; row++) {
    const fromCell = ws.getCell(`${fromCol}${row}`);
    const toCell = ws.getCell(`${toCol}${row}`);
    
    // Kopiere Formatierung
    if (fromCell.fill) {
      toCell.fill = fromCell.fill;
    }
    if (fromCell.font) {
      toCell.font = fromCell.font;
    }
    if (fromCell.border) {
      toCell.border = fromCell.border;
    }
    if (fromCell.alignment) {
      toCell.alignment = fromCell.alignment;
    }
    if (fromCell.style) {
      // Kopiere andere Style-Eigenschaften
      Object.assign(toCell.style, fromCell.style);
    }
  }
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
function eqWeight(exS, webVal) {
  const { value: wv } = parseWeight(webVal);
  if (wv == null) return false;
  const exNum = toNumber(exS);
  if (exNum == null) return false;
  return Math.abs(exNum - wv) < 1e-9;
}
function eqDimension(exVal, webDimText, dimType) {
  const exNum = toNumber(exVal);
  const webDim = parseDimensionsToLBH(webDimText);
  if (exNum == null) return false;
  
  let webVal = null;
  if (dimType === 'L') webVal = webDim.L;
  else if (dimType === 'B') webVal = webDim.B;
  else if (dimType === 'H') webVal = webDim.H;
  
  if (webVal == null) return false;
  return exNum === webVal;
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

    // 1) A2V-Nummern aus Spalte Z ab ursprünglicher Zeile 4 einsammeln
    const tasks = [];
    const rowsPerSheet = new Map(); // ws -> [rowIndex,...]
    
    for (const ws of wb.worksheets) {
      const indices = [];
      const last = ws.lastRow?.number || 0;
      
      // A2V-Nummern einsammeln (vor Strukturänderung)
      for (let r = FIRST_DATA_ROW - 1; r <= last; r++) { // -1 weil wir noch die alte Struktur haben
        const a2v = (ws.getCell(`${ORIGINAL_COLS.Z}${r}`).value || '').toString().trim().toUpperCase();
        if (a2v.startsWith('A2V')) {
          indices.push(r);
          tasks.push(a2v);
        }
      }
      rowsPerSheet.set(ws, indices);
    }

    // 2) Scrapen
    console.log(`Scraping ${tasks.length} A2V numbers...`);
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);

    // 3) Excel-Struktur für jedes Worksheet umbauen
    for (const ws of wb.worksheets) {
      console.log(`Processing worksheet: ${ws.name}`);
      
      // 3.1) Neue Spaltenstruktur berechnen
      const structure = calculateNewColumnStructure(ws);
      
      // 3.2) Spalten einfügen (von rechts nach links, um Indizes nicht zu verschieben)
      const pairsReversed = [...structure.pairs].reverse();
      for (const pair of pairsReversed) {
        const insertPos = getColumnIndex(pair.original) + 1;
        ws.spliceColumns(insertPos, 0, [null]);
        console.log(`Inserted column after ${pair.original} for ${pair.label}`);
      }
      
      // 3.3) Label-Zeile (Zeile 4) einfügen
      ws.spliceRows(LABEL_ROW, 0, [null]);
      console.log(`Inserted label row at position ${LABEL_ROW}`);
      
      // 3.4) Zeile 2 (technische Codes) und Zeile 3 (Klartext) für neue Spalten duplizieren
      for (const pair of structure.pairs) {
        // Werte aus der DB-Spalte holen
        const dbTechCode = ws.getCell(`${pair.dbCol}2`).value;
        const dbKlartext = ws.getCell(`${pair.dbCol}3`).value;
        
        // In Web-Spalte duplizieren
        ws.getCell(`${pair.webCol}2`).value = dbTechCode;
        ws.getCell(`${pair.webCol}3`).value = dbKlartext;
        
        // Formatierung von DB-Spalte auf Web-Spalte kopieren (Zeilen 1-3)
        copyColumnFormatting(ws, pair.dbCol, pair.webCol, 1, 3);
        
        // Label-Zeile befüllen
        ws.getCell(`${pair.dbCol}${LABEL_ROW}`).value = 'DB-Wert';
        ws.getCell(`${pair.webCol}${LABEL_ROW}`).value = 'Web-Wert';
        
        console.log(`Set labels for ${pair.label}: ${pair.dbCol} (DB-Wert), ${pair.webCol} (Web-Wert)`);
      }
      
      // 3.5) A2V-Daten verarbeiten und Web-Werte eintragen
      const prodRows = rowsPerSheet.get(ws) || [];
      
      for (const originalRow of prodRows) {
        // Zeile wurde um 1 nach unten verschoben durch Label-Zeile
        const currentRow = originalRow + 1;
        
        // A2V-Nummer bestimmen - neue Spaltenposition für Z berechnen
        let zCol = ORIGINAL_COLS.Z;
        if (structure.otherCols.has(ORIGINAL_COLS.Z)) {
          zCol = structure.otherCols.get(ORIGINAL_COLS.Z);
        }
        
        const a2v = (ws.getCell(`${zCol}${currentRow}`).value || '').toString().trim().toUpperCase();
        const web = resultsMap.get(a2v) || {};
        
        console.log(`Processing row ${currentRow}, A2V: ${a2v}`);
        
        // 3.6) Web-Werte in die entsprechenden Web-Spalten eintragen
        for (const pair of structure.pairs) {
          const dbValue = ws.getCell(`${pair.dbCol}${currentRow}`).value;
          let webValue = null;
          let isEqual = false;
          
          switch (pair.original) {
            case 'C': // Material-Kurztext
              webValue = (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') ? web.Produkttitel : null;
              isEqual = webValue ? eqText(dbValue || '', webValue) : false;
              break;
              
            case 'E': // Herstellartikelnummer
              webValue = (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden') 
                ? web['Weitere Artikelnummer'] : a2v;
              isEqual = eqPart(dbValue || a2v, webValue);
              break;
              
            case 'N': // Fert./Prüfhinweis (Materialklassifizierung)
              if (web.Materialklassifizierung && web.Materialklassifizierung !== 'Nicht gefunden') {
                const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung));
                if (code) {
                  webValue = code;
                  isEqual = eqN(dbValue || '', code);
                }
              }
              break;
              
            case 'P': // Werkstoff
              webValue = (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') ? web.Werkstoff : null;
              isEqual = webValue ? eqText(dbValue || '', webValue) : false;
              break;
              
            case 'S': // Nettogewicht
              if (web.Gewicht && web.Gewicht !== 'Nicht gefunden') {
                const { value } = parseWeight(web.Gewicht);
                if (value != null) {
                  webValue = value;
                  isEqual = eqWeight(dbValue, web.Gewicht);
                }
              }
              break;
              
            case 'U': // Länge
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.L != null) {
                  webValue = d.L;
                  isEqual = eqDimension(dbValue, web.Abmessung, 'L');
                }
              }
              break;
              
            case 'V': // Breite
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.B != null) {
                  webValue = d.B;
                  isEqual = eqDimension(dbValue, web.Abmessung, 'B');
                }
              }
              break;
              
            case 'W': // Höhe
              if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
                const d = parseDimensionsToLBH(web.Abmessung);
                if (d.H != null) {
                  webValue = d.H;
                  isEqual = eqDimension(dbValue, web.Abmessung, 'H');
                }
              }
              break;
          }
          
          // Web-Wert eintragen falls vorhanden
          if (webValue !== null) {
            ws.getCell(`${pair.webCol}${currentRow}`).value = webValue;
            
            // Farbkodierung NUR für Web-Spalte
            const color = isEqual ? 'green' : 'red';
            fillColor(ws, `${pair.webCol}${currentRow}`, color);
            
            console.log(`${pair.label}: DB="${dbValue}" vs Web="${webValue}" -> ${isEqual ? 'EQUAL' : 'DIFFERENT'}`);
          } else {
            // Kein Web-Wert verfügbar - orange markieren wenn DB-Wert vorhanden ist
            if (dbValue !== null && dbValue !== undefined && dbValue !== '') {
              fillColor(ws, `${pair.webCol}${currentRow}`, 'orange');
              console.log(`${pair.label}: DB="${dbValue}" vs Web=MISSING -> ORANGE`);
            }
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
