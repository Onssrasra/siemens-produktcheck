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
  normalizeNCode,
  compareTextExact,
  compareWeightExact
} = require('./utils');
const { SiemensProductScraper, a2vUrl } = require('./scraper');

const app = express();
const PORT = process.env.PORT || 3000;
const SCRAPE_CONCURRENCY = Number(process.env.SCRAPE_CONCURRENCY || 4);
const HEADER_ROW = 3;
const FIRST_DATA_ROW = 4;

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const scraper = new SiemensProductScraper();

// Neue Farbkodierung: grün (gleich), rot (ungleich), orange (fehlt)
function fillColor(ws, addr, color) {
  if (!color) return;
  const map = {
    green: 'FFD5F4E6',   // Grün für exakte Übereinstimmung
    red: 'FFFDEAEA',     // Rot für Unterschiede
    orange: 'FFFFE6CC'   // Orange für fehlende Werte
  };
  ws.getCell(addr).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: map[color] || map.green } };
}

// Exakte Gleichheitstests ohne Toleranz
function eqText(a, b) {
  return compareTextExact(a, b);
}

function eqPart(a, b) { 
  const normA = normPartNo(a);
  const normB = normPartNo(b);
  return normA === normB;
}

function eqN(a, b) { 
  return normalizeNCode(a) === normalizeNCode(b); 
}

function eqDim(exL, exB, exH, webTxt) {
  const L = toNumber(exL);
  const B = toNumber(exB);
  const H = toNumber(exH);
  const w = parseDimensionsToLBH(webTxt);
  
  if (L == null || B == null || H == null || w.L == null || w.B == null || w.H == null) return false;
  return L === w.L && B === w.B && H === w.H;
}

function eqWeight(exWeight, webVal) {
  return compareWeightExact(exWeight, webVal);
}

// Neue Funktion zum Erstellen der Ausgangstabelle mit dem gewünschten Layout
function createOutputLayout(ws, headerRow, firstDataRow) {
  // Spaltenblöcke definieren: [startCol, endCol, title, dbCol, webCol]
  const blocks = [
    ['C', 'D', 'Materialkurztext', 'C', 'D'],
    ['E', 'F', 'Her.-Artikelnummer', 'E', 'F'],
    ['G', 'H', 'Fert./Prüfhinweis', 'G', 'H'],
    ['I', 'J', 'Werkstoff', 'I', 'J'],
    ['K', 'L', 'Nettogewicht', 'K', 'L'],
    ['M', 'N', 'Länge', 'M', 'N'],
    ['O', 'P', 'Breite', 'O', 'P'],
    ['Q', 'R', 'Höhe', 'Q', 'R']
  ];
  
  // Header-Zeile 3: Hauptüberschriften als Blöcke
  blocks.forEach(([startCol, endCol, title]) => {
    ws.mergeCells(`${startCol}${headerRow}:${endCol}${headerRow}`);
    ws.getCell(`${startCol}${headerRow}`).value = title;
    ws.getCell(`${startCol}${headerRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(`${startCol}${headerRow}`).font = { bold: true };
  });
  
  // Header-Zeile 4: Unterüberschriften
  blocks.forEach(([startCol, endCol, title, dbCol, webCol]) => {
    ws.getCell(`${dbCol}${headerRow + 1}`).value = 'DB-Wert';
    ws.getCell(`${webCol}${headerRow + 1}`).value = 'Web-Wert';
    
    // Formatierung der Unterüberschriften
    [dbCol, webCol].forEach(col => {
      ws.getCell(`${col}${headerRow + 1}`).alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getCell(`${col}${headerRow + 1}`).font = { bold: true };
    });
  });
  
  // Produkt-ID Spalte (A2V) - Spalte A
  ws.getCell(`A${headerRow}`).value = 'Produkt-ID';
  ws.getCell(`A${headerRow + 1}`).value = 'A2V';
  ws.getCell(`A${headerRow}`).font = { bold: true };
  ws.getCell(`A${headerRow + 1}`).font = { bold: true };
  ws.getCell(`A${headerRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getCell(`A${headerRow + 1}`).alignment = { horizontal: 'center', vertical: 'middle' };
}

// Neue Funktion zum Mapping der Spalten basierend auf Header-Text
function findColumnsByHeader(ws, headerRow) {
  const columns = {};
  const headerMap = {
    'Materialkurztext': 'C',
    'Her.-Artikelnummer': 'E', 
    'Fert./Prüfhinweis': 'G',
    'Werkstoff': 'I',
    'Nettogewicht': 'K',
    'Länge': 'M',
    'Breite': 'O',
    'Höhe': 'Q'
  };
  
  // Suche nach den relevanten Spalten basierend auf Header-Text
  for (let col = 1; col <= ws.columnCount; col++) {
    const cellValue = ws.getCell(col, headerRow).value;
    if (cellValue) {
      const headerText = String(cellValue).trim();
      for (const [key, defaultCol] of Object.entries(headerMap)) {
        if (headerText.toLowerCase().includes(key.toLowerCase())) {
          columns[key] = ExcelJS.utils.getColumnKey(col);
          break;
        }
      }
    }
  }
  
  // Fallback auf Standard-Spalten falls nicht gefunden
  Object.keys(headerMap).forEach(key => {
    if (!columns[key]) {
      columns[key] = headerMap[key];
    }
  });
  
  return columns;
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

    // Aktives Blatt verarbeiten
    const ws = wb.getWorksheet(1); // Erstes Blatt
    if (!ws) return res.status(400).json({ error: 'Kein Arbeitsblatt gefunden.' });

    // Spalten basierend auf Header-Text finden
    const columns = findColumnsByHeader(ws, HEADER_ROW);
    
    // A2V-Spalte finden (normalerweise Spalte Z oder durch Header-Text)
    let a2vColumn = 'Z'; // Fallback
    for (let col = 1; col <= ws.columnCount; col++) {
      const cellValue = ws.getCell(col, HEADER_ROW).value;
      if (cellValue && String(cellValue).toLowerCase().includes('produkt') && 
          String(cellValue).toLowerCase().includes('id')) {
        a2vColumn = ExcelJS.utils.getColumnKey(col);
        break;
      }
    }

    // 1) A2V-Nummern aus der gefundenen Spalte ab Zeile 4 einsammeln
    const tasks = [];
    const dataRows = [];
    const last = ws.lastRow?.number || 0;
    
    for (let r = FIRST_DATA_ROW; r <= last; r++) {
      const a2v = (ws.getCell(`${a2vColumn}${r}`).value || '').toString().trim().toUpperCase();
      if (a2v.startsWith('A2V')) {
        tasks.push(a2v);
        dataRows.push(r);
      }
    }

    if (tasks.length === 0) {
      return res.status(400).json({ error: 'Keine A2V-Nummern in der Tabelle gefunden.' });
    }

    // 2) Web-Daten scrapen
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);

    // 3) Neue Ausgangstabelle mit dem gewünschten Layout erstellen
    const newWs = wb.addWorksheet('Produktvergleich');
    createOutputLayout(newWs, HEADER_ROW, FIRST_DATA_ROW);

    // 4) Daten verarbeiten und in die neue Tabelle schreiben
    for (let i = 0; i < dataRows.length; i++) {
      const sourceRow = dataRows[i];
      const targetRow = FIRST_DATA_ROW + i;
      const a2v = tasks[i];
      const web = resultsMap.get(a2v) || {};

      // Produkt-ID (A2V)
      newWs.getCell(`A${targetRow}`).value = a2v;

      // DB-Werte aus der ursprünglichen Tabelle kopieren
      if (columns['Materialkurztext']) {
        newWs.getCell(`C${targetRow}`).value = ws.getCell(`${columns['Materialkurztext']}${sourceRow}`).value;
      }
      if (columns['Her.-Artikelnummer']) {
        newWs.getCell(`E${targetRow}`).value = ws.getCell(`${columns['Her.-Artikelnummer']}${sourceRow}`).value;
      }
      if (columns['Fert./Prüfhinweis']) {
        newWs.getCell(`G${targetRow}`).value = ws.getCell(`${columns['Fert./Prüfhinweis']}${sourceRow}`).value;
      }
      if (columns['Werkstoff']) {
        newWs.getCell(`I${targetRow}`).value = ws.getCell(`${columns['Werkstoff']}${sourceRow}`).value;
      }
      if (columns['Nettogewicht']) {
        newWs.getCell(`K${targetRow}`).value = ws.getCell(`${columns['Nettogewicht']}${sourceRow}`).value;
      }
      if (columns['Länge']) {
        newWs.getCell(`M${targetRow}`).value = ws.getCell(`${columns['Länge']}${sourceRow}`).value;
      }
      if (columns['Breite']) {
        newWs.getCell(`O${targetRow}`).value = ws.getCell(`${columns['Breite']}${sourceRow}`).value;
      }
      if (columns['Höhe']) {
        newWs.getCell(`Q${targetRow}`).value = ws.getCell(`${columns['Höhe']}${sourceRow}`).value;
      }

      // Web-Werte setzen und vergleichen
      
      // Materialkurztext (Web)
      if (web.Produkttitel && web.Produkttitel !== 'Nicht gefunden') {
        newWs.getCell(`D${targetRow}`).value = web.Produkttitel;
        const dbVal = newWs.getCell(`C${targetRow}`).value;
        const isEqual = eqText(dbVal, web.Produkttitel);
        fillColor(newWs, `D${targetRow}`, isEqual ? 'green' : 'red');
      } else {
        fillColor(newWs, `D${targetRow}`, 'orange');
      }

      // Her.-Artikelnummer (Web)
      if (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer'] !== 'Nicht gefunden') {
        newWs.getCell(`F${targetRow}`).value = web['Weitere Artikelnummer'];
        const dbVal = newWs.getCell(`E${targetRow}`).value;
        const isEqual = eqPart(dbVal, web['Weitere Artikelnummer']);
        fillColor(newWs, `F${targetRow}`, isEqual ? 'green' : 'red');
      } else {
        fillColor(newWs, `F${targetRow}`, 'orange');
      }

      // Fert./Prüfhinweis (Web) - wird nicht direkt gescraped, daher orange
      fillColor(newWs, `H${targetRow}`, 'orange');

      // Werkstoff (Web)
      if (web.Werkstoff && web.Werkstoff !== 'Nicht gefunden') {
        newWs.getCell(`J${targetRow}`).value = web.Werkstoff;
        const dbVal = newWs.getCell(`I${targetRow}`).value;
        const isEqual = eqText(dbVal, web.Werkstoff);
        fillColor(newWs, `J${targetRow}`, isEqual ? 'green' : 'red');
      } else {
        fillColor(newWs, `J${targetRow}`, 'orange');
      }

      // Nettogewicht (Web)
      if (web.Gewicht && web.Gewicht !== 'Nicht gefunden') {
        const { value: weightValue, unit: weightUnit } = parseWeight(web.Gewicht);
        if (weightValue != null) {
          newWs.getCell(`L${targetRow}`).value = weightValue;
          const dbVal = newWs.getCell(`K${targetRow}`).value;
          const isEqual = eqWeight(dbVal, web.Gewicht);
          fillColor(newWs, `L${targetRow}`, isEqual ? 'green' : 'red');
        } else {
          fillColor(newWs, `L${targetRow}`, 'orange');
        }
      } else {
        fillColor(newWs, `L${targetRow}`, 'orange');
      }

      // Abmessungen (Web)
      if (web.Abmessung && web.Abmessung !== 'Nicht gefunden') {
        const dims = parseDimensionsToLBH(web.Abmessung);
        
        // Länge
        if (dims.L != null) {
          newWs.getCell(`N${targetRow}`).value = dims.L;
          const dbVal = newWs.getCell(`M${targetRow}`).value;
          const isEqual = toNumber(dbVal) === dims.L;
          fillColor(newWs, `N${targetRow}`, isEqual ? 'green' : 'red');
        } else {
          fillColor(newWs, `N${targetRow}`, 'orange');
        }
        
        // Breite
        if (dims.B != null) {
          newWs.getCell(`P${targetRow}`).value = dims.B;
          const dbVal = newWs.getCell(`O${targetRow}`).value;
          const isEqual = toNumber(dbVal) === dims.B;
          fillColor(newWs, `P${targetRow}`, isEqual ? 'green' : 'red');
        } else {
          fillColor(newWs, `P${targetRow}`, 'orange');
        }
        
        // Höhe
        if (dims.H != null) {
          newWs.getCell(`R${targetRow}`).value = dims.H;
          const dbVal = newWs.getCell(`Q${targetRow}`).value;
          const isEqual = toNumber(dbVal) === dims.H;
          fillColor(newWs, `R${targetRow}`, isEqual ? 'green' : 'red');
        } else {
          fillColor(newWs, `R${targetRow}`, 'orange');
        }
      } else {
        // Alle Abmessungen orange markieren wenn nicht gefunden
        fillColor(newWs, `N${targetRow}`, 'orange');
        fillColor(newWs, `P${targetRow}`, 'orange');
        fillColor(newWs, `R${targetRow}`, 'orange');
      }
    }

    // Ursprüngliches Blatt löschen und neues umbenennen
    wb.removeWorksheet(ws.id);
    newWs.name = 'Produktvergleich';

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