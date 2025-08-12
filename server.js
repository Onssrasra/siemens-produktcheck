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
    try {
      ws.mergeCells(`${startCol}${headerRow}:${endCol}${headerRow}`);
      const cell = ws.getCell(`${startCol}${headerRow}`);
      cell.value = title;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.font = { bold: true };
    } catch (error) {
      console.log(`Fehler beim Zusammenführen von ${startCol}${headerRow}:${endCol}${headerRow}:`, error.message);
    }
  });
  
  // Header-Zeile 4: Unterüberschriften
  blocks.forEach(([startCol, endCol, title, dbCol, webCol]) => {
    // DB-Wert
    const dbCell = ws.getCell(`${dbCol}${headerRow + 1}`);
    dbCell.value = 'DB-Wert';
    dbCell.alignment = { horizontal: 'center', vertical: 'middle' };
    dbCell.font = { bold: true };
    
    // Web-Wert
    const webCell = ws.getCell(`${webCol}${headerRow + 1}`);
    webCell.value = 'Web-Wert';
    webCell.alignment = { horizontal: 'center', vertical: 'middle' };
    webCell.font = { bold: true };
  });
  
  // Produkt-ID Spalte (A2V) - Spalte A
  const idCell = ws.getCell(`A${headerRow}`);
  idCell.value = 'Produkt-ID';
  idCell.font = { bold: true };
  idCell.alignment = { horizontal: 'center', vertical: 'middle' };
  
  const a2vCell = ws.getCell(`A${headerRow + 1}`);
  a2vCell.value = 'A2V';
  a2vCell.font = { bold: true };
  a2vCell.alignment = { horizontal: 'center', vertical: 'middle' };
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

    console.log('Verarbeite Excel-Datei:', req.file.originalname, 'Größe:', req.file.size);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);

    // Aktives Blatt verarbeiten
    const ws = wb.getWorksheet(1); // Erstes Blatt
    if (!ws) return res.status(400).json({ error: 'Kein Arbeitsblatt gefunden.' });

    console.log('Arbeitsblatt gefunden:', ws.name, 'Zeilen:', ws.rowCount, 'Spalten:', ws.columnCount);

    // Spalten basierend auf Header-Text finden
    const columns = findColumnsByHeader(ws, HEADER_ROW);
    console.log('Gefundene Spalten:', columns);
    
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
    console.log('A2V-Spalte gefunden:', a2vColumn);

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

    console.log('Gefundene A2V-Nummern:', tasks.length, tasks.slice(0, 5));

    if (tasks.length === 0) {
      return res.status(400).json({ error: 'Keine A2V-Nummern in der Tabelle gefunden.' });
    }

    // 2) Web-Daten scrapen
    console.log('Starte Web-Scraping für', tasks.length, 'Produkte...');
    const resultsMap = await scraper.scrapeMany(tasks, SCRAPE_CONCURRENCY);
    console.log('Web-Scraping abgeschlossen,', resultsMap.size, 'Ergebnisse erhalten');

    // 3) Neue Ausgangstabelle mit dem gewünschten Layout erstellen
    console.log('Erstelle neue Ausgangstabelle...');
    const newWs = wb.addWorksheet('Produktvergleich');
    createOutputLayout(newWs, HEADER_ROW, FIRST_DATA_ROW);
    console.log('Layout erstellt');

    // 4) Daten verarbeiten und in die neue Tabelle schreiben
    console.log('Verarbeite Daten...');
    for (let i = 0; i < dataRows.length; i++) {
      const sourceRow = dataRows[i];
      const targetRow = FIRST_DATA_ROW + i;
      const a2v = tasks[i];
      const web = resultsMap.get(a2v) || {};

      console.log(`Verarbeite Zeile ${i + 1}/${dataRows.length}: ${a2v}`);

      // Produkt-ID (A2V)
      newWs.getCell(`A${targetRow}`).value = a2v;

      // DB-Werte aus der ursprünglichen Tabelle kopieren
      if (columns['Materialkurztext']) {
        const sourceValue = ws.getCell(`${columns['Materialkurztext']}${sourceRow}`).value;
        newWs.getCell(`C${targetRow}`).value = sourceValue;
      }
      if (columns['Her.-Artikelnummer']) {
        const sourceValue = ws.getCell(`${columns['Her.-Artikelnummer']}${sourceRow}`).value;
        newWs.getCell(`E${targetRow}`).value = sourceValue;
      }
      if (columns['Fert./Prüfhinweis']) {
        const sourceValue = ws.getCell(`${columns['Fert./Prüfhinweis']}${sourceRow}`).value;
        newWs.getCell(`G${targetRow}`).value = sourceValue;
      }
      if (columns['Werkstoff']) {
        const sourceValue = ws.getCell(`${columns['Werkstoff']}${sourceRow}`).value;
        newWs.getCell(`I${targetRow}`).value = sourceValue;
      }
      if (columns['Nettogewicht']) {
        const sourceValue = ws.getCell(`${columns['Nettogewicht']}${sourceRow}`).value;
        newWs.getCell(`K${targetRow}`).value = sourceValue;
      }
      if (columns['Länge']) {
        const sourceValue = ws.getCell(`${columns['Länge']}${sourceRow}`).value;
        newWs.getCell(`M${targetRow}`).value = sourceValue;
      }
      if (columns['Breite']) {
        const sourceValue = ws.getCell(`${columns['Breite']}${sourceRow}`).value;
        newWs.getCell(`O${targetRow}`).value = sourceValue;
      }
      if (columns['Höhe']) {
        const sourceValue = ws.getCell(`${columns['Höhe']}${sourceRow}`).value;
        newWs.getCell(`Q${targetRow}`).value = sourceValue;
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
    console.log('Lösche ursprüngliches Blatt und benenne neues um...');
    wb.removeWorksheet(ws.id);
    newWs.name = 'Produktvergleich';

    console.log('Erstelle Excel-Datei...');
    const out = await wb.xlsx.writeBuffer();
    console.log('Excel-Datei erstellt, Größe:', out.length);
    
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="DB_Produktvergleich_verarbeitet.xlsx"');
    res.send(Buffer.from(out));
  } catch (err) {
    console.error('Fehler bei der Excel-Verarbeitung:', err);
    res.status(500).json({ error: err.message, stack: err.stack });
  }
});

app.listen(PORT, () => console.log(`Server running at http://0.0.0.0:${PORT}`));