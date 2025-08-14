
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

// Vorlage-Layout
const HEADER_ROW = 3;      // Spalten-Namen
const SUBHEADER_ROW = 4;   // DB-Wert / Web-Wert
const FIRST_DATA_ROW = 5;  // erste Datenzeile

app.use(helmet({ contentSecurityPolicy: false }));
app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(__dirname));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// ===== Helpers
function colLetterToIndex(letter) {
  let n = 0; for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64); return n;
}
function colIndexToLetter(idx) {
  let s = ''; let n = idx; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s;
}
function addr(c, r) { return `${colIndexToLetter(c)}${r}`; }
function getStr(val) { if (val == null) return ''; if (typeof val === 'object' && val.text != null) return String(val.text); return String(val); }
function fillColor(ws, col, row, color) {
  if (!color) return; const map = { green: 'FFD5F4E6', red: 'FFFDEAEA', orange: 'FFFFF3CD' };
  ws.getCell(addr(col, row)).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: map[color] || map.green } };
}
function eqText(a, b) { const A = getStr(a).trim().toLowerCase().replace(/\s+/g,' '); const B = getStr(b).trim().toLowerCase().replace(/\s+/g,' '); return !!A && !!B && A===B; }
function eqPart(a,b){ return normPartNo(a)===normPartNo(b); }
function eqN(a,b){ return normalizeNCode(a)===normalizeNCode(b); }
function eqWeight(exVal, exUnit, webVal){ const { value:wv, unit:wu } = parseWeight(webVal); if (wv==null) return false; const exNum = toNumber(exVal); if (exNum==null) return false; const exU=(getStr(exUnit)||'').toLowerCase(); const a=weightToKg(exNum, exU); const b=weightToKg(wv, wu||exU||'kg'); return a!=null && b!=null && Math.abs(a-b)<1e-9; }

// Zielspalten gemäß Vorlage (fix)
const TGT = {
  MAT_DB:  colLetterToIndex('C'), MAT_WEB: colLetterToIndex('D'),
  PART_DB: colLetterToIndex('F'), PART_WEB: colLetterToIndex('G'),
  PRF_DB:  colLetterToIndex('P'), PRF_WEB: colLetterToIndex('Q'),
  WERK_DB: colLetterToIndex('S'), WERK_WEB: colLetterToIndex('T'),
  NET_DB:  colLetterToIndex('W'), NET_WEB: colLetterToIndex('X'), NET_UNIT: colLetterToIndex('Y'),
  L_DB:    colLetterToIndex('Z'), L_WEB:   colLetterToIndex('AA'),
  B_DB:    colLetterToIndex('AB'),B_WEB:   colLetterToIndex('AC'),
  H_DB:    colLetterToIndex('AD'),H_WEB:   colLetterToIndex('AE'),
  DIM_UNIT:colLetterToIndex('AF'),
  A2V:     colLetterToIndex('AH')
};

function findColumnByIncludes(ws, headerRow, ...needles) {
  const row = ws.getRow(headerRow);
  const lastCol = Math.max(ws.actualColumnCount||0, row.actualCellCount||0, row.cellCount||0, 60);
  const nls = needles.map(n=>n.toLowerCase());
  for (let c=1;c<=lastCol;c++){
    const name = getStr(row.getCell(c).value).toLowerCase();
    if (!name) continue; if (nls.every(n=>name.includes(n))) return c;
  }
  return null;
}

function ensureRow4OnlySubheaders(ws, pairs){
  const usedCols = Math.max(ws.actualColumnCount||0, ws.getRow(HEADER_ROW).cellCount||0, 60);
  // 1) Daten in Row4 nach Row5 verschieben, falls vorhanden (ausser DB/Web Titel)
  const allowed = new Set();
  for (const p of pairs){ allowed.add(p.db); allowed.add(p.web); }
  const row4 = ws.getRow(SUBHEADER_ROW);
  let hasDataInRow4 = false;
  const r4Values = [];
  for (let c=1;c<=usedCols;c++){
    const v = row4.getCell(c).value;
    const isAllowedTitle = (v==='DB-Wert' && allowed.has(c)) || (v==='Web-Wert' && allowed.has(c));
    if (v!=null && v!=='' && !isAllowedTitle){ hasDataInRow4 = true; }
    r4Values[c] = v;
  }
  if (hasDataInRow4){
    // Insert a new row 5 to make space, then copy row4 data to row5
    ws.spliceRows(FIRST_DATA_ROW, 0, []);
    const r5 = ws.getRow(FIRST_DATA_ROW);
    for (let c=1;c<=usedCols;c++) r5.getCell(c).value = r4Values[c];
  }
  // 2) Zeile 4 vollständig leeren
  for (let c=1;c<=usedCols;c++) row4.getCell(c).value = null;
  // 3) Nur DB/Web schreiben
  for (const p of pairs){ ws.getRow(SUBHEADER_ROW).getCell(p.db).value = 'DB-Wert'; ws.getRow(SUBHEADER_ROW).getCell(p.web).value = 'Web-Wert'; }
}

app.get('/', (req,res)=> res.sendFile(path.join(__dirname,'index.html')));
app.get('/api/health', (req,res)=> res.json({ ok:true, time:new Date().toISOString() }));

app.post('/api/process-excel', upload.single('file'), async (req,res)=>{
  try{
    if(!req.file) return res.status(400).json({ error:'Bitte Excel-Datei hochladen (file).' });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(req.file.buffer);
    const scraper = new SiemensProductScraper();

    for (const ws of wb.worksheets){
      // 0) Paare für Subheader gemäß Vorlage
      const pairs = [
        { db:TGT.MAT_DB,  web:TGT.MAT_WEB },
        { db:TGT.PART_DB, web:TGT.PART_WEB },
        { db:TGT.PRF_DB,  web:TGT.PRF_WEB },
        { db:TGT.WERK_DB, web:TGT.WERK_WEB },
        { db:TGT.NET_DB,  web:TGT.NET_WEB },
        { db:TGT.L_DB,    web:TGT.L_WEB },
        { db:TGT.B_DB,    web:TGT.B_WEB },
        { db:TGT.H_DB,    web:TGT.H_WEB }
      ];

      // 1) Row4 reinigen + ggf. Row4-Daten nach Row5 verschieben, dann Subheader setzen
      ensureRow4OnlySubheaders(ws, pairs);

      // 2) Ursprungs-DB-Spalten erkennen (nur zum Lesen) — tolerant nach Text
      const SRC = {
        MAT_DB:  findColumnByIncludes(ws, HEADER_ROW, 'material', 'kurz') || TGT.MAT_DB,
        PART_DB: findColumnByIncludes(ws, HEADER_ROW, 'artikel') || TGT.PART_DB,
        PRF_DB:  findColumnByIncludes(ws, HEADER_ROW, 'fert', 'prüf') || findColumnByIncludes(ws, HEADER_ROW, 'fert','pruef') || TGT.PRF_DB,
        WERK_DB: findColumnByIncludes(ws, HEADER_ROW, 'werkstoff') || TGT.WERK_DB,
        NET_DB:  findColumnByIncludes(ws, HEADER_ROW, 'netto') || TGT.NET_DB,
        NET_UNIT:findColumnByIncludes(ws, HEADER_ROW, 'gewicht', 'einheit') || TGT.NET_UNIT,
        L_DB:    findColumnByIncludes(ws, HEADER_ROW, 'länge') || findColumnByIncludes(ws, HEADER_ROW, 'laenge') || TGT.L_DB,
        B_DB:    findColumnByIncludes(ws, HEADER_ROW, 'breite') || TGT.B_DB,
        H_DB:    findColumnByIncludes(ws, HEADER_ROW, 'höhe') || findColumnByIncludes(ws, HEADER_ROW, 'hoehe') || TGT.H_DB,
        DIM_UNIT:findColumnByIncludes(ws, HEADER_ROW, 'einheit', 'abmaß') || findColumnByIncludes(ws, HEADER_ROW, 'einheit', 'abmass') || TGT.DIM_UNIT,
        A2V:     findColumnByIncludes(ws, HEADER_ROW, 'a2v') || findColumnByIncludes(ws, HEADER_ROW, 'materialnummer') || TGT.A2V
      };

      // 3) Zeilen mit A2V sammeln (ab FIRST_DATA_ROW)
      const last = ws.lastRow ? ws.lastRow.number : Math.max(ws.rowCount||0, ws.actualRowCount||0);
      const rows = []; const ids = [];
      for (let r = FIRST_DATA_ROW; r <= last; r++){
        const v = getStr(ws.getRow(r).getCell(SRC.A2V).value).trim().toUpperCase();
        if (v && v.startsWith('A2V')){ rows.push(r); ids.push(v); }
      }
      if (!ids.length) continue;

      // 4) Scrapen gebündelt
      const results = await scraper.scrapeMany(ids, SCRAPE_CONCURRENCY); // Map<A2V, obj>

      // 5) Schreiben in feste Zielspalten (DB-Spalten NICHT überschreiben)
      for (let i=0; i<rows.length; i++){
        const r = rows[i];
        const a2v = ids[i];
        const web = results.get(a2v) || {};

        const db = {
          mat: ws.getRow(r).getCell(SRC.MAT_DB).value,
          part:ws.getRow(r).getCell(SRC.PART_DB).value,
          prf: ws.getRow(r).getCell(SRC.PRF_DB).value,
          werk:ws.getRow(r).getCell(SRC.WERK_DB).value,
          net: ws.getRow(r).getCell(SRC.NET_DB).value,
          netU: SRC.NET_UNIT ? ws.getRow(r).getCell(SRC.NET_UNIT).value : '',
          L:   ws.getRow(r).getCell(SRC.L_DB).value,
          B:   ws.getRow(r).getCell(SRC.B_DB).value,
          H:   ws.getRow(r).getCell(SRC.H_DB).value
        };

        // Materialkurztext → D
        {
          const c = TGT.MAT_WEB; const webVal = (web.Produkttitel && web.Produkttitel!=='Nicht gefunden') ? web.Produkttitel : '';
          if (webVal) ws.getRow(r).getCell(c).value = webVal;
          fillColor(ws, c, r, webVal ? (eqText(db.mat, webVal) ? 'green' : 'red') : 'orange');
        }
        // Hersteller-Artikelnummer → G (Fallback A2V)
        {
          const c = TGT.PART_WEB; let webVal = (web['Weitere Artikelnummer'] && web['Weitere Artikelnummer']!=='Nicht gefunden') ? web['Weitere Artikelnummer'] : '';
          if (!webVal && a2v) webVal = a2v;
          if (webVal) ws.getRow(r).getCell(c).value = webVal;
          const excelVal = db.part || a2v; fillColor(ws, c, r, webVal ? (eqPart(excelVal, webVal) ? 'green' : 'red') : 'orange');
        }
        // Fertigung/Prüfhinweis → Q
        {
          const c = TGT.PRF_WEB; const code = normalizeNCode(mapMaterialClassificationToExcel(web.Materialklassifizierung || ''));
          if (code) ws.getRow(r).getCell(c).value = code;
          fillColor(ws, c, r, code ? (eqN(db.prf, code) ? 'green' : 'red') : 'orange');
        }
        // Werkstoff → T
        {
          const c = TGT.WERK_WEB; const webVal = (web.Werkstoff && web.Werkstoff!=='Nicht gefunden') ? web.Werkstoff : '';
          if (webVal) ws.getRow(r).getCell(c).value = webVal;
          fillColor(ws, c, r, webVal ? (eqText(db.werk, webVal) ? 'green' : 'red') : 'orange');
        }
        // Nettogewicht → X (+ ggf. Einheit Y)
        {
          const c = TGT.NET_WEB; const webVal = (web.Gewicht && web.Gewicht!=='Nicht gefunden') ? web.Gewicht : '';
          if (webVal){ const { value, unit } = parseWeight(webVal); ws.getRow(r).getCell(c).value = (value!=null) ? value : webVal; if (unit && TGT.NET_UNIT){ const uCell = ws.getRow(r).getCell(TGT.NET_UNIT); if (!uCell.value) uCell.value = unit; } fillColor(ws, c, r, eqWeight(db.net, db.netU, webVal) ? 'green' : 'red'); } else { fillColor(ws, c, r, 'orange'); }
        }
        // Abmessungen L/B/H → AA/AC/AE (+ Einheit AF optional)
        {
          const dims = (web.Abmessung && web.Abmessung!=='Nicht gefunden') ? parseDimensionsToLBH(web.Abmessung) : { L:null,B:null,H:null };
          // L
          { const c=TGT.L_WEB; if (dims.L!=null) ws.getRow(r).getCell(c).value = dims.L; if (toNumber(db.L)!=null && dims.L!=null) fillColor(ws, c, r, toNumber(db.L)===dims.L ? 'green':'red'); else if (dims.L==null) fillColor(ws, c, r, 'orange'); }
          // B
          { const c=TGT.B_WEB; if (dims.B!=null) ws.getRow(r).getCell(c).value = dims.B; if (toNumber(db.B)!=null && dims.B!=null) fillColor(ws, c, r, toNumber(db.B)===dims.B ? 'green':'red'); else if (dims.B==null) fillColor(ws, c, r, 'orange'); }
          // H
          { const c=TGT.H_WEB; if (dims.H!=null) ws.getRow(r).getCell(c).value = dims.H; if (toNumber(db.H)!=null && dims.H!=null) fillColor(ws, c, r, toNumber(db.H)===dims.H ? 'green':'red'); else if (dims.H==null) fillColor(ws, c, r, 'orange'); }
          // Einheit Abmaße
          if (TGT.DIM_UNIT){ const u = ws.getRow(r).getCell(TGT.DIM_UNIT); if (!u.value) u.value = 'MM'; }
        }
      }
    }

    const out = await new ExcelJS.Workbook().xlsx.writeBuffer?.call({}) // ensure function presence
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition','attachment; filename="DB_Produktvergleich_verarbeitet.xlsx"');
    res.send(Buffer.from(buf));
  } catch (err) {
    console.error('PROCESS ERROR:', err && err.stack || err);
    res.status(500).json({ error: String(err && err.message || err) });
  }
});

app.listen(PORT, ()=> console.log(`Server running at http://0.0.0.0:${PORT}`));
