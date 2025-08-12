// utils.js - Normalisierung & Mapping für QMP Siemens Produktcheck

function a2vUrl(a2v) {
  const id = (a2v || '').toString().trim();
  return `https://www.mymobase.com/de/p/${id}`;
}

function cleanNumberString(s) {
  if (s == null) return null;
  const str = String(s).replace(/\s+/g, '').replace(',', '.'); // 12,3 -> 12.3
  return str;
}

function toNumber(val) {
  if (val == null || val === '') return null;
  const s = cleanNumberString(val);
  const m = s ? s.match(/-?\d+(?:\.\d+)?/) : null;
  if (!m) return null;
  return parseFloat(m[0]);
}

// Gewicht: parse "0,162 kg" => { value: 0.162, unit: 'kg' }
function parseWeight(value) {
  if (!value && value !== 0) return { value: null, unit: '' };
  const s = String(value).toLowerCase().replace(',', '.').trim();
  const m = s.match(/-?\d+(?:\.\d+)?/);
  const num = m ? parseFloat(m[0]) : null;
  let unit = '';
  if (/mg\b/.test(s)) unit = 'mg';
  else if (/\bg\b/.test(s) && !/\bkg\b/.test(s)) unit = 'g';
  else if (/\bkg\b/.test(s)) unit = 'kg';
  else if (/\bt\b/.test(s)) unit = 't';
  return { value: num, unit };
}

// Für Vergleich: in kg umrechnen
function weightToKg(value, unit) {
  if (value == null) return null;
  const u = (unit || '').toLowerCase();
  if (u === 'mg') return value / 1e6;
  if (u === 'g') return value / 1000;
  if (u === 'kg' || u === '') return value;
  if (u === 't') return value * 1000;
  return value;
}

/**
 * Dimensions-Parser:
 * - akzeptiert "L×B×H", "LxBxH", "40X40X42", "30x20x10 mm", etc.
 * - Standard-Reihenfolge: Länge × Breite × Höhe (L×B×H)
 * - unterstützt auch Zylinder-Formate: "D×H", "DxH", "20x30 mm" (Durchmesser x Höhe)
 * - Ergebnis in mm (falls Einheiten erkennbar), sonst roh.
 */
function parseDimensionsToLBH(text) {
  if (!text) return { L: null, B: null, H: null };
  const raw = String(text).trim();

  // 1) Einheit in mm|cm|m erkennen, aber entfernen
  let scale = 1;
  let s = raw.toLowerCase()
             .replace(/[,;]/g, '.')          // Komma oder Semikolon → Punkt
             .replace(/\s+/g, '')            // Leerzeichen killen
             .replace(/[×xX*/]/g, 'x');      // *, ×, X, / → x

  if (/cm\b/.test(s)) { scale = 10;  s = s.replace(/cm\b/g, ''); }
  if (/(^|\D)m\b/.test(s)) { scale = 1000; s = s.replace(/m\b/g, ''); }

  // 2) Bis zu 3 Zahlen extrahieren - verbesserte Regex für verschiedene Formate
  const nums = (s.match(/-?\d+(?:\.\d+)?/g) || []).map(parseFloat);
  
  let L, B, H;
  
  if (nums.length === 2) {
    // Zylinder-Format: Durchmesser x Höhe
    // Durchmesser = Breite (B), Höhe = Höhe (H), Länge = null
    L = null;
    B = nums[0] != null ? Math.round(nums[0] * scale) : null;
    H = nums[1] != null ? Math.round(nums[1] * scale) : null;
    console.log(`Parsed cylinder dimensions: "${raw}" -> Durchmesser:${B}, Höhe:${H}`);
  } else if (nums.length === 3) {
    // Quader-Format: Länge x Breite x Höhe (Standard-Reihenfolge)
    [L, B, H] = nums.map(n => n != null ? Math.round(n * scale) : null);
    console.log(`Parsed cuboid dimensions: "${raw}" -> Länge:${L}, Breite:${B}, Höhe:${H}`);
  } else {
    // Unbekanntes Format
    L = B = H = null;
    console.log(`Unknown dimension format: "${raw}"`);
  }

  return { L, B, H };
}

/**
 * Artikelnummer normalisieren für exakte Vergleiche
 * Entfernt Leerzeichen, Bindestriche, Unterstriche und Schrägstriche
 */
function normPartNo(s) {
  if (!s) return '';
  return String(s).toUpperCase().replace(/[\s\-\/_]+/g, '');
}

/**
 * Exakte Gewichtsvergleiche ohne Toleranz
 * Vergleicht zwei Gewichtswerte in kg
 */
function compareWeightExact(weight1, weight2) {
  if (weight1 == null || weight2 == null) return false;
  const w1 = toNumber(weight1);
  const w2 = toNumber(weight2);
  if (w1 == null || w2 == null) return false;
  
  // Exakte Gleichheit ohne Toleranz
  return Math.abs(w1 - w2) < 1e-9;
}

/**
 * Exakte Dimensionsvergleiche ohne Toleranz
 * Vergleicht Länge, Breite und Höhe einzeln
 */
function compareDimensionsExact(dims1, dims2) {
  if (!dims1 || !dims2) return false;
  
  const L1 = toNumber(dims1.L);
  const B1 = toNumber(dims1.B);
  const H1 = toNumber(dims1.H);
  
  const L2 = toNumber(dims2.L);
  const B2 = toNumber(dims2.B);
  const H2 = toNumber(dims2.H);
  
  // Exakte Gleichheit für jede Dimension
  return L1 === L2 && B1 === B2 && H1 === H2;
}

/**
 * Materialklassifizierung zu Excel-Code mappen
 */
function mapMaterialClassificationToExcel(text) {
  if (!text) return '';
  const s = String(text).toLowerCase();
  const hasNicht = /nicht/.test(s);
  const hasSchweiss = /schwei|schweiß|schweiss/.test(s);
  const hasGuss = /guss/.test(s);
  const hasKlebe = /klebe/.test(s);
  const hasSchmiede = /schmiede/.test(s);
  const hasRelevant = /relev/.test(s);
  if (hasNicht && (hasSchweiss || hasGuss || hasKlebe || hasSchmiede) && hasRelevant) {
    return 'OHNE/N/N/N/N';
  }
  return '';
}

/**
 * N-Code normalisieren für exakte Vergleiche
 * Entfernt alle Leerzeichen und konvertiert zu Großbuchstaben
 */
function normalizeNCode(s) {
  if (!s) return '';
  return String(s).replace(/\s+/g,'').toUpperCase();
}

/**
 * Exakte Textvergleiche ohne Normalisierung
 * Vergleicht nur nach Trim, case-sensitive
 */
function compareTextExact(text1, text2) {
  if (text1 == null || text2 == null) return false;
  const t1 = String(text1).trim();
  const t2 = String(text2).trim();
  return t1 === t2;
}

module.exports = {
  a2vUrl,
  cleanNumberString,
  toNumber,
  parseWeight,
  weightToKg,
  parseDimensionsToLBH,
  normPartNo,
  compareWeightExact,
  compareDimensionsExact,
  mapMaterialClassificationToExcel,
  normalizeNCode,
  compareTextExact
};
