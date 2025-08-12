// utils.js - Normalisierung & Mapping (aktualisiert)

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
 * - Ergebnis in mm (falls Einheiten erkennbar), sonst roh.
 */
function parseDimensionsToLBH(text) {
  if (!text) return { L:null, B:null, H:null };
  const raw = String(text).trim();
  let s = raw.toLowerCase().replace(/[×x]/g, 'x').replace(',', '.').replace(/\s+/g, '');

  // Einheit erkennen (mm, cm, m)
  let scale = 1; // default mm
  if (/(^|\D)cm\b/.test(s) || s.endsWith('cm')) scale = 10;
  if (/(^|\D)m\b/.test(s) || s.endsWith('m')) scale = 1000;

  const nums = (s.match(/-?\d+(?:\.\d+)?/g) || []).map(parseFloat);
  const L = nums.length > 0 ? Math.round(nums[0] * scale) : null;
  const B = nums.length > 1 ? Math.round(nums[1] * scale) : null;
  const H = nums.length > 2 ? Math.round(nums[2] * scale) : null;
  return { L, B, H };
}

function normPartNo(s) {
  if (!s) return '';
  return String(s).toUpperCase().replace(/[\s\-\/_]+/g, '');
}

function withinToleranceKG(exKg, wbKg, tolPct) {
  if (exKg == null || wbKg == null) return false;
  const diff = Math.abs(exKg - wbKg);
  if (!tolPct || tolPct <= 0) return diff < 1e-9; // streng
  const tol = Math.abs(exKg) * (tolPct / 100);
  return diff <= tol;
}

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

// „OHNE/N  /N  /N/N “ -> "OHNE/N/N/N/N"
function normalizeNCode(s) {
  if (!s) return '';
  return String(s).replace(/\s+/g,'').toUpperCase();
}

module.exports = {
  a2vUrl,
  cleanNumberString,
  toNumber,
  parseWeight,
  weightToKg,
  parseDimensionsToLBH,
  normPartNo,
  withinToleranceKG,
  mapMaterialClassificationToExcel,
  normalizeNCode
};