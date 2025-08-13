// scraper.js - A2V-only scraper with structured JSON extraction and optional Playwright fallback.
// Playwright is lazy-required and can be disabled by setting DISABLE_PLAYWRIGHT=1.

const cheerio = require('cheerio');
const NAV_TIMEOUT_MS = Number(process.env.NAV_TIMEOUT_MS || 18000);
const DISABLE_PLAYWRIGHT = String(process.env.DISABLE_PLAYWRIGHT || '0') === '1';

function a2vUrl(a2v) {
  const id = String(a2v || '').trim();
  return `https://www.mymobase.com/de/p/${id}`;
}

function extractJsonInitialData(html) {
  // Verschiedene mögliche JSON-Formate versuchen
  const patterns = [
    /window\.initialData\[['"]product\/dataProduct['"]]\s*=\s*(\{[\s\S]*?\});\s*<\/script>/i,
    /window\.initialData\s*=\s*(\{[\s\S]*?\});\s*<\/script>/i,
    /initialData\[['"]product\/dataProduct['"]]\s*=\s*(\{[\s\S]*?\});/i,
    /initialData\s*=\s*(\{[\s\S]*?\});/i
  ];
  
  for (const pattern of patterns) {
    const m = html.match(pattern);
    if (m) {
      try { 
        const parsed = JSON.parse(m[1]);
        console.log(`Successfully extracted JSON for product data`);
        return parsed; 
      } catch (e) {
        console.log(`Failed to parse JSON with pattern: ${pattern}`);
      }
    }
  }
  return null;
}

function mapFromInitialData(obj, a2v, url) {
  try {
    const product = obj?.data?.product || {};
    const ts = product?.localizations?.technicalSpecifications || product?.technicalSpecifications || [];
    const tsMap = {};
    for (const item of ts) {
      if (!item || typeof item !== 'object') continue;
      const k = String(item.key || '').toLowerCase();
      tsMap[k] = item.value || '';
    }
    const pickTs = (...needles) => {
      for (const [k,v] of Object.entries(tsMap)) if (needles.every(n => k.includes(n))) return v;
      return null;
    };

    const weitere = pickTs('weitere','artikelnummer') || product.additionalMaterialNumbers || product.baseProductAdditionalMaterialNumbers || 'Nicht gefunden';
    let gewicht = pickTs('gewicht') || null;
    if (!gewicht && typeof product.weight === 'number') gewicht = `${product.weight.toString().replace('.', ',')} kg`;
    if (!gewicht) gewicht = 'Nicht gefunden';
    const materialklass = pickTs('materialklassifizierung') || product.materialClassification || 'Nicht gefunden';
    const name = product.name || 'Nicht gefunden';
    const code = product.code || a2v;
    
    // Abmessungen extrahieren - erweiterte Suche
    let abmessung = pickTs('abmess') || pickTs('dimension') || pickTs('größe') || pickTs('size') || pickTs('maße') || pickTs('measure') || pickTs('länge') || pickTs('breite') || pickTs('höhe') || pickTs('length') || pickTs('width') || pickTs('height') || pickTs('abmessungen') || pickTs('dimensions') || pickTs('abmessung') || 'Nicht gefunden';
    
    // Debug-Logging für Abmessungen
    if (abmessung !== 'Nicht gefunden') {
      console.log(`Found dimensions for ${a2v}: "${abmessung}"`);
    } else {
      console.log(`No dimensions found for ${a2v}. Available keys:`, Object.keys(tsMap).slice(0, 10)); // Nur erste 10 Keys anzeigen
    }
    
    // Werkstoff extrahieren
    let werkstoff = pickTs('werkstoff') || (pickTs('material') && !pickTs('material','klass')) || 'Nicht gefunden';

    // Fert./Prüfhinweis extrahieren
    let fertPruefhinweis = pickTs('fertigungs') || pickTs('prüf') || pickTs('fertigung') || pickTs('pruef') || pickTs('hinweis') || pickTs('note') || pickTs('bemerkung') || pickTs('remark') || 'Nicht gefunden';

    return {
      A2V: code,
      URL: url,
      Produkttitel: name,
      'Weitere Artikelnummer': weitere,
      Gewicht: gewicht,
      Abmessung: abmessung,
      Werkstoff: werkstoff,
      Materialklassifizierung: materialklass,
      FertPruefhinweis: fertPruefhinweis,
      Status: 'initialData JSON'
    };
  } catch { return null; }
}

class SiemensProductScraper {
  constructor() {
    this.cache = new Map();
    this.browser = null;
    this.context = null;
  }

  async _httpGet(url) {
    const resp = await fetch(url, {
      headers: {
        'User-Agent':
          'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36'
      }
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    return await resp.text();
  }

  _parseWithCheerio(url, html, a2v) {
    const $ = cheerio.load(html);
    const kv = {};
    $('table').each((_, t) => {
      $(t).find('tr').each((_, tr) => {
        const tds = $(tr).find('td,th');
        if (tds.length >= 2) {
          const k = $(tds[0]).text().trim().toLowerCase();
          const v = $(tds[1]).text().trim();
          if (k && v && !kv[k]) kv[k] = v;
        }
      });
    });
    $('dl').each((_, dl) => {
      const dts = $(dl).find('dt'); const dds = $(dl).find('dd');
      for (let i=0;i<Math.min(dts.length, dds.length); i++) {
        const k = $(dts[i]).text().trim().toLowerCase();
        const v = $(dds[i]).text().trim();
        if (k && v && !kv[k]) kv[k] = v;
      }
    });
    const pick = (needles) => {
      for (const k of Object.keys(kv)) if (needles.every(n => k.includes(n))) return kv[k];
      return null;
    };
    const title = ($('h1, .product-title').first().text() || $('title').first().text() || '').replace(' | MoBase','').trim();
    
    // Fert./Prüfhinweis extrahieren
    const fertPruefhinweis = pick(['fertigungs']) || pick(['prüf']) || pick(['fertigung']) || pick(['pruef']) || pick(['hinweis']) || pick(['note']) || pick(['bemerkung']) || pick(['remark']) || 'Nicht gefunden';
    
    return {
      A2V: a2v,
      URL: url,
      Produkttitel: title || 'Nicht gefunden',
      'Weitere Artikelnummer':
        pick(['weitere','artikelnummer']) ||
        pick(['additional','material','number']) ||
        pick(['part','number']) || 'Nicht gefunden',
      Gewicht:  pick(['gewicht']) || pick(['weight']) || 'Nicht gefunden',
      Abmessung: pick(['abmess']) || pick(['dimension']) || pick(['größe']) || pick(['size']) || 'Nicht gefunden',
      Werkstoff: (pick(['werkstoff']) || (pick(['material']) && !pick(['material','klass']))) || 'Nicht gefunden',
      Materialklassifizierung: pick(['material','klass']) || pick(['material','class']) || 'Nicht gefunden',
      FertPruefhinweis: fertPruefhinweis,
      Status: 'HTTP-Parser'
    };
  }

  async httpScrapeA2V(a2v) {
    const url = a2vUrl(a2v);
    const html = await this._httpGet(url);
    const initObj = extractJsonInitialData(html);
    if (initObj) {
      const mapped = mapFromInitialData(initObj, a2v, url);
      if (mapped) return mapped;
    }
    return this._parseWithCheerio(url, html, a2v);
  }

  async _getChromium() {
    if (DISABLE_PLAYWRIGHT) return null;
    try {
      const { chromium } = require('playwright');
      return chromium;
    } catch {
      return null;
    }
  }

  async _initPlaywright() {
    const chromium = await this._getChromium();
    if (!chromium) return false;
    if (!this.browser) {
      this.browser = await chromium.launch({
        headless: true,
        args: ['--no-sandbox','--disable-setuid-sandbox','--disable-dev-shm-usage']
      });
    }
    if (!this.context) {
      this.context = await this.browser.newContext({
        bypassCSP: true,
        viewport: { width: 1200, height: 900 },
        userAgent:
          'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36'
      });
      await this.context.route('**/*', (route) => {
        const type = route.request().resourceType();
        if (['image','stylesheet','font','media','websocket','other'].includes(type)) return route.abort();
        route.continue();
      });
    }
    return true;
  }

  async pwScrapeA2V(a2v) {
    const ok = await this._initPlaywright();
    if (!ok) throw new Error('Playwright nicht verfügbar');
    const url = a2vUrl(a2v);
    const page = await this.context.newPage();
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: NAV_TIMEOUT_MS });
    const initJson = await page.evaluate(() => {
      const re = /window\.initialData\[['"]product\/dataProduct['"]]\s*=\s*(\{[\s\S]*?\});/i;
      for (const s of document.querySelectorAll('script')) {
        const t = s.textContent || '';
        const m = t.match(re);
        if (m) return m[1];
      }
      return null;
    });
    if (initJson) {
      try {
        const obj = JSON.parse(initJson);
        const mapped = mapFromInitialData(obj, a2v, url);
        if (mapped) { await page.close(); return mapped; }
      } catch {}
    }
    const kv = await page.evaluate(() => {
      const data = {};
      const add = (k, v) => { if (!k || !v) return; k=k.trim().toLowerCase(); v=v.trim(); if (!data[k]) data[k] = v; };
      document.querySelectorAll('table').forEach(t => {
        t.querySelectorAll('tr').forEach(tr => {
          const tds = tr.querySelectorAll('td,th');
          if (tds.length >= 2) add(tds[0].textContent, tds[1].textContent);
        });
      });
      document.querySelectorAll('dl').forEach(dl => {
        const dts = dl.querySelectorAll('dt'); const dds = dl.querySelectorAll('dd');
        for (let i=0;i<Math.min(dts.length, dds.length); i++) add(dts[i].textContent, dds[i].textContent);
      });
      return data;
    });
    const pick = (needles) => {
      for (const k of Object.keys(kv)) {
        const low = k.toLowerCase();
        if (needles.every(n => low.includes(n))) return kv[k];
      }
      return null;
    };
    const title = (await page.locator('h1, .product-title').first().textContent().catch(()=>''))?.replace(' | MoBase','').trim();
    
    // Fert./Prüfhinweis extrahieren
    const fertPruefhinweis = pick(['fertigungs']) || pick(['prüf']) || pick(['fertigung']) || pick(['pruef']) || pick(['hinweis']) || pick(['note']) || pick(['bemerkung']) || pick(['remark']) || 'Nicht gefunden';
    
    await page.close();
    return {
      A2V: a2v,
      URL: url,
      Produkttitel: title || 'Nicht gefunden',
      'Weitere Artikelnummer':
        pick(['weitere','artikelnummer']) || pick(['additional','material','number']) || pick(['part','number']) || 'Nicht gefunden',
      Gewicht:  pick(['gewicht']) || pick(['weight']) || 'Nicht gefunden',
      Abmessung: pick(['abmess']) || pick(['dimension']) || pick(['größe']) || pick(['size']) || 'Nicht gefunden',
      Werkstoff: (pick(['werkstoff']) || (pick(['material']) && !pick(['material','klass']))) || 'Nicht gefunden',
      Materialklassifizierung: pick(['material','klass']) || pick(['material','class']) || 'Nicht gefunden',
      FertPruefhinweis: fertPruefhinweis,
      Status: 'Playwright'
    };
  }

  async scrapeOne(a2v) {
    const key = String(a2v || '').trim().toUpperCase();
    if (!key.startsWith('A2V')) throw new Error('Nur A2V-Nummern sind erlaubt.');
    if (this.cache.has(key)) return this.cache.get(key);
    let out;
    try {
      out = await this.httpScrapeA2V(key);
    } catch (e) {
      try { out = await this.pwScrapeA2V(key); }
      catch (err) {
        out = { A2V: key, URL: a2vUrl(key), Produkttitel:'Nicht gefunden', 'Weitere Artikelnummer':'Nicht gefunden', Abmessung:'Nicht gefunden', Gewicht:'Nicht gefunden', Werkstoff:'Nicht gefunden', Materialklassifizierung:'Nicht gefunden', FertPruefhinweis:'Nicht gefunden', Status:'Fehler: '+err.message };
      }
    }
    this.cache.set(key, out);
    return out;
  }

  async scrapeMany(list, concurrency = 6) {
    const unique = Array.from(new Set(list.filter(Boolean).map(x => String(x).trim().toUpperCase())));
    const results = new Map();
    let i = 0;
    const worker = async () => {
      while (i < unique.length) {
        const idx = i++;
        const id = unique[idx];
        const r = await this.scrapeOne(id);
        results.set(id, r);
      }
    };
    await Promise.all(Array.from({ length: Math.max(1, concurrency) }, () => worker()));
    return results;
  }

  async close() {
    if (this.context) { await this.context.close(); this.context = null; }
    if (this.browser)  { await this.browser.close(); this.browser = null; }
  }
}

module.exports = { SiemensProductScraper, a2vUrl };