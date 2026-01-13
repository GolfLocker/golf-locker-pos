/** ================== CONFIG ================== **/
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();

const TABS = ['Clubs','Sets','Tassen',"Trolley's",'Overig','Diensten']; // voorraad-tabs

const COL  = {
  sku:        1,
  desc:       2,
  buy:        4,
  party:      5,
  expected:   6,
  sale:       7,
  saleDate:   8,
  expMargin:  9,
  backupExp: 12,
  channel:   15   // kolom O
};

const CACHE_MIN     = 30;   // winkelmandje 30 min per gebruiker
const INDEX_TTL_MIN = 120;    // SKU-index 120 min cache

// Waar 80mm-bon PDF's worden opgeslagen
const TICKET80_FOLDER_ID = '1uI70DxaLW_RaYxpC2FIcEdhtIVWG6nO1';

// Branding gedeeld door front-end, PDF en 80mm
const BRAND = {
  name:       'Golf Locker',
  line1:      'Dorpstraat 38',
  line2:      '3981 EB, Bunnik',
  phone:      '06 3839 0722',
  email:      'info@golf-locker.nl',
  vat:        'NL861782495B01',
  kvk:        '80742300',
  extra:      'Btw inbegrepen',
  logoUrl:    'https://shop.golf-locker.nl/wp-content/uploads/2024/04/Golf-Locker-Logo-1-scaled.png',
  webshopUrl: 'https://shop.golf-locker.nl'
};

// Log-sheets
const LOG = {
  headSheet: 'Sales',
  lineSheet: 'Sales_Lines'
};


/** ================== HELPER: 80mm TICKET HTML ================== **/

/**
 * Bouwt de HTML voor de 80mm-bon.
 * Wordt gebruikt door:
 *  - _serveTicket_ (live herprint via ?file=ticket&no=..)
 *  - apiBookAndReceipt (voor PDF-archief in Drive)
 */
function _build80mmTicketHtml_(opts) {
  const {
    receiptNo,
    payMethod,
    customerEmail,
    total,
    subtotal = 0,
    discount = 0,
    dateString,
    items
  } = opts;


  const enc = encodeURIComponent;
  const fmt = n => Utilities.formatString(
    "€ %s",
    Number(n || 0).toFixed(2).replace('.', ',')
  );

  const qrUrl = `https://quickchart.io/qr?text=${enc(BRAND.webshopUrl || '')}&size=180&margin=1&format=png`;
  const code128Url = `https://bwipjs-api.metafloor.com/?bcid=code128&text=${enc(receiptNo || '')}&scale=3&height=12&includetext&textxalign=center`;

  // let op: </script> escapen als <\/script> in template literal
  // let op: </script> escapen
  const autoPrintScript = `
    <script>
      window.onload = function() {
        try { window.print(); } catch(e){}
      };
    <\\/script>
  `;

  const rowsHtml = (items || []).map(it => {
    const skuShort = String(it.sku || '').slice(0, 6).padEnd(6, ' ');
    const descSafe = String(it.desc || '').replace(/</g, '&lt;');
    const price    = Number(it.price || 0);
    return `
      <tr>
        <td class="sku mono">${skuShort}</td>
        <td class="desc">${descSafe}</td>
        <td class="price">${fmt(price)}</td>
      </tr>`;
  }).join('');

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Bon ${receiptNo || ''}</title>
  <style>
    @page { size: 80mm auto; margin: 4mm 4mm; }
    * { box-sizing:border-box; }
    body {
      font:11px/1.25 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
      color:#000;
      margin:0;
    }
    .ticket { width:72mm; max-width:72mm; margin:0 auto; }
    .center { text-align:center; }
    .right { text-align:right; }
    .muted { color:#555; }
    .logo { display:block; width:54mm; margin:0 auto 8px; filter:grayscale(100%); }
    h1 { font-size:14px; margin:0 0 6px; text-align:center; }
    .small { font-size:10px; }
    .mono { font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; }
    .row { display:flex; justify-content:space-between; }
    hr { border:0; border-top:1px dashed #000; margin:6px 0; }
    table { width:100%; border-collapse:collapse; }
    th,td { padding:2px 0; vertical-align:top; }
    th { text-align:left; font-weight:700; }
    .sku   { width:15mm; font-size:9px; }
    .desc  { width:40mm; padding-right:1,5mm; }
    .price { width:18mm; text-align:right; }
    .total { font-size:13px; font-weight:700; }
    .qr { width:36mm; margin:6px auto 0; }
    .barcode { width:60mm; margin:6px auto 0; display:block; }
    @media print { .noprint { display:none!important; } }
  </style>
</head>
<body>
  <div class="ticket">
    ${BRAND.logoUrl ? `<img class="logo" src="${BRAND.logoUrl}" alt="logo">` : ''}
    <h1>${BRAND.name || ''}</h1>
    <div class="center small muted">${BRAND.line1 || ''} • ${BRAND.line2 || ''}</div>
    <div class="center small muted">${BRAND.phone || ''} • ${BRAND.email || ''}</div>
    <div class="center small muted">BTW: ${BRAND.vat || '-'} • KvK: ${BRAND.kvk || '-'}</div>

    <hr>
    <div class="row small">
      <div>Betaalwijze: ${payMethod || '-'}</div>
      <div class="right">${dateString || ''}</div>
    </div>
    ${customerEmail ? `<div class="small">Klant: ${customerEmail}</div>` : ''}
    <hr>

    <table>
      <thead>
        <tr>
          <th class="sku">SKU</th>
          <th class="desc">Artikel</th>
          <th class="price">Prijs</th>
        </tr>
      </thead>
      <tbody>
        ${rowsHtml}
      </tbody>
    </table>

    <hr>
    <table>
      ${discount > 0 ? `
        <tr>
          <td class="right" colspan="5">Subtotaal: ${fmt(subtotal)}</td>
        </tr>
        <tr>
          <td class="right" colspan="5"><b>Korting: - ${fmt(discount)}</b></td>
        </tr>
      ` : ``}
      <tr>
        <td class="right total" colspan="5">Totaal: ${fmt(total)}</td>
      </tr>
    </table>

    ${BRAND.extra ? `<div class="right small muted" style="margin-top:3px">${BRAND.extra}</div>` : ''}

    <div class="center">
      <img class="qr" src="${qrUrl}" alt="qr">
    </div>
    <div class="center small muted" style="margin-top:4px">
      ${BRAND.webshopUrl || ''}
    </div>

    <img class="barcode" src="${code128Url}" alt="Bon ${receiptNo || ''}">
    <div class="center small muted" style="margin-top:6px">Bedankt voor uw aankoop!<br>Ruilen binnen 14 dagen met deze kassabon.</div>

    <div class="noprint center" style="margin-top:8px">
      <button onclick="window.print()">Print</button>
    </div>
  </div>

  ${autoPrintScript}
</body>
</html>`;
}


/** ================== WEB ENTRY (PWA + TICKET) ================== **/

function doGet(e) {
  const p = (e && e.parameter) || {};
  const f = p.file || '';

  if (f === 'manifest') return _serveManifest_();
  if (f === 'sw')       return _serveServiceWorker_();
  if (f === 'ticket')   return _serveTicket_(e);

  // 🔥 NIEUWE ROUTE → retourbon
  if (f === 'returnticket') {
    const no = p.no || '';
    return HtmlService.createHtmlOutput(buildReturnTicket80mm(no))
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Default → laad POS UI
  return HtmlService.createHtmlOutputFromFile('app')
    .setTitle('Golf Locker POS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function _serveManifest_() {
  const manifest = {
    name: "Golf Locker POS",
    short_name: "GL POS",
    description: "Snel afrekenen, bon en e-mail — ook als PWA.",
    start_url: "./",
    scope: "./",
    display: "standalone",
    background_color: "#ffffff",
    theme_color: "#297900",
    icons: [
      {
        src: BRAND.logoUrl,
        sizes: "192x192",
        type: "image/png",
        purpose: "any"
      },
      {
        src: BRAND.logoUrl,
        sizes: "512x512",
        type: "image/png",
        purpose: "any"
      }
    ]
  };

  return ContentService.createTextOutput(JSON.stringify(manifest))
    .setMimeType(ContentService.MimeType.JSON);
}

function _serveServiceWorker_() {
  const sw = [
    "const NAME = 'gl-pos-v3';",
    "const STATIC = [",
    "  self.registration.scope,",
    "  self.registration.scope + '?file=manifest',",
    "  'https://unpkg.com/@zxing/library@0.20.0',",
    `  '${BRAND.logoUrl}'`,
    "];",
    "",
    "self.addEventListener('install', e => {",
    "  e.waitUntil(",
    "    caches.open(NAME)",
    "      .then(c => c.addAll(STATIC))",
    "      .then(() => self.skipWaiting())",
    "  );",
    "});",
    "",
    "self.addEventListener('activate', e => {",
    "  e.waitUntil(",
    "    caches.keys().then(keys =>",
    "      Promise.all(keys.map(k => k === NAME ? null : caches.delete(k)))",
    "    )",
    "  );",
    "});",
    "",
    "self.addEventListener('fetch', e => {",
    "  const req = e.request;",
    "  const accept = req.headers.get('accept') || '';",
    "  const isHTML = accept.includes('text/html');",
    "  if (isHTML) {",
    "    e.respondWith(",
    "      fetch(req).then(res => {",
    "        const copy = res.clone();",
    "        caches.open(NAME).then(c => c.put(req, copy));",
    "        return res;",
    "      }).catch(() => caches.match(req))",
    "    );",
    "  } else {",
    "    e.respondWith(",
    "      caches.match(req).then(hit => hit ||",
    "        fetch(req).then(res => {",
    "          const copy = res.clone();",
    "          caches.open(NAME).then(c => c.put(req, copy));",
    "          return res;",
    "        })",
    "      )",
    "    );",
    "  }",
    "});"
  ].join('\n');

  return ContentService.createTextOutput(sw)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * HTML endpoint voor 80mm-ticket (?file=ticket&no=...)
 * Leest uit Sales + Sales_Lines en gebruikt dezelfde layout-helper.
 */
function _serveTicket_(e) {
  const no = (e && e.parameter && e.parameter.no) || '';
  if (!no) {
    return HtmlService.createHtmlOutput('Bonnummer ontbreekt');
  }

  const ss    = SpreadsheetApp.getActive();
  const head  = ss.getSheetByName(LOG.headSheet);
  const lines = ss.getSheetByName(LOG.lineSheet);

  if (!head || !lines) {
    return HtmlService.createHtmlOutput('Log-sheets niet gevonden');
  }

  // Kopregel zoeken
  const lastHead = head.getLastRow();
  const headVals = lastHead > 1
    ? head.getRange(2, 1, lastHead - 1, head.getLastColumn()).getValues()
    : [];

  const hRow = headVals.find(r => String(r[0]) === String(no));
  if (!hRow) {
    return HtmlService.createHtmlOutput('Bon niet gevonden: ' + no);
  }

  const receiptNo  = String(hRow[0]);
  const datum      = hRow[1];
  const payMethod  = String(hRow[2] || '');
  const total      = Number(hRow[3] || 0);
  const custEmail  = String(hRow[4] || '');
  const subtotal   = Number(hRow[8] || 0);   // 🔥 nieuw
  const discount   = Number(hRow[9] || 0);   // 🔥 nieuw

  const tz = Session.getScriptTimeZone() || 'Europe/Amsterdam';
  const dt = datum instanceof Date
    ? Utilities.formatDate(datum, tz, 'dd-MM-yyyy HH:mm')
    : String(datum || '');

  // Regels ophalen
  const lastLine = lines.getLastRow();
  const lineVals = lastLine > 1
    ? lines.getRange(2, 1, lastLine - 1, 7).getValues()
    : [];

  const items = lineVals
    .filter(r => String(r[0]) === receiptNo)
    .map(r => ({
      sku:   String(r[1] || ''),
      desc:  String(r[2] || ''),
      price: Number(r[3] || 0),
      qty:   Number(r[4] || 1),
      subt:  Number(r[5] || 0)
    }));

    const html = _build80mmTicketHtml_({
      receiptNo,
      payMethod,
      customerEmail: custEmail,
      total,
      subtotal,   
      discount,   
      dateString: dt,
      items
    });


  return HtmlService.createHtmlOutput(html)
    .setTitle('Ticket ' + receiptNo)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTML endpoint voor 80mm RETOUR-bon (?file=returnticket&no=RT-...)
 * Leest uit tab 'Retouren' en gebruikt dezelfde 80mm layout
 */
function buildReturnTicket80mm(returnNo) {
  if (!returnNo) {
    return HtmlService.createHtmlOutput('Retournummer ontbreekt');
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Retouren');
  if (!sh) {
    return HtmlService.createHtmlOutput('Tab "Retouren" niet gevonden');
  }

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) {
    return HtmlService.createHtmlOutput('Geen retourdata');
  }

  const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  // Filter op retournummer (kolom B)
  const items = rows.filter(r => String(r[1]) === String(returnNo));
  if (!items.length) {
    return HtmlService.createHtmlOutput('Retour niet gevonden: ' + returnNo);
  }

  // Neem eerste regel als "kop"
  const first = items[0];

  const receiptNo  = String(first[2] || ''); // originele bon
  const date       = first[0];
  const tz         = Session.getScriptTimeZone() || 'Europe/Amsterdam';
  const dateString = date instanceof Date
    ? Utilities.formatDate(date, tz, 'dd-MM-yyyy HH:mm')
    : String(date || '');

  const ticketItems = items.map(r => {
  const price = Number(r[5] || 0);

    return {
      sku:  String(r[3] || ''),
      desc: String(r[4] || ''),
      price: -Math.abs(price),
      subt:  -Math.abs(price)
    };
  });

  const total = ticketItems.reduce((s, it) => s + it.subt, 0);

  const html = _build80mmTicketHtml_({
    receiptNo: returnNo,
    payMethod: 'RETOUR',
    customerEmail: '',
    total,
    dateString,
    items: ticketItems
  });

  return html;
}



/** ================== KLEINE PING ================== **/

function apiPing() {
  return 'ok';
}


/** ================== CART STORAGE ================== **/

function _cartKey_() {
  const email = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || 'anon';
  return email + '::' + SpreadsheetApp.getActive().getId();
}

function getCart() {
  const cache = CacheService.getUserCache();
  const raw = cache.get(_cartKey_());
  return raw ? JSON.parse(raw) : [];
}

function saveCart(cart) {
  CacheService.getUserCache().put(_cartKey_(), JSON.stringify(cart), CACHE_MIN);
  return cart;
}

function clearCart() {
  CacheService.getUserCache().remove(_cartKey_());
  return [];
}

function apiGetCart()       { return getCart(); }
function apiClearCart()     { clearCart(); return true; }


/** ================== NIEUWE SUPER-SAFE SKU INDEX ================== **/

/**
 * Maak per sheet een aparte index.
 * Cache key per sheet, nooit te groot.
 */
function _indexKeyForSheet_(sheetName) {
  const ssId = SpreadsheetApp.getActive().getId();
  return `SKU_INDEX_V2::${ssId}::${sheetName}`;
}

/** Ophalen index voor een sheet */
function _getIndexForSheet_(sheetName) {
  const raw = CacheService.getScriptCache().get(_indexKeyForSheet_(sheetName));
  return raw ? JSON.parse(raw) : null;
}

/** Opslaan index voor een sheet (nooit te groot) */
function _saveIndexForSheet_(sheetName, data) {
  const key = _indexKeyForSheet_(sheetName);
  CacheService.getScriptCache().put(key, JSON.stringify(data), INDEX_TTL_MIN * 60);
}

/** Invalideer ALLE indexen voor ALLE voorraadtabs */
function invalidateSkuIndex_() {
  const ssId = SpreadsheetApp.getActive().getId();
  TABS.forEach(name => {
    CacheService.getScriptCache().remove(`SKU_INDEX_V2::${ssId}::${name}`);
  });
}

/**
 * Bouw index per sheet:
 * {
 *   "1513": [ { row: 12, free:true }, { row:18, free:false } ]
 * }
 */
function _buildIndexForSheet_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return null;

  const last = sh.getLastRow();
  if (last <= 1) {
    _saveIndexForSheet_(sheetName, {});
    return {};
  }

  const data = sh.getRange(2, 1, last - 1, 7).getValues();
  const index = {};

  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const sku = String(r[0] || '').trim();
    if (!sku) continue;

    const saleVal = r[COL.sale - 1];
    const free = (saleVal === '' || saleVal === null);

    if (!index[sku]) index[sku] = [];
    index[sku].push({
      row: i + 2,
      free,

      // 🔥 EXTRA DATA (zodat we GEEN sheet read meer nodig hebben)
      desc: r[1],                               // kolom B
      expected: Number(r[COL.expected - 1] || 0), // kolom F
      channel: r[COL.channel - 1] || ''          // kolom O
    });
  }

  _saveIndexForSheet_(sheetName, index);
  return index;
}

/** Haal index voor sheet op, bouw indien nodig */
function _ensureIndexForSheet_(sheetName) {
  let idx = _getIndexForSheet_(sheetName);
  if (!idx) idx = _buildIndexForSheet_(sheetName);
  return idx;
}

/**
 * NIEUWE snelle findBySku:
 * ✓ zoekt per sheet
 * ✓ first-free
 * ✓ nooit te groot
 */
function findBySku(sku, opts) {
  opts = opts || {};
  const key = String(sku).trim();
  const ss = SpreadsheetApp.getActive();

  for (let t = 0; t < TABS.length; t++) {
    const name = TABS[t];
    const index = _ensureIndexForSheet_(name);
    const list = index[key];
    if (!list || !list.length) continue;

    // free-first
    const freeOne = list.find(x => x.free);

    // ❌ GEEN vrije voorraad → al verkocht
    if (!freeOne) {
      // neem laatste verkochte regel
      const last = list[list.length - 1];

      const sh = SpreadsheetApp.getActive().getSheetByName(name);
      const saleDate = sh.getRange(last.row, COL.saleDate).getValue();
      const receipt  = sh.getRange(last.row, COL.receipt || 1).getValue(); // fallback
      const info = getLastSaleInfoBySku_(sku);

      throw new Error(
        'ALREADY_SOLD|' +
        (info?.receipt || '') +
        '|' +
        (info?.receiptUrl || '')
      );
    }

    // ✅ normale flow
    return {
      item: {
        sheetName: name,
        row: freeOne.row,
        desc: freeOne.desc,
        expected: freeOne.expected,
        channel: freeOne.channel,
        positions: list
      },
      positions: list
    };
  }

  return null; // Niet gevonden
}

/** Warm de hele index op */
function apiWarmIndex() {
  TABS.forEach(name => _buildIndexForSheet_(name));
  return true;
}


/** ================== CART API ================== **/

function apiAddFast(sku) {
  const rawSku = String(sku).trim();

  // ===========================
  // GENERATOR-SKU LOGICA (blijft)
  // ===========================
  if (GENERATOR_SKUS && GENERATOR_SKUS[rawSku]) {
    const generated = createGeneratedItem(rawSku);

    const line = {
      sku:       generated.sku,
      sheetName: generated.sheetName,
      row:       generated.row,
      desc:      generated.description,
      price:     generated.expected,
      qty:       1,
      party:     "",
      edited:    false,
      channel:   '',
      positions: [{ row: generated.row, free: true }]
    };

    // ❌ geen backend cart
    return { line };
  }

  // ===========================
  // NORMALE SKU
  // ===========================
  const res = findBySku(rawSku);
  if (!res) throw new Error('SKU niet gevonden: ' + rawSku);

  const { item, positions } = res;

  const line = {
    sku:       rawSku,
    sheetName: item.sheetName,
    row:       item.row,
    desc:      item.desc,
    price:     item.expected,
    qty:       1,
    party:     item.party,
    edited:    false,
    channel:   item.channel || '',
    positions: positions
  };

  // ❌ geen saveCart
  // ❌ geen total
  return { line };
}


function apiRemoveFast(sku) {
  const cart = getCart();
  const i = cart.findIndex(x => String(x.sku) === String(sku));
  if (i >= 0) cart.splice(i, 1);
  saveCart(cart);

  const total = cart.reduce((s, it) =>
    s + (Number(it.price) || 0) * (Number(it.qty) || 1), 0);

  return { ok: true, total };
}

function apiSetQtyFast(sku, qty) {
  const cart = getCart();
  const it = cart.find(x => String(x.sku) === String(sku));
  if (!it) throw new Error('Niet in mandje');

  it.qty = Math.max(1, Number(qty) || 1);
  saveCart(cart);

  const total = cart.reduce((s, i) =>
    s + (Number(i.price) || 0) * (Number(i.qty) || 1), 0);

  return { ok: true, total, line: it };
}

function apiSetPriceFast(sku, price) {
  const cart = getCart();
  const it = cart.find(x => String(x.sku) === String(sku));
  if (!it) throw new Error('Niet in mandje');

  it.price  = Math.max(0, Number(String(price).replace(',', '.')) || 0);
  it.edited = true;
  saveCart(cart);

  const total = cart.reduce((s, i) =>
    s + (Number(i.price) || 0) * (Number(i.qty) || 1), 0);

  return { ok: true, total, line: it };
}

function getLastSaleInfoBySku_(sku){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Sales_Lines');
  if (!sh) return null;

  const data = sh.getDataRange().getValues();
  const header = data.shift();

  const COL_SKU     = header.indexOf('sku');
  const COL_RECEIPT = header.indexOf('receipt_no');
  const RECEIPT_BASE_URL =
  'https://script.google.com/macros/s/AKfycby0SI0p9-oWKtk-JIh2CsLt_A4_NuYqsF8DlVX90k1g6pFjuXWeLkbVdRoc4ZAPA2yuZg/exec?file=ticket&no=';

  if (COL_SKU === -1 || COL_RECEIPT === -1) return null;

  for (let i = data.length - 1; i >= 0; i--){
    if (String(data[i][COL_SKU]).trim() === String(sku).trim()) {
      const receiptNo = data[i][COL_RECEIPT];
      if (!receiptNo) return null;

      return {
        receipt: receiptNo,
        receiptUrl: RECEIPT_BASE_URL + encodeURIComponent(receiptNo)
      };
    }
  }

  return null;
}

/** ================== REFRESH VANUIT SHEET ================== **/

function apiRefreshFromSheet() {
  const ss   = SpreadsheetApp.getActive();
  const cart = getCart();
  if (!cart.length) return { cart, total: 0, changed: 0 };

  let changed = 0;
  const bySheet = {};

  cart.forEach(it => {
    if (!bySheet[it.sheetName]) bySheet[it.sheetName] = [];
    bySheet[it.sheetName].push(it);
  });

  Object.keys(bySheet).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const items = bySheet[name];
    items.forEach(it => {
      if (it.edited) return; // handmatige prijs respecteren

      const desc     = sh.getRange(it.row, COL.desc).getValue();
      const expected = Number(sh.getRange(it.row, COL.expected).getValue()) || 0;

      if (it.desc !== desc || it.price !== expected) {
        it.desc  = desc;
        it.price = expected;
        changed++;
      }
    });
  });

  saveCart(cart);

  const total = cart.reduce((s, i) =>
    s + (Number(i.price) || 0) * (Number(i.qty) || 1), 0);

  return { cart, total, changed };
}


/** ================== MINI DATABASE (Sales / Sales_Lines) ================== **/

function ensureLogSheets_() {
  const ss = SpreadsheetApp.getActive();

  // HEAD-sheet
  let head = ss.getSheetByName(LOG.headSheet);
  if (!head) {
    head = ss.insertSheet(LOG.headSheet);
    head.getRange(1, 1, 1, 10).setValues([[
      'receipt_no', 'datum', 'betaalwijze', 'totaal',
      'klant_email', 'pdfUrl', 'mail_status', 'bon80Url', 'subtotal', 'discount'
    ]]);
    head.setFrozenRows(1);
  } else {
    const lastCol    = head.getLastColumn();
    const headerVals = head.getRange(1, 1, 1, lastCol).getValues()[0];

    // zorg dat pdfUrl, mail_status, bon80Url bestaan (zonder data te slopen)
    const needed = ['pdfUrl', 'mail_status', 'bon80Url'];
    needed.forEach(name => {
      if (!headerVals.includes(name)) {
        const newCol = head.getLastColumn() + 1;
        head.getRange(1, newCol).setValue(name);
      }
    });
  }

  // LINES-sheet
  let lines = ss.getSheetByName(LOG.lineSheet);
  if (!lines) {
    lines = ss.insertSheet(LOG.lineSheet);
    lines.getRange(1, 1, 1, 7).setValues([[
      'receipt_no','sku','omschrijving','prijs','aantal','subtotaal','partij'
    ]]);
    lines.setFrozenRows(1);
  }
}

function nextReceiptNo_() {
  const props = PropertiesService.getDocumentProperties();
  const d  = new Date();
  const yyyy = d.getFullYear();
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const dd   = String(d.getDate()).padStart(2, '0');
  const key  = `RCPT_SEQ_${yyyy}${mm}${dd}`;
  const cur  = Number(props.getProperty(key) || '0') + 1;
  props.setProperty(key, String(cur));
  const seq  = String(cur).padStart(3, '0');
  return `GL-${yyyy}${mm}${dd}-${seq}`;
}


/** ================== BON-STYLING SHEET 'Bon' ================== **/

function styleBon_(bon, receiptNo, payMethod, customerEmail, total, now, opts) {
  opts = opts || {};
  const subtotal = Number(opts.subtotal || 0);
  const discount = Number(opts.discount || 0);
  try {
    bon.clearFormats();
    bon.setHiddenGridlines(true);

    // Kolombreedtes
    bon.setColumnWidth(1, 140);
    bon.setColumnWidth(2, 180);
    bon.setColumnWidth(3, 90);
    bon.setColumnWidth(4, 70);
    bon.setColumnWidth(5, 110);

    /* ---------------------------
       LOGO (rij 1, kolom 1)
    ---------------------------- */
    try {
      bon.insertImage(BRAND.logoUrl, 1, 1)
        .setAnchorCell(bon.getRange("A1"))
        .setWidth(160)
        .setHeight(60);
    } catch (e) {
      Logger.log("⚠ Logo kon niet worden geladen: " + e);
      bon.getRange("A1").setValue(BRAND.name).setFontWeight("bold").setFontSize(18);
    }

    /* ---------------------------
       ADRESBLOK (linker zijde)
    ---------------------------- */
    bon.getRange("A3").setValue(BRAND.line1);
    bon.getRange("A4").setValue(BRAND.line2);
    bon.getRange("A5").setValue("Tel: " + BRAND.phone);
    bon.getRange("A6").setValue("E-mail: " + BRAND.email);
    bon.getRange("A7").setValue("KvK: " + BRAND.kvk + " | BTW: " + BRAND.vat);

    /* ---------------------------
       FACTUUR INFO (rechter zijde)
    ---------------------------- */
    bon.getRange("C1").setValue("FACTUUR")
      .setFontWeight("bold")
      .setFontSize(22);

    bon.getRange("C3").setValue("Factuurnummer:");
    bon.getRange("D3").setValue(receiptNo).setFontWeight("bold");

    bon.getRange("C4").setValue("Datum:");
    bon.getRange("D4").setValue(
      Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm")
    );

    bon.getRange("C5").setValue("Betaalwijze:");
    bon.getRange("D5").setValue(payMethod);

    bon.getRange("C6").setValue("Klant:");
    bon.getRange("D6").setValue(customerEmail || "-");

    /* ---------------------------
       TABEL HEADER
    ---------------------------- */
    const header = bon.getRange("A10:E10");
    header.setValues([["SKU","Omschrijving","Prijs","Aantal","Subtotaal"]]);
    header.setFontWeight("bold")
      .setBackground("#e6e6e6")
      .setHorizontalAlignment("center");

    // Zet body opmaak
    const last = bon.getLastRow();
    if (last > 10) {
      const body = bon.getRange(11, 1, last - 10, 5);
      body.setFontSize(11).setVerticalAlignment("middle");

      // prijs/subtotaal rechts uitlijnen
      bon.getRange(11, 3, last - 10, 1).setHorizontalAlignment("right");
      bon.getRange(11, 5, last - 10, 1).setHorizontalAlignment("right");

      // bedragformaten
      bon.getRange(11, 3, last - 10, 1).setNumberFormat("€ #,##0.00");
      bon.getRange(11, 5, last - 10, 1).setNumberFormat("€ #,##0.00");
    }

    /* ---------------------------
       TOTAALREGEL
    ---------------------------- */
    const totRow = bon.getLastRow() + 2;
    let row = totRow;

    // Cadeaubon (betaling met giftcard)
    if (opts.giftcard?.applied > 0) {
      bon.getRange(row, 4)
        .setValue('Cadeaubon:')
        .setFontWeight('bold');

      bon.getRange(row, 5)
        .setValue(-opts.giftcard.applied)
        .setNumberFormat('€ #,##0.00');

      row++;
    }

    // Kortingscode (alleen als er géén giftcard is gebruikt)
    if (discount > 0 && !opts.giftcard) {
      bon.getRange(row, 4)
        .setValue("Korting:")
        .setFontWeight("bold");

      bon.getRange(row, 5)
        .setValue(-discount)
        .setNumberFormat("€ #,##0.00")
        .setFontWeight("bold");

      row++;
    }

    // Totaal (netto)
    bon.getRange(row, 4)
      .setValue("Totaal:")
      .setFontWeight("bold")
      .setHorizontalAlignment("right");

    bon.getRange(row, 5)
      .setValue(total)
      .setNumberFormat("€ #,##0.00")
      .setFontWeight("bold")
      .setHorizontalAlignment("right");


    /* ---------------------------
       FOOTER
    ---------------------------- */
    bon.getRange(totRow + 3, 1)
      .setValue("Bedankt voor uw aankoop!")
      .setFontSize(12)
      .setFontWeight("bold");

  } catch (err) {
    Logger.log("❌ styleBon_ fout: " + err);
  }
}

/** ================== BOOK + RECEIPT (batch, locked) ================== **/

function apiBookAndReceipt(payMethod, customerEmail) {
  const lock = LockService.getDocumentLock();
  lock.tryLock(5000);

  try {
    ensureLogSheets_();

    const ss   = SpreadsheetApp.getActive();
    const cart = getCart();
    if (!cart.length) throw new Error('Mandje is leeg');

    const now   = new Date();
    const tz    = Session.getScriptTimeZone() || 'Europe/Amsterdam';
    const email = (customerEmail || '').trim();
    const receiptNo = nextReceiptNo_();

    let total = 0;

    // groepeer per sheet
    const bySheet = {};
    cart.forEach(it => {
      if (!bySheet[it.sheetName]) bySheet[it.sheetName] = [];
      bySheet[it.sheetName].push(it);
    });

    // schrijf sales in voorraad-tabbladen
    Object.keys(bySheet).forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) throw new Error('Tab niet gevonden: ' + name);

      const items = bySheet[name];
      

      // Verzamel exacte updates: row -> data
      const updates = [];

      items.forEach(it => {
        let need = it.qty;
        const positions = it.positions || []; // uit index

        const freeRows = positions.filter(p => p.free).map(p => p.row);

        if (freeRows.length < need) {
          throw new Error(`Niet genoeg voorraad voor ${it.sku} in ${name}`);
        }

        for (let i = 0; i < need; i++) {
          updates.push({
            row: freeRows[i],
            price: it.price
          });
          total += Number(it.price) || 0;
        }
      });

      // 🔥 schrijf ALLEEN de benodigde rijen
      updates.forEach(u => {
        sh.getRange(u.row, COL.sale).setValue(u.price);
        sh.getRange(u.row, COL.saleDate).setValue(now);
        sh.getRange(u.row, COL.expected).setValue(0);
        sh.getRange(u.row, COL.expMargin).setValue(0);
      });
    });

    // URL voor live 80mm-ticket (?file=ticket&no=...)
    const baseUrl   = ScriptApp.getService().getUrl();
    const ticketUrl = baseUrl + '?file=ticket&no=' + encodeURIComponent(receiptNo);

    // Loggen in Sales + Sales_Lines
    const head = ss.getSheetByName(LOG.headSheet);
    head.appendRow([
      receiptNo,
      now,
      String(payMethod || ''),
      total,
      String(email || ''),
      '',           // pdfUrl → wordt later ingevuld
      'pending',    // mail_status
      '',
      total, // subtotal = totaal
      0      // discount
    ]);

    const headRowIndex = head.getLastRow(); // net toegevoegde regel

    const linesSheet = ss.getSheetByName(LOG.lineSheet);
    const lineRows = cart.map(it => [
      receiptNo,
      it.sku,
      it.desc || '',
      Number(it.price) || 0,
      Number(it.qty) || 1,
      (Number(it.price) || 0) * (Number(it.qty) || 1),
      it.party || ''
    ]);
    if (lineRows.length) {
      linesSheet
        .getRange(linesSheet.getLastRow() + 1, 1, lineRows.length, 7)
        .setValues(lineRows);
    }
    
    // 80mm-bon: geen PDF meer opslaan, alleen de HTML-URL bewaren in kolom H
    try {
      const baseUrl = ScriptApp.getService().getUrl();
      const bon80Url = baseUrl + '?file=ticket&no=' + encodeURIComponent(receiptNo);

      // schrijf HTML-bon-URL in kolom H (bon80Url)
      head.getRange(headRowIndex, 8).setValue(bon80Url);
    } catch (err80) {
      Logger.log('Fout bij opslaan 80mm-bon URL: ' + err80);
    }

    // 🔥 ALTijd codes wissen na succesvolle boeking
    apiClearDiscount();
    apiClearGiftcard();

    // Cart & index reset
    clearCart();
    invalidateSkuIndex_();

    const props = PropertiesService.getDocumentProperties();
    props.setProperty(
      `POST_PROCESS_${receiptNo}`,
      JSON.stringify({ receiptNo, email })
    );
    ensurePostProcessTrigger_();

    // Bouw itemsSold in hetzelfde formaat als apiBookWithDiscountNet
    // 🔥 Kanaal opnieuw ophalen uit voorraad-sheet (enige betrouwbare bron)
    const itemsSold = cart.map(it => {
      let channel = '';

      try {
        const sh = ss.getSheetByName(it.sheetName);
        if (sh && it.row) {
          channel = String(
            sh.getRange(it.row, COL.channel).getValue() || ''
          ).trim();
        }
      } catch (e) {
        channel = '';
      }

      return {
        sku: it.sku,
        desc: it.desc || '',
        price: Number(it.price) || 0,
        qty: Number(it.qty) || 1,
        channel
      };
    });

    // 🔥 GIFT CARD ISSUE (alleen bij verkoop)
    try {
      giftcardIssueForBookedSale_({
        receiptNo: receiptNo,
        when: new Date(),
        items: itemsSold
      });
    } catch (e) {
      // bewust niet crashen, maar wel loggen
      Logger.log('Giftcard issue failed: ' + e.message);
    }

    // front-end gebruikt ticketUrl voor directe (live) 80mm-print
    return { total, receiptNo, ticketUrl, itemsSold };

  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** ================== BONNEN-OVERZICHT API ================== **/
/**
 * Data voor 80mm-herprint vanuit Bonnen-tab.
 * Front-end kan hiermee via buildReceipt80mmHtml opnieuw een venster openen.
 */
function apiGetReceiptForPrint(receiptNo) {
  ensureLogSheets_();

  const ss    = SpreadsheetApp.getActive();
  const head  = ss.getSheetByName(LOG.headSheet);
  const lines = ss.getSheetByName(LOG.lineSheet);
  if (!head || !lines) {
    throw new Error('Sales sheets ontbreken');
  }

  const lastHead = head.getLastRow();
  const hVals = lastHead > 1
    ? head.getRange(2, 1, lastHead - 1, Math.max(8, head.getLastColumn())).getValues()
    : [];

  const h = hVals.find(r => String(r[0]) === String(receiptNo));
  if (!h) {
    throw new Error('Bon niet gevonden');
  }

  const headObj = {
    receipt_no: String(h[0]),
    date:       h[1],
    pay:        String(h[2] || ''),
    total:      Number(h[3] || 0),
    email:      String(h[4] || ''),
    pdfUrl:     String(h[5] || ''),
    mail:       String(h[6] || ''),
    bon80Url:   String(h[7] || '')
  };

  const lastLine = lines.getLastRow();
  const lVals = lastLine > 1
    ? lines.getRange(2, 1, lastLine - 1, 7).getValues()
    : [];

  const items = lVals
    .filter(r => String(r[0]) === String(receiptNo))
    .map(r => ({
      sku:      String(r[1]),
      desc:     String(r[2] || ''),
      price:    Number(r[3] || 0),
      qty:      Number(r[4] || 1),
      subtotal: Number(r[5] || 0),
      party:    String(r[6] || '')
    }));

  return { head: headObj, items };
}


function apiListReceipts(limit, query) {
  try {
    // Defaults als er niks wordt meegegeven
    limit = Number(limit) || 200;
    query = (query || '').toString().trim();
    Logger.log('apiListReceipts START — limit=%s, query="%s"', limit, query);

    ensureLogSheets_();

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(LOG.headSheet); // 'Sales'
    if (!sh) {
      Logger.log('apiListReceipts: ❌ Sales-sheet niet gevonden');
      return [];
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    Logger.log('apiListReceipts: lastRow=%s, lastCol=%s', lastRow, lastCol);

    if (lastRow <= 1) {
      Logger.log('apiListReceipts: geen data onder de header (alleen rij 1)');
      return [];
    }

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    Logger.log('apiListReceipts: header=%s', JSON.stringify(header));

    const vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Header → index mapping
    const idx = {};
    header.forEach((h, i) => { idx[h] = i; });
    Logger.log('apiListReceipts: idx=%s', JSON.stringify(idx));

    // Als één van de kolommen niet bestaat, meteen loggen
    const required = [
      'receipt_no',
      'datum',
      'betaalwijze',
      'totaal',
      'klant_email',
      'pdfUrl',
      'mail_status',
      'bon80Url'
    ];
    required.forEach(k => {
      if (!(k in idx)) {
        Logger.log('apiListReceipts: ⚠️ kolom "%s" niet gevonden in header', k);
      }
    });

    // Rijen mappen naar objecten
    let rows = vals.map(r => {
      return {
        receipt_no: r[idx['receipt_no']] || '',
        date:       r[idx['datum']] || '',
        pay:        r[idx['betaalwijze']] || '',
        total:      Number(r[idx['totaal']] || 0),
        email:      r[idx['klant_email']] || '',
        pdfUrl:     r[idx['pdfUrl']] || '',
        mail:       r[idx['mail_status']] || '',
        bon80Url:   r[idx['bon80Url']] || ''
      };
    });

    Logger.log('apiListReceipts: ruwe mapped rows (eerste 3)=%s',
      JSON.stringify(rows.slice(0, 3))
    );

    // Filter op zoekterm (bonnummer of e-mail)
    const q = query.toLowerCase();
    if (q) {
      rows = rows.filter(r =>
        (r.receipt_no && String(r.receipt_no).toLowerCase().includes(q)) ||
        (r.email      && String(r.email).toLowerCase().includes(q))
      );
      Logger.log('apiListReceipts: na filter, rows=%s', rows.length);
    }

    // Nieuwste bovenaan op datum
    rows.sort((a, b) => {
      const da = new Date(a.date || 0).getTime();
      const db = new Date(b.date || 0).getTime();
      return db - da;
    });

    if (limit > 0 && rows.length > limit) {
      rows = rows.slice(0, limit);
    }

    Logger.log('apiListReceipts: RETURN rows.length=%s', rows.length);

    // Hard serialiseerbaar maken (voor de zekerheid)
    const safeRows = JSON.parse(JSON.stringify(rows));
    return safeRows;

  } catch (e) {
    Logger.log('apiListReceipts ERROR: ' + (e && e.message ? e.message : String(e)));
    throw e; // zodat withFailureHandler wordt getriggerd in de front-end
  }
}

function debugListReceipts() {
  const ss   = SpreadsheetApp.getActive();
  const sh   = ss.getSheetByName('Sales');
  if (!sh) {
    Logger.log('❌ Sheet "Sales" bestaat niet.');
    return;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  Logger.log('👉 LastRow: ' + lastRow + ', LastCol: ' + lastCol);

  const header = sh.getRange(1,1,1,lastCol).getValues()[0];
  Logger.log('👉 Header: ' + JSON.stringify(header));

  if (lastRow <= 1) {
    Logger.log('❌ Geen data onder de headers');
    return;
  }

  const rows = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  Logger.log('👉 Eerste 3 rijen:');
  Logger.log(JSON.stringify(rows.slice(0,3)));

  // Probeer mapping zoals apiListReceipts het doet
  const idx = {};
  header.forEach((h, i) => idx[h] = i);
  Logger.log('👉 Index mapping: ' + JSON.stringify(idx));

  const mapped = rows.map(r => ({
    receipt_no: r[idx['receipt_no']],
    datum:      r[idx['datum']],
    betaalwijze:r[idx['betaalwijze']],
    totaal:     r[idx['totaal']],
    klant_email:r[idx['klant_email']],
    pdfUrl:     r[idx['pdfUrl']],
    mail_status:r[idx['mail_status']],
    bon80Url:   r[idx['bon80Url']]
  }));

  Logger.log('👉 Eerste 3 mapped: ' + JSON.stringify(mapped.slice(0,3)));
}

function migrateSalesAddDiscountColumns_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Sales');
  if (!sh) throw new Error('Sales sheet niet gevonden');

  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];

  if (!header.includes('subtotal')) {
    sh.insertColumnAfter(sh.getLastColumn());
    sh.getRange(1, sh.getLastColumn()).setValue('subtotal');
  }

  if (!header.includes('discount')) {
    sh.insertColumnAfter(sh.getLastColumn());
    sh.getRange(1, sh.getLastColumn()).setValue('discount');
  }
}

function ensurePostProcessTrigger_() {
  const handler = 'processPostProcessQueue';
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction && t.getHandlerFunction() === handler);
  if (exists) return;

  // elke minuut de queue wegwerken
  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(1)
    .create();
}

function processPostProcessQueue() {
  const props = PropertiesService.getDocumentProperties();
  const keys = props.getKeys().filter(k => k.startsWith('POST_PROCESS_'));
  if (!keys.length) return;

  const ss = SpreadsheetApp.getActive();

  keys.forEach(key => {
    try {
      const { receiptNo, email } = JSON.parse(props.getProperty(key));
      props.deleteProperty(key);

      // --- haal header uit Sales ---
      const head = ss.getSheetByName(LOG.headSheet);
      const headVals = head.getRange(2, 1, head.getLastRow() - 1, 10).getValues();
      const hIdx = headVals.findIndex(r => String(r[0]) === String(receiptNo));
      if (hIdx < 0) return;

      const now       = headVals[hIdx][1];
      const payMethod = headVals[hIdx][2];
      const total     = Number(headVals[hIdx][3] || 0);

      // --- haal regels uit Sales_Lines ---
      const linesSheet = ss.getSheetByName(LOG.lineSheet);
      const lineVals = linesSheet.getRange(2, 1, linesSheet.getLastRow() - 1, 7).getValues();
      const myLines = lineVals.filter(r => String(r[0]) === String(receiptNo));
      if (!myLines.length) return;

      const rows = myLines.map(r => {
        const price = Number(r[3] || 0);
        const qty   = Number(r[4] || 1);
        return [r[1], r[2] || '', price, qty, price * qty];
      });

      // --- TEMP bon-sheet ---
      const tmpName = ('Bon_' + receiptNo).slice(0, 99);
      let bon = ss.getSheetByName(tmpName);
      if (bon) ss.deleteSheet(bon);
      bon = ss.insertSheet(tmpName);

      bon.getRange('A10:E10').setValues([[
        'SKU','Omschrijving','Prijs','Aantal','Subtotaal'
      ]]).setFontWeight('bold');

      bon.getRange(11, 1, rows.length, 5).setValues(rows);

      const totRow = 11 + rows.length;
      bon.getRange(totRow, 4)
        .setValue('Totaal:')
        .setFontWeight('bold')
        .setHorizontalAlignment('right');

      bon.getRange(totRow, 5)
        .setValue(total)
        .setNumberFormat('€ #,##0.00')
        .setFontWeight('bold');

      styleBon_(bon, receiptNo, payMethod, email || '', total, now);
      SpreadsheetApp.flush();

      // --- PDF ---
      const pdfUrl = Utilities.formatString(
        'https://docs.google.com/spreadsheets/d/%s/export?format=pdf&gid=%s&portrait=true&size=A5&gridlines=false',
        ss.getId(),
        bon.getSheetId()
      );

      // --- Mail ---
      let mailStatus = 'no email';
      if (email && /\S+@\S+\.\S+/.test(email)) {
        const token = ScriptApp.getOAuthToken();
        const pdfBlob = UrlFetchApp.fetch(pdfUrl, {
          headers: { Authorization: 'Bearer ' + token }
        }).getBlob().setName(`Golf-Locker-${receiptNo}.pdf`);

        const htmlBody = `
            <div style="font-family:Arial,Helvetica,sans-serif; color:#1f2937;">
              <div style="border-bottom:2px solid #297900; padding-bottom:10px; margin-bottom:20px;">
                <img src="https://shop.golf-locker.nl/wp-content/uploads/2024/04/Golf-Locker-Logo-1-scaled.png"
                    style="height:50px;" alt="Golf Locker">
              </div>

              <h2 style="color:#297900;">Bedankt voor je aankoop!</h2>

              <p>
                Bedankt voor je aankoop bij <strong>Golf Locker</strong>.<br>
                In de bijlage vind je de bijhorende factuur.
              </p>

              <table style="margin-top:15px; border-collapse:collapse;">
                <tr>
                  <td style="padding:4px 8px;"><strong>Factuurnummer:</strong></td>
                  <td style="padding:4px 8px;">${receiptNo}</td>
                </tr>
                <tr>
                  <td style="padding:4px 8px;"><strong>Totaal:</strong></td>
                  <td style="padding:4px 8px;">€ ${total.toFixed(2)}</td>
                </tr>
                 <tr>
                  <td style="padding:4px 8px;"><strong>Status:</strong></td>
                  <td style="padding:4px 8px;">Betaald</td>
                </tr>
                <tr>
                  <td style="padding:4px 8px;"><strong>Betaalwijze:</strong></td>
                  <td style="padding:4px 8px;">${payMethod || '-'}</td>
                </tr>
              </table>

              <p style="margin-top:25px;">
                Met sportieve groet,<br>
                <strong>Golf Locker</strong><br>
                <a href="https://shop.golf-locker.nl">shop.golf-locker.nl</a>
              </p>

              <hr style="margin-top:30px; border:none; border-top:1px solid #e5e7eb;">
              <p style="font-size:12px; color:#6b7280;">
                Deze e-mail is automatisch verstuurd na je aankoop in onze winkel.
              </p>
            </div>
          `;
          
        MailApp.sendEmail({
            to: email,
            subject: `Factuur ${receiptNo} – Golf Locker`,
            htmlBody: htmlBody,
            name: 'Golf Locker',
            replyTo: 'info@golf-locker.nl',
            attachments: [pdfBlob]
          });

        mailStatus = 'sent';
      }

      // --- schrijf terug in Sales ---
      head.getRange(hIdx + 2, 6).setValue(pdfUrl);
      head.getRange(hIdx + 2, 7).setValue(mailStatus);

      ss.deleteSheet(bon);

    } catch (e) {
      Logger.log('Queue error: ' + e);
    }
  });
}

function apiSetCart(cart){
  if (!Array.isArray(cart)) throw new Error('Ongeldige cart');
  saveCart(cart);
  return true;
}
