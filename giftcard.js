/** ============================
 *  GIFTCARD (state + lookup)
 *  Sheet: "Codes"
 *  Type: "GIFTCARD"
 *  ============================ */

/** Key per gebruiker (zelfde patroon als discount) */
function _giftcardKey_() {
  const email = Session.getActiveUser()?.getEmail() || 'anon';
  return `GIFTCARD::${email}::${SpreadsheetApp.getActive().getId()}`;
}

function _getActiveGiftcard_() {
  const raw = CacheService.getUserCache().get(_giftcardKey_());
  return raw ? JSON.parse(raw) : null;
}

function _setActiveGiftcard_(obj) {
  CacheService.getUserCache().put(_giftcardKey_(), JSON.stringify(obj), 7200);
}

function apiClearGiftcard() {
  CacheService.getUserCache().remove(_giftcardKey_());
  return { ok:true };
}


function apiGetActiveGiftcard() {
  return { ok: true, giftcard: _getActiveGiftcard_() };
}

/**
 * Activeer giftcard in de POS (op basis van tab Codes).
 * Verwacht: type == "GIFTCARD"
 * Vereist: actief, niet verlopen, saldo > 0
 */
function apiApplyGiftcardCode(codeRaw) {
  const lookup = apiCodeLookup(codeRaw); // komt uit discount.gs (bestaat al)
  if (!lookup?.ok) return { ok: false, error: lookup?.error || 'Lookup mislukt' };
  if (!lookup.found) return { ok: false, error: 'Code niet gevonden' };

  const type = String(lookup.type || '').toUpperCase();
  if (type !== 'GIFTCARD') return { ok: false, error: 'Geen cadeaukaart' };
  if (!lookup.isUsable) return { ok: false, error: 'Cadeaukaart niet actief of verlopen' };

  // saldo kan in sheet als getal of tekst staan
  const saldo = Number(String(lookup.saldo ?? '0').replace(',', '.')) || 0;
  if (saldo <= 0) return { ok: false, error: 'Saldo is 0' };

  const gc = {
    code: lookup.code,
    saldo: Math.round(saldo * 100) / 100
  };

  _setActiveGiftcard_(gc);

  // 🔥 behandel giftcard als vaste korting
  _setActiveDiscount_({
    code: gc.code,
    waarde: gc.saldo,
    waarde_type: 'FIXED',
    isGiftcard: true
  });

  return {
    ok: true,
    giftcard: gc
  };
}

/** ============================
 *  TOTALS WITH GIFTCARD (read-only)
 *  ============================ */

function apiGetCartTotalsWithGiftcard() {
  const cart = getCart();
  if (!cart || cart.length === 0) {
    // 🔥 Geen cart = geen cadeaubon
    apiClearGiftcard();
    return { ok: true, totalToPay: 0 };
  }
  const gc   = _getActiveGiftcard_();  // uit giftcard.gs

  const subtotal = cart.reduce(
    (s, it) => s + (Number(it.price) || 0) * (Number(it.qty) || 1),
    0
  );

  if (!gc) {
    return {
      ok: true,
      subtotal,
      giftcard: null,
      totalToPay: subtotal
    };
  }

  const saldo = Number(gc.saldo) || 0;
  const applied = Math.min(saldo, subtotal);
  const remaining = Math.round((saldo - applied) * 100) / 100;

  return {
    ok: true,
    subtotal,
    giftcard: {
      code: gc.code,
      applied,
      remaining
    },
    totalToPay: Math.round((subtotal - applied) * 100) / 100
  };
}

function apiBookWithGiftcardNet(payMethod, customerEmail) {
  const totals = apiGetCartTotalsWithGiftcard();
  const giftcardAmount = Number(totals?.giftcard?.applied || 0);
  if (!totals || totals.ok === false) {
    throw new Error('Kon giftcard-totals niet bepalen');
  }

  const res = apiBookAndReceipt(payMethod, customerEmail);

  res.subtotal = totals.subtotal;
  res.giftcard = totals.giftcard;
  res.total    = totals.totalToPay;

  // 🔥 NIEUW: saldo afboeken
  if (totals.giftcard?.code && totals.giftcard.applied > 0) {
    _applyGiftcardTransaction_(
      totals.giftcard.code,
      totals.giftcard.applied,
      res.receiptNo
    );
  }

  return res;
}

/***********************
 * STEP 4A — APPLY GIFTCARD TRANSACTION
 ***********************/

function _applyGiftcardTransaction_(code, amount, receiptNo) {
  if (!code || amount <= 0) return;

  const sh = SpreadsheetApp.getActive().getSheetByName('Codes');
  if (!sh) throw new Error('Codes-sheet niet gevonden');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('Geen giftcards gevonden');

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (String(row[idx.code]).toUpperCase() !== String(code).toUpperCase()) continue;

    const saldoCol   = idx.saldo + 1;
    const gebruiktCol = idx.gebruikt + 1;
    const lastTxCol  = idx.laatste_transactie + 1;
    const r = i + 2;

    const currentSaldo = Number(row[idx.saldo]) || 0;
    const newSaldo = Math.max(0, currentSaldo - amount);

    sh.getRange(r, saldoCol).setValue(newSaldo);
    sh.getRange(r, gebruiktCol).setValue((Number(row[idx.gebruikt]) || 0) + 1);
    sh.getRange(r, lastTxCol).setValue(
      `${receiptNo} | -€${amount.toFixed(2)}`
    );

    return;
  }

  throw new Error('Giftcard niet gevonden bij afboeken');
}

/***********************
 * STEP 6 — Issue giftcards on checkout
 ***********************/

// 🔧 HIER kun je later alles aanpassen (prefix / code prefix)
const GIFTCARD_CFG = {
  skuPrefix: 'GIFTCARD',  // matcht GIFTCARD+1, GIFTCARD+2, ...
  codePrefix: 'GC+',       // codes worden GC-XXXXXXXX
  sheetName: 'Codes',
  typeValue: 'GIFTCARD'
};

function _ensureCodesCols_(sh, wanted) {
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const idx = {};
  header.forEach((h,i)=> idx[h] = i + 1);

  wanted.forEach(name => {
    if (!idx[name]) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(name);
      idx[name] = sh.getLastColumn();
    }
  });

  return idx; // 1-based col index
}

function _randomCode_(len) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // geen O/0/I/1
  let s = '';
  for (let i=0;i<len;i++) s += chars[Math.floor(Math.random()*chars.length)];
  return s;
}

function _generateUniqueGiftcardCode_(existingSet) {
  for (let tries = 0; tries < 50; tries++) {
    const code = GIFTCARD_CFG.codePrefix + _randomCode_(8);
    if (!existingSet.has(code)) return code;
  }
  throw new Error('Kon geen unieke giftcard-code genereren');
}

/**
 * Maakt giftcards aan in Codes op basis van verkochte items.
 * @param {Object} opts
 * @param {string} opts.receiptNo
 * @param {Date}   opts.when
 * @param {Array}  opts.items  (itemsNet uit apiBookWithDiscountNet)
 */
function giftcardIssueForBookedSale_(opts) {
  const receiptNo = String(opts?.receiptNo || '').trim();
  const when = opts?.when || new Date();
  const items = Array.isArray(opts?.items) ? opts.items : [];

  if (!receiptNo) throw new Error('giftcardIssueForBookedSale_: receiptNo ontbreekt');

  // 1) Zoek giftcard-regels
  const gcLines = items.filter(it => {
    const sku = String(it?.sku || '').toUpperCase().trim();
    return sku === GIFTCARD_CFG.skuPrefix || sku.startsWith(GIFTCARD_CFG.skuPrefix + '+');
  });

  if (!gcLines.length) return []; // niks te doen

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(GIFTCARD_CFG.sheetName);
  if (!sh) throw new Error(`Sheet '${GIFTCARD_CFG.sheetName}' niet gevonden`);

  // 2) Zorg dat kolommen bestaan (minimaal jouw bestaande headers + extra logging)
  // bestaande uit discount: code,type,waarde,waarde_type,saldo,actief,vervaldatum,gebruikt
  const cols = _ensureCodesCols_(sh, [
    'code','type','waarde','waarde_type','saldo','actief','vervaldatum','gebruikt',
    'bron','receipt_no','created_at'
  ]);

  // 3) Existing codes set (uniqueness)
  const lastRow = sh.getLastRow();
  const existing = new Set();
  if (lastRow >= 2) {
    const codeVals = sh.getRange(2, cols.code, lastRow - 1, 1).getValues();
    codeVals.forEach(r => {
      const c = String(r[0] || '').trim().toUpperCase();
      if (c) existing.add(c);
    });
  }

  // 4) Bouw rows (1 code per qty)
  const out = [];
  const rowsToAppend = [];

  gcLines.forEach(it => {
    const qty = Math.max(1, Number(it.qty) || 1);
    const unit = Math.max(0, Number(it.price) || 0); // NETTO unit (jouw systeem)
    for (let i=0;i<qty;i++) {
      const code = _generateUniqueGiftcardCode_(existing);
      existing.add(code);

      // defaults
      const row = [];
      row[cols.code - 1]        = code;
      row[cols.type - 1]        = GIFTCARD_CFG.typeValue;
      row[cols.waarde - 1]      = '';        // niet gebruikt bij giftcard
      row[cols.waarde_type - 1] = '';        // niet gebruikt bij giftcard
      row[cols.saldo - 1]       = unit;      // startsaldo = verkoopwaarde
      row[cols.actief - 1]      = true;
      row[cols.vervaldatum - 1] = '';
      row[cols.gebruikt - 1]    = '';

      // extra logging
      row[cols.bron - 1]        = 'POS';
      row[cols.receipt_no - 1]  = receiptNo;
      row[cols.created_at - 1]  = when;

      rowsToAppend.push(row);
      out.push({ code, amount: unit });
    }
  });

  // 5) Append
  sh.getRange(sh.getLastRow() + 1, 1, rowsToAppend.length, sh.getLastColumn()).setValues(rowsToAppend);

  if (out.length) {
    CacheService.getUserCache().put(
      'LAST_GIFTCARD_CREATED',
      JSON.stringify({
        code: out[out.length - 1].code,
        amount: out[out.length - 1].amount,
        receiptNo: receiptNo,
        createdAt: when
      }),
      3600
    );
  }

  return out; // terug naar caller (handig voor later printen/UI)
}

/*********************************
 * COMPAT SHIM — legacy call support
 * (frontend verwacht giftcardGetLastCreated_)
 *********************************/
function giftcardGetLastCreated_() {
  const cache = CacheService.getUserCache();
  const raw = cache.get('LAST_GIFTCARD_CREATED');
  return raw ? JSON.parse(raw) : null;
}

/**
 * (optioneel, maar handig) Frontend endpoint
 * Als je app.html via google.script.run dit wil ophalen.
 */
function apiGiftcardGetLastCreated() {
  return giftcardGetLastCreated_();
}

function apiDebugGetActiveGiftcard() {
  return _getActiveGiftcard_();
}

