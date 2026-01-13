/** ============================
 *  CODES (read-only lookup)
 *  Sheet: "Codes"
 *  ============================ */

function apiCodeLookup(codeRaw) {
  try {
    const code = _normCode_(codeRaw);
    if (!code) return { ok:true, found:false, reason:"EMPTY" };

    const sh = SpreadsheetApp.getActive().getSheetByName('Codes');
    if (!sh) return { ok:false, error:"Sheet 'Codes' niet gevonden." };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2) return { ok:true, found:false, reason:"NO_DATA" };

    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
    const idx = _headerIndexMap_(headers);
    const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowCode = _normCode_(row[idx.code] ?? '');
      if (rowCode !== code) continue;

      const out = {
        ok: true,
        found: true,
        rowNumber: i + 2,
        code: String(row[idx.code] ?? ''),
        type: String(row[idx.type] ?? ''),
        waarde: Number(row[idx.waarde] || 0),
        waarde_type: String(row[idx.waarde_type] || '').toUpperCase(),
        saldo: row[idx.saldo] ?? '',
        actief: row[idx.actief] ?? '',
        vervaldatum: row[idx.vervaldatum] ?? '',
        gebruikt: row[idx.gebruikt] ?? ''
      };

      out.isActive  = _truthy_(out.actief);
      out.isExpired = _isExpired_(out.vervaldatum);
      out.isUsable  = out.isActive && !out.isExpired;

      return out;
    }

    return { ok:true, found:false, reason:"NOT_FOUND" };

  } catch (e) {
    return { ok:false, error: String(e.message || e) };
  }
}

function _headerIndexMap_(headers) {
  const map = {};
  headers.forEach((h,i) => map[h] = i);
  const get = k => (k in map ? map[k] : -1);

  return {
    code: get('code'),
    type: get('type'),
    waarde: get('waarde'),
    waarde_type: get('waarde_type'),
    saldo: get('saldo'),
    actief: get('actief'),
    vervaldatum: get('vervaldatum'),
    gebruikt: get('gebruikt')
  };
}

function _normCode_(v) {
  return String(v || '').replace(/[\u200B-\u200D\uFEFF]/g,'').trim().toUpperCase();
}
function _truthy_(v) {
  return v === true || v === 1 || ['true','ja','yes','1'].includes(String(v).toLowerCase());
}
function _isExpired_(v) {
  if (!v) return false;
  const d = v instanceof Date ? v : new Date(v);
  if (isNaN(d)) return false;
  return new Date() > new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23,59,59);
}

/** ============================
 *  DISCOUNT STATE (UserCache)
 *  ============================ */

function _discountKey_() {
  const email = Session.getActiveUser()?.getEmail() || 'anon';
  return `DISCOUNT::${email}::${SpreadsheetApp.getActive().getId()}`;
}

function _getActiveDiscount_() {
  const raw = CacheService.getUserCache().get(_discountKey_());
  return raw ? JSON.parse(raw) : null;
}

function _setActiveDiscount_(obj) {
  CacheService.getUserCache().put(_discountKey_(), JSON.stringify(obj), 7200);
}

function apiClearDiscount() {
  CacheService.getUserCache().remove(_discountKey_());
  return { ok:true };
}

/** ============================
 *  APPLY DISCOUNT CODE
 *  ============================ */

function apiApplyDiscountCode(codeRaw) {
  const lookup = apiCodeLookup(codeRaw);
  if (!lookup?.found) return { ok:false, error:'Code niet gevonden' };
  if (lookup.type !== 'DISCOUNT') return { ok:false, error:'Geen kortingscode' };
  if (!lookup.isUsable) return { ok:false, error:'Code niet actief of verlopen' };

  const discount = {
    code: lookup.code,
    waarde: lookup.waarde,
    waarde_type: lookup.waarde_type
  };

  _setActiveDiscount_(discount);
  return { ok:true, discount };
}
/**
 * Apply code vanuit 1 invoerveld:
 * - DISCOUNT  -> apiApplyDiscountCode
 * - GIFTCARD  -> apiApplyGiftcardCode (moet bestaan)
 */
function apiApplyAnyCode(codeRaw) {
  const lookup = apiCodeLookup(codeRaw);
  if (!lookup?.found) return { ok:false, error:'Code niet gevonden' };
  if (!lookup.isUsable) return { ok:false, error:'Code niet actief of verlopen' };

  if (lookup.type === 'DISCOUNT') {
    return apiApplyDiscountCode(codeRaw);
  }

  if (lookup.type === 'GIFTCARD') {
    if (typeof apiApplyGiftcardCode !== 'function') {
      return { ok:false, error:'Giftcard functie ontbreekt (apiApplyGiftcardCode)' };
    }
    return apiApplyGiftcardCode(codeRaw);
  }

  return { ok:false, error:'Onbekend code type: ' + lookup.type };
}


/** ============================
 *  TOTALS WITH DISCOUNT
 *  ============================ */

function apiGetCartTotalsWithDiscount() {
  const cart = getCart();
  if (!cart || cart.length === 0) {
    // 🔥 Geen cart = geen korting
    apiClearDiscount();
    return { ok: true, total: 0 };
  }
  const discount = _getActiveDiscount_();

  let subtotal = cart.reduce((s,it)=> s + (Number(it.price)||0)*(Number(it.qty)||1), 0);
  let discountAmount = 0;

  if (discount) {
    discountAmount = discount.waarde_type === 'PERCENT'
      ? subtotal * (discount.waarde / 100)
      : discount.waarde;
  }

  discountAmount = Math.min(discountAmount, subtotal);

  return {
    ok: true,
    subtotal,
    discount: discount ? { code: discount.code, amount: discountAmount } : null,
    total: subtotal - discountAmount
  };
}

/** ============================
 *  BOOK WITH DISCOUNT (wrapper)
 *  ============================ */

function apiBookWithDiscount(payMethod, customerEmail) {
  const totals = apiGetCartTotalsWithDiscount();
  const res = apiBookAndReceipt(payMethod, customerEmail);

  res.total = totals.total;
  res.discount = totals.discount;

  apiClearDiscount();
  return res;

}

/***********************
 * STEP 3.6 — NETTO BOEKEN (zonder pos.gs te wijzigen)
 * Nieuwe endpoint: apiBookWithDiscountNet()
 ***********************/

function _round2_(n){
  return Math.round((Number(n) || 0) * 100) / 100;
}

/**
 * Verdeelt totale korting proportioneel over cartregels.
 * Result: itemsNet met aangepaste unit prijzen (netto).
 */
function _applyDiscountToCartLines_(cart, discountAmount){
  discountAmount = _round2_(discountAmount);

  if (!cart || !cart.length || discountAmount <= 0) {
    return { itemsNet: (cart || []).map(x => Object.assign({}, x)), discountAmount: 0 };
  }

  const lines = cart.map(it => {
    const price = Number(it.price) || 0;
    const qty   = Number(it.qty) || 1;
    return { it, grossLine: _round2_(price * qty) };
  });

  const grossTotal = _round2_(lines.reduce((s,l)=> s + l.grossLine, 0));
  if (grossTotal <= 0) return { itemsNet: cart.map(x => Object.assign({}, x)), discountAmount: 0 };

  if (discountAmount > grossTotal) discountAmount = grossTotal;

  const targetNetTotal = _round2_(grossTotal - discountAmount);

  let runningNet = 0;
  const out = [];

  for (let i = 0; i < lines.length; i++){
    const l = lines[i];

    let netLine;
    if (i === lines.length - 1) {
      netLine = _round2_(targetNetTotal - runningNet);
    } else {
      const share = l.grossLine / grossTotal;
      netLine = _round2_(l.grossLine - (discountAmount * share));
      runningNet = _round2_(runningNet + netLine);
    }

    const qty = Number(l.it.qty) || 1;
    const unitNet = qty > 0 ? _round2_(netLine / qty) : 0;

    out.push(Object.assign({}, l.it, { price: unitNet }));
  }

  return { itemsNet: out, discountAmount };
}

/**
 * ✅ Dit is de functie die je frontend moet aanroepen i.p.v. apiBookAndReceipt
 * - boekt NETTO in voorraad (kolom G)
 * - schrijft Sales + Sales_Lines NETTO
 * - bouwt Bon-sheet (A5) NETTO
 * - ticketUrl blijft werken
 */
function apiBookWithDiscountNet(payMethod, customerEmail) {
 
  const lock = LockService.getDocumentLock();
  lock.tryLock(5000);

  try {
    ensureLogSheets_();

    const ss   = SpreadsheetApp.getActive();
    const cart = getCart() || [];
    if (cart.length === 0) {
      throw new Error('Mandje is leeg (backend)');
    }
    if (!cart || cart.length === 0) {
      apiClearDiscount();
      apiClearGiftcard();
      throw new Error('Mandje is leeg');
    }

    // 1) Bepaal kortingbedrag via jouw bestaande totals-functie
    const totals = apiGetCartTotalsWithDiscount();
    // 🔥 Laatste waarheid: als er GEEN actieve discount meer is → forceer 0
    const activeDiscount = _getActiveDiscount_?.() || null;

    if (!activeDiscount) {
      totals.discount = null;
    }

    if (!totals || totals.ok === false) throw new Error(totals?.error || 'Kon korting niet bepalen');

    const discountAmount = activeDiscount
      ? Number(totals.discount?.amount || 0)
      : 0;
    const activeGiftcard = _getActiveGiftcard_?.() || null;
    const giftcardApplied = activeGiftcard && discountAmount > 0
      ? {
          code: activeGiftcard.code,
          applied: discountAmount
        }
      : null;


    // 2) Maak netto-regels (unitprijzen aangepast)
    const { itemsNet } = _applyDiscountToCartLines_(cart, discountAmount);
    // Kanaal (kolom O) toevoegen aan itemsNet (voor post-sale overzicht)
    itemsNet.forEach(it => {
      try {
        const sh = ss.getSheetByName(it.sheetName);
        if (!sh) { it.channel = ''; return; }
        it.channel = String(sh.getRange(it.row, COL.channel).getValue() || '').trim();
      } catch(e) {
        it.channel = '';
      }
    });


    const now   = new Date();
    const email = (customerEmail || '').trim();
    const receiptNo = nextReceiptNo_();

    let total = 0;

    // groepeer per sheet
    const bySheet = {};
    itemsNet.forEach(it => {
      if (!bySheet[it.sheetName]) bySheet[it.sheetName] = [];
      bySheet[it.sheetName].push(it);
    });

    // 3) schrijf NETTO sales in voorraad-tabbladen
    Object.keys(bySheet).forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) throw new Error('Tab niet gevonden: ' + name);

      const items = bySheet[name];
      const last  = sh.getLastRow();
      if (last <= 1) throw new Error('Geen data in ' + name);

      const skuCol  = sh.getRange(2, COL.sku,  last - 1, 1).getValues().map(r => String(r[0]).trim());
      const saleCol = sh.getRange(2, COL.sale, last - 1, 1).getValues().map(r => r[0]);

      items.forEach(it => {
        let need = Number(it.qty) || 1;
        const rowsToWrite = [];

        for (let i = 0; i < skuCol.length && need > 0; i++) {
          if (skuCol[i] === String(it.sku).trim() && (saleCol[i] === '' || saleCol[i] === null)) {
            rowsToWrite.push(i + 2);
            need--;
          }
        }

        if (need > 0) {
          throw new Error(`Niet genoeg vrije regels voor SKU ${it.sku} in ${name} (ontbreken: ${need})`);
        }

        rowsToWrite.forEach(r => {
          const unit = Number(it.price) || 0;
          sh.getRange(r, COL.sale).setValue(unit);              // ✅ netto unit
          sh.getRange(r, COL.saleDate).setValue(now).setNumberFormat('dd-mm-yyyy');
          sh.getRange(r, COL.expected).setValue(0);
          sh.getRange(r, COL.expMargin).setValue(0);
          total += unit;                                       // ✅ totaal netto (som units)
          saleCol[r - 2] = unit;
        });
      });
    });

    total = _round2_(total);

    // 4) Bon-sheet (A5) opbouwen met NETTO regels
    const bon = ss.getSheetByName('Bon') || ss.insertSheet('Bon');
    bon.clear();

    bon.getRange('A10:E10').setValues([[
      'SKU','Omschrijving','Prijs','Aantal','Subtotaal'
    ]]).setFontWeight('bold');

    const rows = itemsNet.map(it => {
      const unit = Number(it.price) || 0;
      const qty  = Number(it.qty) || 1;
      return [
        it.sku,
        it.desc || '',
        unit,
        qty,
        _round2_(unit * qty)
      ];
    });

    if (rows.length) bon.getRange(11, 1, rows.length, 5).setValues(rows);

    const giftcard = giftcardGetLastCreated_?.() || null;

    styleBon_(bon, receiptNo, payMethod, email, total, now, {
      subtotal: total + discountAmount,
      discount: giftcardApplied ? 0 : discountAmount,
      giftcard: giftcardApplied
    });
    SpreadsheetApp.flush();

    // PDF-url
    const pdfUrl = Utilities.formatString(
      'https://docs.google.com/spreadsheets/d/%s/export?format=pdf&gid=%s&portrait=true&size=A5&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gridlines=false',
      ss.getId(),
      bon.getSheetId()
    );

    // mail (zelfde gedrag als pos.gs)
    let mailStatus = 'no email';
    const looksLikeEmail = /\S+@\S+\.\S+/.test(email);
    if (email) {
      if (looksLikeEmail) {
        try {
          const token = ScriptApp.getOAuthToken();
          const pdfBlob = UrlFetchApp.fetch(pdfUrl, {
            headers: { Authorization: 'Bearer ' + token }
          }).getBlob().setName(`Golf-Locker-Factuur-${receiptNo}.pdf`);
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
        } catch (err) {
          mailStatus = 'error: ' + (err?.message || String(err));
        }
      } else {
        mailStatus = 'invalid email';
      }
    }

    const baseUrl   = ScriptApp.getService().getUrl();
    const ticketUrl = baseUrl + '?file=ticket&no=' + encodeURIComponent(receiptNo);

    // 5) Loggen Sales + Sales_Lines (NETTO)
    const head = ss.getSheetByName(LOG.headSheet);
      head.appendRow([
        receiptNo,
        now,
        String(payMethod || ''),
        total,                              // netto totaal
        String(email || ''),
        pdfUrl,
        mailStatus,
        ticketUrl,                          // bon80Url
        _round2_(total + discountAmount),   // subtotal
        _round2_(discountAmount)            // discount
      ]);
    const headRowIndex = head.getLastRow();

    const linesSheet = ss.getSheetByName(LOG.lineSheet);
    const lineRows = itemsNet.map(it => {
      const unit = Number(it.price) || 0;
      const qty  = Number(it.qty) || 1;
      return [
        receiptNo,
        it.sku,
        it.desc || '',
        unit,
        qty,
        _round2_(unit * qty),
        it.party || ''
      ];
    });

    if (lineRows.length) {
      linesSheet.getRange(linesSheet.getLastRow() + 1, 1, lineRows.length, 7).setValues(lineRows);
    }

    // bon80Url in kolom H
    try {
      head.getRange(headRowIndex, 8).setValue(ticketUrl);
    } catch (e) {}

    // 6) Giftcards uitgeven (als er GIFTCARD+ regels in itemsNet zitten)
    let giftcardsIssued = [];
    try {
      giftcardsIssued = giftcardIssueForBookedSale_({
        receiptNo,
        when: now,
        items: itemsNet
      });
    } catch (e) {
      // Niet hard failen op giftcards, maar wél loggen
      Logger.log('giftcardIssueForBookedSale_ ERROR: ' + (e?.message || e));
    }
    // 6B) Giftcard SALDO AFBOEKEN (bij gebruik als betaalmiddel)

    if (activeGiftcard && discountAmount > 0) {
      _applyGiftcardTransaction_(
        activeGiftcard.code,
        discountAmount,
        receiptNo
      );

      apiClearGiftcard(); // 🔥 heel belangrijk: session resetten
    }

    // opruimen
    clearCart();
    apiClearDiscount();
    apiClearGiftcard(); 
    invalidateSkuIndex_();

    return {
      pdfUrl,
      total,
      receiptNo,
      ticketUrl,
      discountAmount: _round2_(discountAmount),
      subtotal: _round2_(total + discountAmount),
      giftcardsIssued, // [{code, amount}, ...]
      itemsSold: itemsNet.map(it => ({
        sku: String(it.sku || ''),
        desc: String(it.desc || ''),
        channel: String(it.channel || '')
      }))
    };


  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function apiDebugGetActiveDiscount() {
  return _getActiveDiscount_();
}

