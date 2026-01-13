function findSkuGlobal_(sku) {
  const ss = SpreadsheetApp.getActive();
  const tabs = ["Clubs","Sets","Tassen","Trolley's","Overig","Diensten"];

  for (let t of tabs) {
    const sh = ss.getSheetByName(t);
    if (!sh) continue;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue; // alleen header
    const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === sku) {
        return { sheet: t, row: i + 2 };
      }
    }
  }
  return null;
}

function norm(s) {
  return String(s || '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')  // zero width weg
    .replace(/\s/g, '')                     // spaties weg
    .toUpperCase();
}

function jsonSafe(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function apiReturnGetSafe(receiptNo) {
  Logger.log("SAFE START raw:", receiptNo);

  const no = String(receiptNo || "").trim();
  Logger.log("SAFE cleaned:", no);

  const data = apiGetReceiptForPrint(no);
  Logger.log("SAFE after apiGetReceiptForPrint:", JSON.stringify(data));

  if (!data || !data.head) {
    throw new Error("Bon niet gevonden: " + no);
  }

  const result = {
    head: {
      receipt_no: data.head.receipt_no || "",
      date:       String(data.head.date || ""),
      pay:        data.head.pay || "",
      total:      Number(data.head.total || 0),
      email:      data.head.email || "",
      pdfUrl:     String(data.head.pdfUrl || ""),
      mail:       data.head.mail || "",
      bon80Url:   String(data.head.bon80Url || "")
    },
    lines: (data.items || []).map(it => {
      const sku = String(it.sku || "");

      return {
        sku:      sku,
        desc:     String(it.desc || ""),
        qty:      Number(it.qty || 0),
        price:    Number(it.price || 0),
        subtotal: Number(it.subtotal || 0),
        alreadyReturned: isAlreadyReturned_(no, sku)
      };
    })
  };

  Logger.log("SAFE FINAL RESULT:", JSON.stringify(result));
  return jsonSafe(result);
}

function debugReturnGet() {
  const testNo = "GL-20251118-009";  // <== vul zelf tijdelijk een bestaand bonnummer in
  const res = apiReturnGetSafe(testNo);
  Logger.log("DEBUG RESULT:\n" + JSON.stringify(res, null, 2));
  return res;
}

function apiProcessReturn(receiptNo, payload) {
  Logger.log('[RET] apiProcessReturn START');
  Logger.log('[RET] receiptNo = ' + receiptNo);
  Logger.log('[RET] payload = ' + JSON.stringify(payload));

  if (!payload || !payload.items) {
    throw new Error("Geen retourdata ontvangen (payload.items ontbreekt)");
  }

  // payload kan een object of array zijn → altijd array maken
  let items = payload.items;
  if (!Array.isArray(items)) {
    items = Object.values(items);
  }

  // Alleen items met "selected: true"
  const itemsToReturn = items.filter(it => it.selected === true);

  if (!itemsToReturn.length) {
    throw new Error("Geen artikelen geselecteerd voor retour.");
  }

  const ss = SpreadsheetApp.getActive();
  const retourSheet = ss.getSheetByName("Retouren");
  if (!retourSheet) throw new Error("Tab 'Retouren' ontbreekt.");

  const now = new Date();
  const returnNo = generateReturnNumber();
  const baseUrl = ScriptApp.getService().getUrl();
  const returnTicketUrl =
    baseUrl + '?file=returnticket&no=' + encodeURIComponent(returnNo);

  // 🔒 VALIDATIE: check vooraf of er al retouren zijn
  itemsToReturn.forEach(it => {
    const sku = String(it.sku || "").trim();
    if (!sku) return;

    if (isAlreadyReturned_(receiptNo, sku)) {
      throw new Error(
        `SKU ${sku} van bon ${receiptNo} is al geretourneerd`
      );
    }
  });

    // ✅ Bereken refund total (som van returned items)
  const refundTotal = itemsToReturn.reduce((s, it) => {
    const p = Math.abs(Number(it.price || 0));
    const q = Math.max(1, Number(it.qty || 1)); // als jij geen qty meegeeft: blijft 1
    return s + (p * q);
  }, 0);

  // ✅ Pas Sales totaal aan (D = oud - refundTotal)
  _applyReturnToSalesTotal_(receiptNo, refundTotal);

  // ✅ Maak Sales_Lines regel(s) negatief
  _applyReturnToSalesLines_(receiptNo, itemsToReturn);
  
  // Voor elk geselecteerd item
  itemsToReturn.forEach(it => {
    const sku = String(it.sku || "").trim();
    const reason = String(it.reason || "");
    const price = Number(it.price || 0);
    const desc = String(it.desc || "");

    if (!sku) return;

    // 1️⃣ Voorraad herstellen
    _revertSoldItem_(sku);

    // 2️⃣ Loggen in retour-sheet
    retourSheet.appendRow([
      now,           // Datum (kolom A)
      returnNo,      // Retournummer (kolom B)
      receiptNo,     // Originele Bon (kolom C)
      sku,           // SKU (kolom D)
      desc,          // Omschrijving (kolom E)
      -Math.abs(price), // Prijs (kolom F) -> negatief op sheet
      reason,        // Reden (kolom G)
      returnTicketUrl // URL (kolom H) (of laat leeg als je wilt)
    ]);
  });

  Logger.log("[RET] Retour succesvol: " + returnNo);

  return {
    ok: true,
    returnNo: returnNo,
    itemsCount: itemsToReturn.length,
    receiptNo: receiptNo,
    returnTicketUrl: returnTicketUrl
  };
}

function generateReturnNumber() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Retouren');
  if (!sh) throw new Error("Tab 'Retouren' ontbreekt.");

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd'
  );

  // Tel bestaande retouren van vandaag
  const lastRow = sh.getLastRow();
  let count = 0;

  if (lastRow > 1) {
    const vals = sh.getRange(2, 2, lastRow - 1, 1).getValues(); // kolom B = retourNo
    vals.forEach(v => {
      if (String(v[0]).includes(today)) count++;
    });
  }

  const seq = String(count + 1).padStart(3, '0');
  return `RT-${today}-${seq}`;
}

/**
 * Herstelt een verkochte SKU in de voorraad.
 * - zet COL.F (verkoopprijs) = backup expected prijs (COL.L)
 * - zet COL.G (verkoopdatum) = ""
 * - zet COL.J (marge) = 0
 */
function _revertSoldItem_(sku) {
  console.log("[RET] _revertSoldItem_ CALLED →", sku);

  const ss = SpreadsheetApp.getActive();
  const sheets = ['Clubs','Sets','Tassen',"Trolley's",'Overig','Diensten'];

  sku = String(sku).trim();
  if (!sku) return;

  for (const name of sheets) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const vals = sh.getDataRange().getValues();

    for (let r = 1; r < vals.length; r++) {
      if (String(vals[r][0]).trim() === sku) {

        const row = r + 1;
        console.log("[RET] SKU FOUND in sheet:", name, "row:", row);

        const expectedBackup = vals[r][11];

        sh.getRange(row, 6).setValue(expectedBackup);
        sh.getRange(row, 7).setValue("");
        sh.getRange(row, 8).setValue("");
        // Stel formules in voor de nieuwe rij
        sh.getRange(row, 9).setFormula('=IF(ISBLANK(F' + row + ');0;F' + row + '-D' + row + ')');
        sh.getRange(row, 10).setFormula('=IF(ISBLANK(G' + row + ');0;G' + row + '-D' + row + ')');

        return;
      }
    }
  }

  throw new Error("SKU niet gevonden in voorraad: " + sku);
}

function isAlreadyReturned_(receiptNo, sku) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Retouren');
  if (!sh) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const vals = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  // A = datum
  // B = returnNo
  // C = receiptNo
  // D = sku

  return vals.some(r =>
    String(r[2]) === String(receiptNo) &&
    String(r[3]) === String(sku)
  );
}

// ====== CONFIG: pas aan naar jouw sheet kolommen ======
const SALES_COLS = {
  receipt: 1, // meestal Bonnummer kolom (bijv. B)
  total:   4  // totaalbedrag kolom (bijv. D)
};

const SALES_LINES_COLS = {
  receipt: 1, // bonnummer kolom
  sku:     2, // sku kolom
  price:   4, // prijs kolom
  qty:     5, // qty kolom
  subtotal:6  // subtotaal kolom
};

// Zoek rij-index (1-based) van value in een kolom
function _findRowByValue_(sh, col, value) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  const needle = String(value).trim();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === needle) return i + 2;
  }
  return 0;
}

// Zoek alle rijen waar receiptNo + sku matcht (kan meerdere zijn)
function _findSalesLineRows_(sh, receiptNo, sku) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, Math.max(
    SALES_LINES_COLS.receipt,
    SALES_LINES_COLS.sku,
    SALES_LINES_COLS.price,
    SALES_LINES_COLS.qty,
    SALES_LINES_COLS.subtotal
  )).getValues();

  const rNeedle = String(receiptNo).trim();
  const sNeedle = String(sku).trim();

  const rows = [];
  for (let i = 0; i < data.length; i++) {
    const r = String(data[i][SALES_LINES_COLS.receipt - 1]).trim();
    const s = String(data[i][SALES_LINES_COLS.sku - 1]).trim();
    if (r === rNeedle && s === sNeedle) rows.push(i + 2);
  }
  return rows;
}

// Pas totaal in Sales aan: nieuw = oud - refundTotal
function _applyReturnToSalesTotal_(receiptNo, refundTotal) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Sales');
  if (!sh) throw new Error("Tab 'Sales' ontbreekt.");

  const row = _findRowByValue_(sh, SALES_COLS.receipt, receiptNo);
  if (!row) throw new Error(`Sales: bon ${receiptNo} niet gevonden.`);

  const oldTotal = Number(sh.getRange(row, SALES_COLS.total).getValue() || 0);
  const newTotal = oldTotal - Number(refundTotal || 0);

  sh.getRange(row, SALES_COLS.total).setValue(newTotal);
  return { row, oldTotal, newTotal };
}

// Zet de bijbehorende Sales_Lines regel negatief
function _applyReturnToSalesLines_(receiptNo, itemsToReturn) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Sales_Lines');
  if (!sh) throw new Error("Tab 'Sales_Lines' ontbreekt.");

  itemsToReturn.forEach(it => {
    const sku = String(it.sku || '').trim();
    if (!sku) return;

    const rows = _findSalesLineRows_(sh, receiptNo, sku);
    if (!rows.length) throw new Error(`Sales_Lines: regel niet gevonden voor bon ${receiptNo}, SKU ${sku}.`);

    // Kies de "beste" rij: bij voorkeur eentje die nog positief is
    let targetRow = rows[0];
    for (const r of rows) {
      const subt = Number(sh.getRange(r, SALES_LINES_COLS.subtotal).getValue() || 0);
      if (subt > 0) { targetRow = r; break; }
    }

    const qty   = Number(sh.getRange(targetRow, SALES_LINES_COLS.qty).getValue() || 1);
    const price = Number(sh.getRange(targetRow, SALES_LINES_COLS.price).getValue() || 0);
    const subt  = Number(sh.getRange(targetRow, SALES_LINES_COLS.subtotal).getValue() || (price * qty));

    // Als jouw regel altijd qty=1 per product: simpel negatief maken
    // (Dit matcht jouw wens: "die artikelregel op -€10")
    sh.getRange(targetRow, SALES_LINES_COLS.price).setValue(-Math.abs(price || it.price || 0));
    sh.getRange(targetRow, SALES_LINES_COLS.subtotal).setValue(-Math.abs(subt || it.price || 0));

    // Optioneel: qty ook negatief (als je dat consistent wil)
    // sh.getRange(targetRow, SALES_LINES_COLS.qty).setValue(-Math.abs(qty));
  });
}

function apiGetRecentReturns(limit) {
  limit = limit || 20;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Retouren');
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh
    .getRange(2, 1, lastRow - 1, 8)
    .getValues()
    .reverse()
    .slice(0, limit);

  return values.map(r => ({
    date: r[0],
    returnNo: r[1],
    receiptNo: r[2],
    sku: r[3],
    description: r[4],
    amount: Number(r[5]) || 0,
    reason: r[6],
    url: r[7]
  }));
}

function debugRecentReturns() {
  const res = apiGetRecentReturns(20);
  Logger.log(JSON.stringify(res, null, 2));
}

function apiGetRecentReturnsSafe(limit) {
  try {
    // 1) haal data op via je bestaande functie
    const res = apiGetRecentReturns(limit);

    // 2) force JSON-serialiseerbaar (cruciaal voor google.script.run)
    return JSON.parse(JSON.stringify(res || []));
  } catch (e) {
    // force failureHandler (zodat je nooit weer "null" stilletjes krijgt)
    throw new Error('apiGetRecentReturnsSafe failed: ' + (e && e.message ? e.message : e));
  }
}


