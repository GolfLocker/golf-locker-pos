/***** CONFIG *****/
const CONFIG = {
  excludedSheets: ['Dashboard','Brandstof','Overige kosten','Info','Berekeningen','POS','Charts','Sales','Sales_Lines','Bon','Overige inkomsten','Missende Clubs','Retouren','Kas', 'Verkocht','Codes','Verhuur'], // deze tabs overslaan voor de voorraad-automatisering
  headerRow: 1,                  // koprij
  skuCol: 1,                     // kolom A = SKU
  triggerColForSku: 2,           // kolom B triggert SKU-aanmaak
  skuBaseline: 1513,             // LAATST gebruikte SKU --> volgende wordt 1514
  propKey: 'LAST_SKU',           // opslag voor teller
  initFlagKey: 'SKU_INIT_DONE'   // 1x initialisatie-flag
};

// vertrekpunt voor afstand in tab Brandstof
const KM_ORIGIN = 'Dorpsstraat 38, Utrecht, Nederland';

// === POS Config ===
const POS = {
  sheet: 'POS',
  input: {
    sku: 'B3',     // SKU invoer
    price: 'B5'    // verkoopprijs invoer
  },
  out: {           // weergavevelden in POS (pas desnoods aan)
    desc:   'B7',   // omschrijving
    partij: 'B8',   // partijnummer
    inkoop: 'B9',   // inkoopprijs
    verwacht:'B10', // verwachte verkoop (F)
    verkocht:'B11', // actuele verkoop (G)
    status:  'B13'  // status/feedback
  },
  // Kolomindexen in je artikelbladen (1=A, 2=B, ...)
  cols: {
    sku: 1,              // A
    desc: 2,             // B
    purchase: 4,         // D (inkoop)
    partij: 5,           // E
    expected: 6,         // F (verwachte verkoop)
    sale: 7,             // G (verkoopprijs)
    saleDate: 8,         // H (verkoopdatum)
    expectedMargin: 9,   // I (verwachte marge)
    backupExpected: 12   // L (backup van verwachte verkoop)
  }
};

/***** HELPERS *****/
function isExcludedSheet_(sheet) {
  const ex = (CONFIG.excludedSheets || []).map(s => String(s).toLowerCase());
  return ex.includes(sheet.getName().toLowerCase());
}

// 1x hard init: zet teller op baseline (1513) zodat eerste uitgifte 1514 is
function hardInitCounterIfNeeded_() {
  const props = PropertiesService.getDocumentProperties();
  const done = props.getProperty(CONFIG.initFlagKey);
  if (!done) {
    props.setProperty(CONFIG.propKey, String(CONFIG.skuBaseline));
    props.setProperty(CONFIG.initFlagKey, '1');
  }
}

// Haal volgende SKU STRIKT uit de teller (we scannen niet door de sheets)
function getNextSku_() {
  const props = PropertiesService.getDocumentProperties();
  let last = Number(props.getProperty(CONFIG.propKey));
  if (isNaN(last) || last < CONFIG.skuBaseline || last > CONFIG.skuBaseline + 1000000) {
    last = CONFIG.skuBaseline;
  }
  last += 1; // volgende
  props.setProperty(CONFIG.propKey, String(last));
  return last;
}

// Afstand (km, 1 decimaal) via Apps Script Maps service
function getKmBetween_(origin, destination) {
  if (!destination) return '';
  try {
    const dir = Maps.newDirectionFinder()
      .setOrigin(origin)
      .setDestination(destination)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .setRegion('nl')
      .getDirections();

    const routes = dir && dir.routes;
    if (!routes || !routes.length) return '';
    const meters = routes[0].legs?.reduce((s, leg) => s + (leg.distance?.value || 0), 0) || 0;
    if (!meters) return '';
    return Math.round((meters / 1000) * 10) / 10; // bv. 12.3 km
  } catch (err) {
    return '';
  }
}

// POS helpers
function toast_(msg){ SpreadsheetApp.getActiveSpreadsheet().toast(msg); }

function clearPosOutput_(posSheet) {
  posSheet
    .getRangeList([
      POS.out.desc,
      POS.out.partij,
      POS.out.inkoop,
      POS.out.verwacht,
      POS.out.verkocht,
      POS.out.status
    ])
    .clearContent();
}

function findItemBySku_(sku) {
  if (!sku) return null;
  const ss = SpreadsheetApp.getActive();
  const skuStr = String(sku).trim();
  for (const sh of ss.getSheets()) {
    const name = sh.getName();
    // sla uitgesloten tabs + POS over
    if (isExcludedSheet_(sh) || name === POS.sheet) continue;

    const last = sh.getLastRow();
    if (last <= CONFIG.headerRow) continue;

    const r = sh.getRange(CONFIG.headerRow + 1, POS.cols.sku, last - CONFIG.headerRow, 1).getValues();
    for (let i = 0; i < r.length; i++) {
      if (String(r[i][0]).trim() === skuStr) {
        return { sheet: sh, row: CONFIG.headerRow + 1 + i };
      }
    }
  }
  return null;
}

/***** MENU / TRIGGERS *****/
function onOpen() {
  hardInitCounterIfNeeded_();
  SpreadsheetApp.getUi()
  .createMenu('Golf Locker')
  .addItem('Reset SKU naar 1513', 'hardResetSkuCounter')
  .addSeparator()
  .addItem('POS: Zoek SKU (B3)', 'posZoekKnop')
  .addItem('POS: Opslaan verkoop', 'posOpslaanKnop')
  .addItem('POS: Annuleer verkoop', 'posAnnuleerKnop')
  .addSeparator()
  .addItem('POS stylen', 'stylePOS_')   // <-- deze toevoegen
  .addToUi();
}

function hardResetSkuCounter() {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(CONFIG.propKey, String(CONFIG.skuBaseline));
  props.setProperty(CONFIG.initFlagKey, '1');
  toast_('SKU teller is gereset naar ' + CONFIG.skuBaseline);
}

/***** HOOFDTRIGGER *****/
function onEdit(e) {
  const sh = e.range.getSheet();
  const sheetName = sh.getName();

  const rowStart = e.range.getRow();
  const colStart = e.range.getColumn();
  const rows = e.range.getNumRows();
  const cols = e.range.getNumColumns();

  /***** BRANDSTOF: vul kolom C (km) zodra kolom B (bestemming) is ingevuld/plakt *****/
  if (sheetName === 'Brandstof') {
    const destCol = 2; // B = bestemming
    const kmCol   = 3; // C = kilometers
    const overlaps = destCol >= colStart && destCol < colStart + cols;

    if (overlaps && rowStart > 1) { // skip koprij
      const destVals = sh.getRange(rowStart, destCol, rows, 1).getValues();
      const out = destVals.map(r => [ r[0] ? getKmBetween_(KM_ORIGIN, r[0]) : '' ]);
      sh.getRange(rowStart, kmCol, rows, 1).setValues(out);
    }
    return; // niets anders draaien op de Brandstof-tab
  }

  // Vanaf hier: overige tabs (Dashboard/Overige kosten/Info/Berekeningen/POS worden uitgesloten)
  if (isExcludedSheet_(sh)) return;

  hardInitCounterIfNeeded_();
  if (rowStart <= CONFIG.headerRow) return;

  /***** AUTO-SKU — ALLEEN als kolom B geraakt wordt *****/
  const overlapsSkuTrigger =
    CONFIG.triggerColForSku >= colStart &&
    CONFIG.triggerColForSku < colStart + cols;

  if (overlapsSkuTrigger) {
    const skuRange = sh.getRange(rowStart, CONFIG.skuCol, rows, 1);           // kolom A
    const bRange   = sh.getRange(rowStart, CONFIG.triggerColForSku, rows, 1); // kolom B
    const skuVals  = skuRange.getValues();
    const bVals    = bRange.getValues();

    let changed = false;
    for (let i = 0; i < rows; i++) {
      const alreadyHasSku = skuVals[i][0] !== "";
      const bFilled       = bVals[i][0] !== "";
      if (!alreadyHasSku && bFilled) {
        skuVals[i][0] = getNextSku_(); // 1514, 1515, ...
        changed = true;
      }
    }
    if (changed) {
      skuRange.setValues(skuVals);
      // skuRange.setNumberFormat('@'); // indien als tekst gewenst
    }
  }

  /***** AUTO-DATUMSTEMPELS *****/
  const now = new Date();

  function stampDateIfNeeded(sourceCol, targetCol) {
    const overlaps = (sourceCol >= colStart && sourceCol < colStart + cols);
    if (!overlaps) return;

    const srcRange = sh.getRange(rowStart, sourceCol, rows, 1);
    const tgtRange = sh.getRange(rowStart, targetCol, rows, 1);
    const srcVals  = srcRange.getValues();
    const tgtVals  = tgtRange.getValues();

    let any = false;
    for (let i = 0; i < rows; i++) {
      const hasSrc = srcVals[i][0] !== "";
      const hasTgt = tgtVals[i][0] !== "";
      if (hasSrc && !hasTgt) {
        tgtVals[i][0] = now; // 1x, vast
        any = true;
      }
    }
    if (any) {
      tgtRange.setValues(tgtVals);
      tgtRange.setNumberFormat("dd-mm-yyyy");
    }
  }

  /***** KOPIEER F --> L ALS HARD WAARDE (bulk) *****/
  if (colStart === 6 && rowStart > CONFIG.headerRow) { 
    const waarden = e.range.getValues(); // alles wat je in F plakt
    const doelRange = sh.getRange(rowStart, 12, waarden.length, 1); // kolom L
    doelRange.setValues(waarden); // zet de waarden hard neer in kolom L
  }

  // B -> C (aankoopdatum)
  stampDateIfNeeded(2, 3);

  // G -> H (verkoopdatum)
  stampDateIfNeeded(7, 8);

  /***** VERWACHTE VERKOOP (F) + VERWACHTE MARGE (I) OP 0 ALS G IS INGEVULD *****/
  const overlapsSold = (7 >= colStart && 7 < colStart + cols);
  if (overlapsSold) {
      // Alleen reageren op single-cell edits (geen bulk)
    const soldRange     = sh.getRange(rowStart, 7, rows, 1); // G
    const expectedRange = sh.getRange(rowStart, 6, rows, 1); // F
    const marginRange   = sh.getRange(rowStart, 9, rows, 1); // I

    const soldVals     = soldRange.getValues();
    const expectedVals = expectedRange.getValues();
    const marginVals   = marginRange.getValues();

    let any = false;

    for (let i = 0; i < rows; i++) {
      const hasSold = soldVals[i][0] !== "";

      if (hasSold) {
        expectedVals[i][0] = 0;
        marginVals[i][0]   = 0;
        any = true;

        // ✅ deze functie voorkomt zelf al dubbele kopieën
        copySoldRowToArchive(sh, rowStart + i);
      }
    }

    if (any) {
      expectedRange.setValues(expectedVals);
      marginRange.setValues(marginVals);
    }
  }
}

/***** POS KNOPFUNCTIES *****/
// 💾 Opslaan: zet prijs in G, datum in H, F & I -> 0, en toon info in POS
function posOpslaanKnop() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pos = ss.getSheetByName(POS.sheet);
  if (!pos) throw new Error('POS tab niet gevonden.');

  clearPosOutput_(pos);

  const sku = pos.getRange(POS.input.sku).getValue();
  const prijs = pos.getRange(POS.input.price).getValue();

  if (!sku) { pos.getRange(POS.out.status).setValue('Vul een SKU in (B3).'); toast_('Vul een SKU in (B3).'); return; }
  if (prijs === '' || prijs === null) { pos.getRange(POS.out.status).setValue('Vul een verkoopprijs in (B5).'); toast_('Vul een verkoopprijs in (B5).'); return; }

  const loc = findItemBySku_(sku);
  if (!loc) { pos.getRange(POS.out.status).setValue('SKU niet gevonden.'); toast_('SKU niet gevonden.'); return; }

  const sh = loc.sheet, row = loc.row, c = POS.cols, now = new Date();

  // Schrijf verkoop en datum + reset verwachte velden
  sh.getRange(row, c.sale).setValue(prijs);
  sh.getRange(row, c.saleDate).setValue(now).setNumberFormat('dd-mm-yyyy');
  sh.getRange(row, c.expected).setValue(0);
  sh.getRange(row, c.expectedMargin).setValue(0);

  // Toon info in POS
  pos.getRange(POS.out.desc).setValue(sh.getRange(row, c.desc).getValue());
  pos.getRange(POS.out.partij).setValue(sh.getRange(row, c.partij).getValue());
  pos.getRange(POS.out.inkoop).setValue(sh.getRange(row, c.purchase).getValue());
  pos.getRange(POS.out.verwacht).setValue(sh.getRange(row, c.expected).getValue());
  pos.getRange(POS.out.verkocht).setValue(prijs);
  pos.getRange(POS.out.status).setValue('✔ Verkoop opgeslagen in ');

  toast_('Verkoop opgeslagen voor SKU ' + sku);
}

function posAnnuleerKnop() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const pos = ss.getSheetByName(POS.sheet);
  if (!pos) throw new Error('POS tab niet gevonden.');

  clearPosOutput_(pos);

  const sku = pos.getRange(POS.input.sku).getValue();
  if (!sku) { pos.getRange(POS.out.status).setValue('Vul een SKU in (B3).'); toast_('Vul een SKU in (B3).'); return; }

  const loc = findItemBySku_(sku);
  if (!loc) { pos.getRange(POS.out.status).setValue('SKU niet gevonden.'); toast_('SKU niet gevonden.'); return; }

  const sh = loc.sheet, row = loc.row, c = POS.cols;

  // 1) Verkoop wissen (G en H)
  sh.getRange(row, c.sale).clearContent();      // G
  sh.getRange(row, c.saleDate).clearContent();  // H

  // 2) Verwachte verkoop (F) terugzetten vanuit backup L (of leegmaken)
  const backup = sh.getRange(row, c.backupExpected).getValue(); // L
  if (backup !== '') sh.getRange(row, c.expected).setValue(backup);
  else sh.getRange(row, c.expected).clearContent();

  // 3) Verwachte marge (I) -> formule herstellen met R1C1:
  //    IF(ISBLANK(F), 0, F - D)  => in R1C1: RC6 (F), RC4 (D)
  sh.getRange(row, c.expectedMargin).setFormulaR1C1('=IF(ISBLANK(RC6);0;RC6-RC4)');

  // 4) POS feedback
  pos.getRange(POS.out.desc).setValue(sh.getRange(row, c.desc).getValue());
  pos.getRange(POS.out.partij).setValue(sh.getRange(row, c.partij).getValue());
  pos.getRange(POS.out.inkoop).setValue(sh.getRange(row, c.purchase).getValue());
  pos.getRange(POS.out.verwacht).setValue(sh.getRange(row, c.expected).getValue());
  pos.getRange(POS.out.verkocht).setValue('');
  pos.getRange(POS.out.status).setValue('Verkoop geannuleerd 🔄');

  toast_('Verkoop geannuleerd en formules hersteld.');
}

function posZoekKnop() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const pos = ss.getSheetByName(POS.sheet);
  if (!pos) throw new Error('POS tab niet gevonden.');

  clearPosOutput_(pos);

  const sku = pos.getRange(POS.input.sku).getValue();
  if (!sku) { pos.getRange(POS.out.status).setValue('Vul een SKU in (B3).'); toast_('Vul een SKU in (B3).'); return; }

  const loc = findItemBySku_(sku);
  if (!loc) { pos.getRange(POS.out.status).setValue('SKU niet gevonden.'); toast_('SKU niet gevonden.'); return; }

  const { sheet: sh, row } = loc;
  const c = POS.cols;

  pos.getRange(POS.out.desc).setValue(sh.getRange(row, c.desc).getValue());
  pos.getRange(POS.out.partij).setValue(sh.getRange(row, c.partij).getValue());
  pos.getRange(POS.out.inkoop).setValue(sh.getRange(row, c.purchase).getValue());
  pos.getRange(POS.out.verwacht).setValue(sh.getRange(row, c.expected).getValue());
  pos.getRange(POS.out.verkocht).setValue(sh.getRange(row, c.sale).getValue());
  pos.getRange(POS.out.status).setValue('Artikel gevonden in: ' + sh.getName() + ' (rij ' + row + ')');

  toast_('Artikel geladen voor SKU ' + sku);
}


function stylePOS_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('POS');
  if (!sh) return;

  // Titel bovenaan
  sh.getRange('A1').setValue('💳 Golf Locker POS')
    .setFontSize(16).setFontWeight('bold').setBackground('#333333').setFontColor('#FFFFFF');

  // Sectie labels
  sh.getRange('A3').setValue('Voer SKU in:').setFontWeight('bold');
  sh.getRange('A4').setValue('Resultaat:').setFontWeight('bold');
  sh.getRange('A6').setValue('Aankoopbedrag:').setFontWeight('bold');

  // Knoppen
  sh.getRange('A8').setValue('🔍 Zoeken').setBackground('#1976d2').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('B8').setValue('💾 Opslaan verkoop').setBackground('#388e3c').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('C8').setValue('↩️ Annuleer verkoop').setBackground('#d32f2f').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  // Optioneel: gridlines verbergen (lokaal voor dit tabblad)
  sh.setHiddenGridlines(true);

  // Kolombreedte wat mooier
  sh.setColumnWidths(1, 3, 150);
}

/***** AUTOMATISCH FILTER: TOON ALLEEN ONVERKOCHTE ITEMS *****/
function applyStockFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const sh of ss.getSheets()) {
    const name = sh.getName();

    // Alleen voorraad-tabs
    if (
      isExcludedSheet_(sh) ||
      name === POS.sheet ||
      name === 'Verkocht'
    ) continue;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= CONFIG.headerRow) continue;

    // Verwijder bestaande filter
    if (sh.getFilter()) {
      sh.getFilter().remove();
    }

    // Zet filter
    const range = sh.getRange(2, 1, lastRow - 2, lastCol);
    range.createFilter();


    const filter = sh.getFilter();

    // Kolom G = verkoopprijs → toon alleen lege cellen
    const criteria = SpreadsheetApp.newFilterCriteria()
      .whenCellEmpty()
      .build();

    filter.setColumnFilterCriteria(POS.cols.sale, criteria);
  }
}

/***** SCAN VOORRAAD: ALS KOLOM G IS GEVULD → KOPIEER NAAR VERKOCHT *****/
function scanSoldItemsToVerkocht() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const soldSheetName = 'Verkocht';

  let soldSheet = ss.getSheetByName(soldSheetName);
  if (!soldSheet) {
    soldSheet = ss.insertSheet(soldSheetName);
  }

  for (const sh of ss.getSheets()) {
    const name = sh.getName();

    // Alleen voorraad-tabs
    if (
      isExcludedSheet_(sh) ||
      name === soldSheetName
    ) continue;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= CONFIG.headerRow) continue;

    // Header 1x
    if (soldSheet.getLastRow() === 0) {
      const header = sh.getRange(1, 1, 1, lastCol).getValues();
      soldSheet.getRange(1, 1, 1, lastCol).setValues(header);
    }

    const data = sh.getRange(3, 1, lastRow - 2, lastCol).getValues(); // vanaf rij 3
    for (let i = 0; i < data.length; i++) {
      const saleValue = data[i][POS.cols.sale - 1]; // kolom G
      if (saleValue === '' || saleValue === null) continue;

      const rowNumber = i + 3;
      const key = 'VERKOCHT_' + sh.getName() + '_' + rowNumber;

      if (props.getProperty(key)) continue;

      soldSheet.appendRow(data[i]);
      props.setProperty(key, new Date().toISOString());
    }
  }
}
