/**
 * =========================
 *  GENERATOR CONFIGURATIE
 * =========================
 * Hier bepaal je welke SKU's worden gebruikt om
 * automatisch nieuwe artikelen te genereren.
 *
 * Format:
 * "BASIS-SKU": {
 *    tab: "Naam van tabblad",
 *    prefix: "1569-",
 *    description: "Omschrijving"
 * }
 */
const GENERATOR_SKUS = {
  "1569": {
    tab: "Overig",
    prefix: "1569+",
    description: "Losse club",
    price: "4"
  },
  "1907": {
    tab: "Overig",
    prefix: "1907+",
    description: "Startersetje Ballen, Tees, etc"
  },
  "1908": {
    tab: "Overig",
    prefix: "1908+",
    description: "Paraplu"
  },
  "1911": {
    tab: "Diensten",
    prefix: "1911+",
    description: "Grip vervangen",
    price: "10"
  },
  "1912": {
    tab: "Diensten",
    prefix: "1912+",
    description: "Club verlengen"
  },
    "GIFTCARD": {
    tab: "Overig",
    prefix: "GIFTCARD+",
    description: "Cadeaubon"
  },
      "VERZENDING": {
    tab: "Overig",
    prefix: "SHIP5",
    description: "Verzending",
    price: "14.5"
  },
  "2032": {
    tab: "Diensten",
    prefix: "2032+",
    description: "Reparatie",
    price: "20"
  }

  // Extra voorbeelden – je kunt ze later zelf aanzetten / toevoegen:
  // "2000": { tab: "Overig", prefix: "2000-", description: "Losse shaft" },
  // "8888": { tab: "Overig", prefix: "8888-", description: "Goedkoop item" },
};


/**
 * =========================
 *  ITEM GENERATOR
 * =========================
 * Maakt een nieuwe regel aan voor een generator-SKU.
 * - SKU altijd als TEXT (met verborgen apostrof) → geen datum-conversie.
 * - Schrijft in juiste tab.
 * - Retourneert POS-ready object.
 */
function createGeneratedItem(baseSku) {
  const cfg = GENERATOR_SKUS[baseSku];
  if (!cfg) throw new Error("Onbekende generator-SKU: " + baseSku);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(cfg.tab);
  if (!sheet) throw new Error("Tab niet gevonden: " + cfg.tab);

  // 1: Bepaal nieuwe SKU op basis van prefix (bijv. 1569- → 1569-1, 1569-2, ...)
  const nextSku = getNextGeneratedSku(sheet, cfg.prefix);

  // Force TEXT in Google Sheets → voorkomt datum-bug
  const skuForSheet = "" + nextSku;

  // 2: Datum
  const today = new Date();
  const tz = Session.getScriptTimeZone() || "Europe/Amsterdam";
  const formattedDate = Utilities.formatDate(today, tz, "dd-MM-yyyy");

  // 3: Nieuwe regel (pas aan op jouw kolom-structuur van Overig)
  // Optionele vaste prijs uit config (fallback = 0)
  const fixedPrice = Number(cfg.price || 0);
  const row = [
    skuForSheet,        // A: SKU (TEXT)
    cfg.description,    // B: Omschrijving
    formattedDate,      // C: Aankoopdatum
    0,                  // D: Inkoopprijs
    "",                 // E: Partij
    fixedPrice,         // F: Verwachte verkoopprijs (optioneel)
    "",                 // G: Verkoopprijs
    "",                 // H: Verkoopdatum
    "",                 // I: Verwachte marge
    "",                 // J: Marge
    cfg.description     // K: Opmerking
  ];

  // 4: Vind laatste echte rij op basis van kolom A
  const lastRow = getLastRealRow(sheet);

  // 5: Schrijf nieuwe regel direct onder de laatste
  sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);

  // 6: Cache ongeldig maken (optioneel, maar lichtgewicht)
  // We hoeven hier géén nieuwe index op te bouwen,
  // omdat generator-SKU's niet via findBySku() lopen.
  invalidateSkuIndex_();

  // 7: Retour: info die we direct in de POS-cart gebruiken
  return {
    sku: nextSku,             // zonder apostrof
    description: cfg.description,
    expected: fixedPrice,
    party: "",
    row: lastRow + 1,
    sheetName: cfg.tab,

    // 🔑 marker: komt uit troephoek generator
    _source: 'TROEPHOEK'
  };
}


/**
 * Vind de laatste echte rij met data in kolom A.
 */
function getLastRealRow(sheet) {
  const last = sheet.getLastRow();
  if (last < 1) return 1;

  const values = sheet.getRange(1, 1, last, 1).getValues();
  let lastReal = 1;

  values.forEach((row, i) => {
    const v = row[0];
    if (v !== "" && v !== null) {
      lastReal = i + 1;
    }
  });

  return lastReal;
}


/**
 * Vind de volgende SKU voor een gegeven prefix.
 * Voorbeeld:
 *  prefix = "1569-"
 *  bestaande: 1569-1, 1569-2, 1569-5
 *  → return "1569-6"
 */
function getNextGeneratedSku(sheet, prefix) {
  const last = sheet.getLastRow();
  if (last < 2) {
    // Alleen header → eerste wordt prefix + "1"
    return prefix + "1";
  }

  const values = sheet.getRange(2, 1, last - 1, 1).getValues();
  let highest = 0;

  // Regex om bv "1569-7" uit prefix "1569-" te herkennen
  const re = new RegExp("^" + prefix.replace("+", "\\+") + "(\\d+)$");

  values.forEach(row => {
    let text = String(row[0] || "").trim();

    // Als de cel als tekst is geschreven met apostrof, verwijder die
    text = text.replace(/^'/, "");

    const m = text.match(re);
    if (m) {
      const num = Number(m[1]);
      if (!isNaN(num) && num > highest) highest = num;
    }
  });

  return prefix + (highest + 1);
}

function apiSaveTroephoekBuyPrices(items) {
  if (!Array.isArray(items) || !items.length) {
    return { ok: true };
  }

  const ss = SpreadsheetApp.getActive();

  items.forEach(it => {
    if (!it || !it.sheetName || !it.row) return;

    const sh = ss.getSheetByName(it.sheetName);
    if (!sh) return;

    // accepteer "12,50", "12.50", "€12,50"
    const raw = String(it.buy || '')
      .replace('€', '')
      .replace(/\s/g, '')
      .replace(',', '.');

    const val = Number(raw);
    if (isNaN(val) || val <= 0) return;

    // kolom D = aankoopprijs
    sh.getRange(Number(it.row), 4).setValue(val);
  });

  return { ok: true };
}

