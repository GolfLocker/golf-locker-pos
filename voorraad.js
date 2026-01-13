/***********************
 * VOORRAAD MODULE (v1)
 * Bestand: voorraad.gs
 *
 * Vereist:
 * - hardInitCounterIfNeeded_() en getNextSku_() bestaan in code.gs (MAG NIET gewijzigd worden)
 ***********************/

const STOCKCOUNT_PROP_FILE_ID = 'STOCKCOUNT_FILE_ID';
const STOCKCOUNT_USERCACHE_KEY = 'STOCKCOUNT_ACTIVE_ID';

const VOORRAAD_TABS = [
  'Clubs',
  'Sets',
  'Tassen',
  "Trolley's",
  'Overig',
  'Diensten'
];

// Kolom-mapping (identiek aan code.gs / POS)
const VOORRAAD_COL = {
  sku: 1,        // A
  desc: 2,       // B
  purchase: 3,   // C
  buy: 4,        // D
  party: 5,      // E
  expected: 6,   // F
  sale: 7,       // G
  channel: 15    // O
};


/** ===== Helpers ===== */

function _assert_(cond, msg){
  if (!cond) throw new Error(msg || 'Ongeldige input');
}

function _tz_(){
  return Session.getScriptTimeZone() || 'Europe/Amsterdam';
}

function _fmtDate_(d){
  return Utilities.formatDate(d, _tz_(), "dd-MM-yyyy");
}

/** Vind laatste “echte” rij met SKU in kolom A (zoals troephoek aanpak). */
function _getLastRealRow_(sheet){
  const last = sheet.getLastRow();
  if (last < 1) return 1;
  const values = sheet.getRange(1, 1, last, 1).getValues();
  let lastReal = 1;
  values.forEach((r, i) => {
    const v = r[0];
    if (v !== "" && v !== null) lastReal = i + 1;
  });
  return lastReal;
}

function _getSheetByNameOrThrow_(name){
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Tab niet gevonden: ' + name);
  return sh;
}

/** ===== Voorraad: Inschrijven ===== */

/**
 * payload:
 * {
 *   tab: 'Clubs'|'Sets'|...,
 *   desc: string,
 *   buy: number|string,
 *   expected: number|string,
 *   party?: string,
 *   channel?: string
 * }
 */
function apiVoorraadCreateItem(payload){
  payload = payload || {};
  const tab = String(payload.tab || '').trim();
  const desc = String(payload.desc || '').trim();
  const buy = payload.buy;
  const expected = payload.expected;
  const party = String(payload.party || '').trim();
  const channel = String(payload.channel || '').trim();

  _assert_(tab && (VOORRAAD_TABS || []).includes(tab), 'Kies een geldige categorie/tab.');
  _assert_(desc, 'Omschrijving is verplicht.');
  _assert_(buy !== '' && buy !== null && buy !== undefined, 'Aankoopprijs is verplicht.');
  _assert_(expected !== '' && expected !== null && expected !== undefined, 'Verwachte verkoopprijs is verplicht.');

  // ✅ Gebruik exact dezelfde SKU-teller als code.gs (zonder wijziging)
  hardInitCounterIfNeeded_();
  const sku = getNextSku_();

  const sh = _getSheetByNameOrThrow_(tab);
  const row = _getLastRealRow_(sh) + 1;

  const now = new Date();
  const purchaseDate = _fmtDate_(now);

  // Schrijf alleen de kolommen die we nodig hebben (COL komt uit pos.gs)
  sh.getRange(row, COL.sku).setValue(String(sku));              // A (als tekst)
  sh.getRange(row, COL.desc).setValue(desc);                    // B
  sh.getRange(row, 3).setValue(purchaseDate);                   // C (aankoopdatum)
  sh.getRange(row, COL.buy).setValue(Number(buy) || 0);         // D
  if (party)   sh.getRange(row, COL.party).setValue(party);     // E
  sh.getRange(row, COL.expected).setValue(Number(expected) || 0); // F

  // I: verwachte marge formule (zelfde als elders in je systeem)
  // IF(ISBLANK(F),0,F-D)
  sh.getRange(row, COL.expMargin).setFormulaR1C1('=IF(ISBLANK(RC6);0;RC6-RC4)');

  // L: backup expected = hard value
  sh.getRange(row, COL.backupExp).setValue(Number(expected) || 0);

  // O: channel
  if (channel) sh.getRange(row, COL.channel).setValue(channel);

  SpreadsheetApp.flush();

  return {
    ok: true,
    item: {
      sku: String(sku),
      tab,
      row,
      purchaseDate,
      desc,
      buy: Number(buy) || 0,
      expected: Number(expected) || 0,
      party,
      channel
    }
  };
}

function _getLastRealRow_(sheet) {
  const last = sheet.getLastRow();
  if (last < 1) return 1;

  const values = sheet.getRange(1, 1, last, 1).getValues(); // alleen kolom A (SKU)
  let lastReal = 1;

  values.forEach((r, i) => {
    const v = r[0];
    if (v !== '' && v !== null) {
      lastReal = i + 1;
    }
  });

  return lastReal;
}

/**
 * Laatste N toegevoegde items, tab-overstijgend.
 * We sorteren op SKU (hoog -> laag), omdat dat jouw teller-volgorde is.
 */
function apiVoorraadListRecent(limit) {
  limit = Number(limit || 20);
  if (!limit || limit < 1) limit = 20;
  if (limit > 100) limit = 100;

  const ss = SpreadsheetApp.getActive();
  const out = [];

  (VOORRAAD_TABS || []).forEach(tab => {
    const sh = ss.getSheetByName(tab);
    if (!sh) return;

    const last = sh.getLastRow();
    if (last < 2) return;

    // ✅ Bulk read A..O (snel) + displayValues (zoals je het in Sheets ziet)
    const rng   = sh.getRange(2, 1, last - 1, 15);
    const vals  = rng.getValues();
    const disp  = rng.getDisplayValues();

    for (let i = 0; i < vals.length; i++) {
      const sku = String(disp[i][0] || '').trim(); // A, display (kan nummer->string zijn)
      if (!sku) continue;

      // Gebruik display voor tekst/datum (zodat je niet tegen Date parsing aanloopt)
      const desc        = disp[i][1] || '';  // B
      const purchaseStr = disp[i][2] || '';  // C
      const party       = disp[i][4] || '';  // E
      const channel     = disp[i][14] || ''; // O

      // Gebruik values voor bedragen (blijft nummer, maar kan ook leeg zijn)
      const buy      = vals[i][3]; // D
      const expected = vals[i][5]; // F

      out.push({
        sku,
        tab,
        row: i + 2,
        desc,
        purchaseDate: purchaseStr,
        buy: (buy === '' || buy === null) ? '' : buy,
        expected: (expected === '' || expected === null) ? '' : expected,
        party,
        channel
      });
    }
  });

  // Recent = hoogste SKU eerst (jouw teller is oplopend)
  out.sort((a, b) => (Number(b.sku) || 0) - (Number(a.sku) || 0));

  return { ok: true, items: out.slice(0, limit) };
}

/** ===== Voorraad: Tellen (met apart bestand) ===== */

function _getOrCreateCountFile_(){
  const props = PropertiesService.getDocumentProperties();
  let id = props.getProperty(STOCKCOUNT_PROP_FILE_ID);

  if (id) {
    try {
      const f = SpreadsheetApp.openById(id);
      return f;
    } catch(e){
      // bestaat niet / geen rechten → opnieuw aanmaken
      id = null;
    }
  }

  const f = SpreadsheetApp.create('Golf Locker - Voorraad Tellingen');
  props.setProperty(STOCKCOUNT_PROP_FILE_ID, f.getId());

  // Sheets + headers
  const ss = SpreadsheetApp.openById(f.getId());

  const s1 = ss.getSheets()[0];
  s1.setName('Tellingen');
  s1.getRange(1,1,1,8).setValues([[
    'TellingID','StartedAt','FinishedAt','ExpectedCount','FoundCount','MissingCount','SoldScans','UnknownScans'
  ]]);

  const exp = ss.insertSheet('Expected');
  exp.getRange(1,1,1,8).setValues([[
    'TellingID','SKU','Tab','Row','Omschrijving','Partij','Verkoopprijs','Status'
  ]]);

  const scans = ss.insertSheet('Scans');
  scans.getRange(1,1,1,8).setValues([[
    'TellingID','ScannedAt','SKU','Status','Tab','Row','Omschrijving','Extra'
  ]]);

  const res = ss.insertSheet('Result');
  res.getRange(1,1,1,8).setValues([[
    'TellingID','Type','SKU','Tab','Row','Omschrijving','Status','Extra'
  ]]);

  return ss;
}

function apiVoorraadTellingStart(){
  const tz = _tz_();
  const now = new Date();
  const id = 'T' + Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss');

  // Snapshot: alle onverkochte items (kolom G leeg)
  const ssMain = SpreadsheetApp.getActive();
  const expectedRows = [];

  (VOORRAAD_TABS || []).forEach(tab => {
    const sh = ssMain.getSheetByName(tab);
    if (!sh) return;

    const last = sh.getLastRow();
    if (last < 2) return;

    const take = last - 1;
    const data = sh.getRange(2, 1, take, 7).getValues(); // A..G is genoeg voor snapshot
    data.forEach((r, i) => {
      const sku = String(r[0] || '').trim();
      if (!sku) return;

      const sale = r[6]; // G
      const isSold = !(sale === '' || sale === null);

      if (!isSold) {
        expectedRows.push([
          id, sku, tab, (2 + i), r[1] || '', r[4] || '', r[6] || '', 'EXPECTED'
        ]);
      }
    });
  });

  const ssCount = _getOrCreateCountFile_();

  // Log telling
  const shT = ssCount.getSheetByName('Tellingen');
  shT.appendRow([id, now.toISOString(), '', expectedRows.length, 0, expectedRows.length, 0, 0]);

  // Save expected snapshot
  const shE = ssCount.getSheetByName('Expected');
  if (expectedRows.length){
    shE.getRange(shE.getLastRow()+1, 1, expectedRows.length, expectedRows[0].length).setValues(expectedRows);
  }

  // Zet active telling in user cache
  CacheService.getUserCache().put(STOCKCOUNT_USERCACHE_KEY, id, 60 * 60 * 6); // 6 uur

  return { ok:true, tellingId:id, expectedCount: expectedRows.length };
}

function apiVoorraadTellingScan(sku){
  const id = CacheService.getUserCache().get(STOCKCOUNT_USERCACHE_KEY);
  _assert_(id, 'Geen actieve telling. Klik eerst op "Start telling".');

  const skuStr = String(sku || '').trim();
  _assert_(skuStr, 'Vul/scan een SKU.');

  // Zoek in alle tabs, pak de rij, bepaal sold status via kolom G
  const ssMain = SpreadsheetApp.getActive();
  let found = null;

  for (let t=0; t<(VOORRAAD_TABS||[]).length; t++){
    const tab = VOORRAAD_TABS[t];
    const sh = ssMain.getSheetByName(tab);
    if (!sh) continue;

    const last = sh.getLastRow();
    if (last < 2) continue;

    // voor performance: zoek in kolom A
    const skus = sh.getRange(2, 1, last-1, 1).getValues();
    for (let i=0; i<skus.length; i++){
      if (String(skus[i][0] || '').trim() === skuStr){
        const row = 2 + i;
        const desc = sh.getRange(row, COL.desc).getValue();
        const sale = sh.getRange(row, COL.sale).getValue();
        const status = (sale === '' || sale === null) ? 'KLOPT' : 'AL_VERKOCHT';
        found = { tab, row, desc, status, sale };
        break;
      }
    }
    if (found) break;
  }

  const ssCount = _getOrCreateCountFile_();
  const shS = ssCount.getSheetByName('Scans');

  const now = new Date();
  if (!found){
    shS.appendRow([id, now.toISOString(), skuStr, 'ONBEKEND', '', '', '', 'Niet gevonden in voorraad-tabs']);
    return { ok:true, tellingId:id, sku:skuStr, status:'ONBEKEND' };
  }

  shS.appendRow([id, now.toISOString(), skuStr, found.status, found.tab, found.row, found.desc, found.sale || '' ]);

  return { ok:true, tellingId:id, sku:skuStr, status:found.status, tab:found.tab, row:found.row, desc:found.desc };
}

function apiVoorraadTellingFinish(opts){
  opts = opts || {};
  const id = CacheService.getUserCache().get(STOCKCOUNT_USERCACHE_KEY);
  _assert_(id, 'Geen actieve telling.');

  // Zorg voor consistente response-structuur
  const emptyArrays = {
    scannedRows: [],
    expectedRows: [],
    resultRows: []
  };

  const ssCount = _getOrCreateCountFile_();
  const shE = ssCount.getSheetByName('Expected');
  const shS = ssCount.getSheetByName('Scans');
  const shR = ssCount.getSheetByName('Result');
  const shT = ssCount.getSheetByName('Tellingen');

  const expVals = shE.getDataRange().getValues();   // incl header
  const scanVals = shS.getDataRange().getValues();  // incl header

  // Expected set
  const expected = [];
  for (let i=1; i<expVals.length; i++){
    if (String(expVals[i][0]) === id){
      expected.push({
        sku: String(expVals[i][1]||'').trim(),
        tab: expVals[i][2]||'',
        row: expVals[i][3]||'',
        desc: expVals[i][4]||''
      });
    }
  }

  // Scanned sets
  const scannedOk = new Set();
  let soldScans = 0, unknownScans = 0;

  // ook één “laatste status per SKU” voor UI
  const scannedRows = [];

  for (let i=1; i<scanVals.length; i++){
    if (String(scanVals[i][0]) !== id) continue;

    const sku = String(scanVals[i][2]||'').trim();
    const status = String(scanVals[i][3]||'').trim();

    scannedRows.push({
      sku,
      status,
      tab: scanVals[i][4]||'',
      row: scanVals[i][5]||'',
      desc: scanVals[i][6]||'',
      extra: scanVals[i][7]||''
    });

    if (status === 'KLOPT') scannedOk.add(sku);
    if (status === 'AL_VERKOCHT') soldScans++;
    if (status === 'ONBEKEND') unknownScans++;
  }

  const expectedSet = new Set(expected.map(x => x.sku));
  const missing = expected.filter(x => !scannedOk.has(x.sku));

  // Schrijf Result regels (append; je kunt later filteren per tellingId)
  const resRows = [];

  scannedRows.forEach(x => {
    resRows.push([id, 'GETELD', x.sku, x.tab, x.row, x.desc, x.status, x.extra]);
  });

  missing.forEach(x => {
    resRows.push([id, 'MISSEND', x.sku, x.tab, x.row, x.desc, 'MISSEND', 'Niet gescand']);
  });

  if (resRows.length){
    shR.getRange(shR.getLastRow()+1, 1, resRows.length, resRows[0].length).setValues(resRows);
  }

  // Update telling header (laatste match)
  const lastRow = shT.getLastRow();
  // zoek telling row
  const tVals = shT.getRange(2,1,Math.max(0,lastRow-1),1).getValues();
  let tRow = -1;
  for (let i=0; i<tVals.length; i++){
    if (String(tVals[i][0]) === id){ tRow = 2+i; break; }
  }
  if (tRow !== -1){
    shT.getRange(tRow, 3).setValue(new Date().toISOString());       // FinishedAt
    shT.getRange(tRow, 5).setValue(scannedOk.size);                 // FoundCount (alleen KLOPT)
    shT.getRange(tRow, 6).setValue(missing.length);                 // MissingCount
    shT.getRange(tRow, 7).setValue(soldScans);
    shT.getRange(tRow, 8).setValue(unknownScans);
  }

  // optioneel: blokkeren als er missing is
  if (missing.length && !opts.force){
    return {
      ok: false,
      tellingId: id,
      missingCount: missing.length,
      missingItems: missing,

      // 👇 ALTIJD AANWEZIG
      scannedRows: [],
      expectedRows: [],
      resultRows: [],

      message: 'Er missen nog artikelen. Kies: verder tellen of toch afronden.'
    };
  }

  // afsluiten
  CacheService.getUserCache().remove(STOCKCOUNT_USERCACHE_KEY);

  return {
    ok: true,
    tellingId: id,
    expectedCount: expected.length,
    foundCount: scannedOk.size,
    missingCount: missing.length,
    missingItems: missing,

    // 👇 ALTIJD AANWEZIG
    scannedRows: scannedRows || [],
    expectedRows: expected || [],
    resultRows: resRows || [],

    soldScans,
    unknownScans
  };
}

/** Alleen voor testen/debug */
function apiVoorraadTellingGetActive(){
  const id = CacheService.getUserCache().get(STOCKCOUNT_USERCACHE_KEY);
  return { ok:true, tellingId: id || '' };
}
