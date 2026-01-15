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

const VOORRAAD_SEARCH_CACHE_KEY = 'VOORRAAD_SEARCH_INDEX_V1';
const VOORRAAD_SEARCH_CACHE_TTL = 60 * 10; // 10 min
const VOORRAAD_SEARCH_CHUNK_SIZE = 50; // items per chunk


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

function _buildSkuIndex_() {
  const ss = SpreadsheetApp.getActive();
  const index = {};

  (VOORRAAD_TABS || []).forEach(tab => {
    const sh = ss.getSheetByName(tab);
    if (!sh) return;

    const last = sh.getLastRow();
    if (last < 2) return;

    const data = sh.getRange(2, 1, last - 1, 7).getValues(); // A..G

    data.forEach((r, i) => {
      const sku = String(r[0] || '').trim();
      if (!sku) return;

      const sale = r[6];
      index[sku] = {
        tab,
        row: 2 + i,
        desc: r[1] || '',
        sold: !(sale === '' || sale === null),
        sale: sale || ''
      };
    });
  });

  return index;
}


function _getSkuIndex_() {
  const raw = CacheService.getUserCache().get('STOCKCOUNT_SKU_INDEX');
  return raw ? JSON.parse(raw) : null;
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

function _voorraadBuildSearchIndex_(){
  const ss = SpreadsheetApp.getActive();
  const index = [];

  (VOORRAAD_TABS || []).forEach(tab => {
    const sh = ss.getSheetByName(tab);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    // Bulk read A..O
    const range = sh.getRange(2, 1, lastRow - 1, 15);
    const values = range.getValues();
    const display = range.getDisplayValues();

    for (let i = 0; i < values.length; i++){
      const sku = String(display[i][0] || '').trim();
      if (!sku) continue;

      const desc    = display[i][1] || '';
      const party   = display[i][4] || '';
      const channel = display[i][14] || '';

      const buy      = values[i][3];
      const expected = values[i][5];
      const sale     = values[i][6];

      const sold = !(sale === '' || sale === null);
      const row  = i + 2;

      const haystack = (
        sku + ' ' +
        desc + ' ' +
        party + ' ' +
        channel
      ).toLowerCase();

      index.push({
        sku,
        tab,
        row,
        desc,
        buy: Number(buy) || 0,
        expected: Number(expected) || 0,
        party,
        channel,
        sold,
        _hay: haystack
      });
    }
  });

  return index;
}

function debugVoorraadSearchIndex(){
  const idx = _voorraadBuildSearchIndex_();
  Logger.log('Index size: ' + idx.length);
  Logger.log(idx.slice(0, 5));
}

function _voorraadGetSearchIndex_(){
  const cache = CacheService.getScriptCache();

  // check chunk count
  const metaRaw = cache.get(VOORRAAD_SEARCH_CACHE_KEY + '_meta');
  if (metaRaw){
    const meta = JSON.parse(metaRaw);
    const all = [];

    for (let i = 0; i < meta.chunks; i++){
      const part = cache.get(VOORRAAD_SEARCH_CACHE_KEY + '_part_' + i);
      if (part){
        all.push(...JSON.parse(part));
      }
    }
    return all;
  }

  // build fresh
  const index = _voorraadBuildSearchIndex_();
  const chunks = [];

  for (let i = 0; i < index.length; i += VOORRAAD_SEARCH_CHUNK_SIZE){
    chunks.push(index.slice(i, i + VOORRAAD_SEARCH_CHUNK_SIZE));
  }

  // store chunks
  for (let i = 0; i < chunks.length; i++){
    cache.put(
      VOORRAAD_SEARCH_CACHE_KEY + '_part_' + i,
      JSON.stringify(chunks[i]),
      VOORRAAD_SEARCH_CACHE_TTL
    );
  }

  // store meta
  cache.put(
    VOORRAAD_SEARCH_CACHE_KEY + '_meta',
    JSON.stringify({ chunks: chunks.length }),
    VOORRAAD_SEARCH_CACHE_TTL
  );

  return index;
}


function _voorraadInvalidateSearchIndex_(){
  const cache = CacheService.getScriptCache();
  const metaRaw = cache.get(VOORRAAD_SEARCH_CACHE_KEY + '_meta');
  if (!metaRaw) return;

  const meta = JSON.parse(metaRaw);
  for (let i = 0; i < meta.chunks; i++){
    cache.remove(VOORRAAD_SEARCH_CACHE_KEY + '_part_' + i);
  }
  cache.remove(VOORRAAD_SEARCH_CACHE_KEY + '_meta');
}

function _voorraadUpdateItemInCache_(updated){
  const cache = CacheService.getScriptCache();
  const metaRaw = cache.get(VOORRAAD_SEARCH_CACHE_KEY + '_meta');
  if (!metaRaw) return; // geen cache → niks doen

  const meta = JSON.parse(metaRaw);

  for (let i = 0; i < meta.chunks; i++){
    const key = VOORRAAD_SEARCH_CACHE_KEY + '_part_' + i;
    const raw = cache.get(key);
    if (!raw) continue;

    const arr = JSON.parse(raw);
    let changed = false;

    for (let j = 0; j < arr.length; j++){
      const it = arr[j];
      if (it.tab === updated.tab && it.row === updated.row){
        // update fields
        it.desc     = updated.desc;
        it.buy      = updated.buy;
        it.expected = updated.expected;
        it.party    = updated.party;
        it.channel  = updated.channel;

        // rebuild haystack
        it._hay = (
          it.sku + ' ' +
          it.desc + ' ' +
          it.party + ' ' +
          it.channel
        ).toLowerCase();

        changed = true;
        break;
      }
    }

    if (changed){
      cache.put(key, JSON.stringify(arr), VOORRAAD_SEARCH_CACHE_TTL);
      return; // klaar
    }
  }
}

function debugVoorraadSearchCache(){
  const t0 = Date.now();
  const idx1 = _voorraadGetSearchIndex_();
  const t1 = Date.now();

  const idx2 = _voorraadGetSearchIndex_();
  const t2 = Date.now();

  Logger.log('First load: ' + (t1 - t0) + ' ms');
  Logger.log('Second load: ' + (t2 - t1) + ' ms');
  Logger.log('Index size: ' + idx2.length);
}

function apiVoorraadSearch(payload){
  payload = payload || {};

  const q = String(payload.q || '').trim().toLowerCase();
  if (!q) return { ok: true, items: [] };

  const status = String(payload.status || 'ALL');
  const tab    = String(payload.tab || 'ALL');
  const limit  = Math.min(Number(payload.limit || 50), 100);

  const index = _voorraadGetSearchIndex_();
  const results = [];

  for (let i = 0; i < index.length; i++){
    const it = index[i];

    if (tab !== 'ALL' && it.tab !== tab) continue;

    if (status === 'FREE' && it.sold) continue;
    if (status === 'SOLD' && !it.sold) continue;

    if (!it._hay.includes(q)) continue;

    results.push({
      sku: it.sku,
      tab: it.tab,
      row: it.row,
      desc: it.desc,
      buy: it.buy,
      expected: it.expected,
      sold: it.sold,
      party: it.party,
      channel: it.channel
    });

    if (results.length >= limit) break;
  }

  return { ok: true, items: results };
}

function debugVoorraadSearch(){
  const t0 = Date.now();
  const res = apiVoorraadSearch({
    q: 'driver',
    status: 'ALL',
    tab: 'ALL',
    limit: 50
  });
  const t1 = Date.now();

  Logger.log('Search ms: ' + (t1 - t0));
  Logger.log('Results: ' + res.items.length);
}

function apiVoorraadGetItem(payload){
  payload = payload || {};
  const tab = String(payload.tab || '').trim();
  const row = Number(payload.row || 0);

  if (!tab) return { ok:false, error:'Tab ontbreekt' };
  if (!row || row < 2) return { ok:false, error:'Rij ongeldig' };

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(tab);
  if (!sh) return { ok:false, error:'Tab niet gevonden: ' + tab };

  // lees A..O
  const range = sh.getRange(row, 1, 1, 15);
  const values = range.getValues()[0];
  const display = range.getDisplayValues()[0];

  const sku = String(display[0] || '').trim();
  if (!sku) return { ok:false, error:'Geen SKU op deze rij' };

  // kolommen (0-based in arrays)
  const desc    = display[1]  || '';   // B
  const buy     = values[3]   || 0;    // D
  const party   = display[4]  || '';   // E
  const expected= values[5]   || 0;    // F
  const sale    = values[6];           // G
  const channel = display[14] || '';   // O

  const sold = !(sale === '' || sale === null);

  return {
    ok: true,
    item: {
      sku: sku,
      tab: tab,
      row: row,
      desc: desc,
      buy: Number(buy) || 0,
      expected: Number(expected) || 0,
      party: party,
      channel: channel,
      sold: sold
    }
  };
}

function apiVoorraadUpdateItem(payload){
  payload = payload || {};

  const tab = String(payload.tab || '').trim();
  const row = Number(payload.row || 0);

  if (!tab) return { ok:false, error:'Tab ontbreekt' };
  if (!row || row < 2) return { ok:false, error:'Rij ongeldig' };

  const desc     = String(payload.desc || '').trim();
  const party    = String(payload.party || '').trim();
  const channel  = String(payload.channel || '').trim();

  const buy      = Number(payload.buy);
  const expected = Number(payload.expected);

  if (!desc) return { ok:false, error:'Omschrijving is verplicht' };
  if (isNaN(buy)) return { ok:false, error:'Inkoop is ongeldig' };
  if (isNaN(expected)) return { ok:false, error:'Verwacht is ongeldig' };

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(tab);
  if (!sh) return { ok:false, error:'Tab niet gevonden: ' + tab };

  // schrijf waarden
  sh.getRange(row, 2).setValue(desc);      // B omschrijving
  sh.getRange(row, 4).setValue(buy);       // D inkoop
  sh.getRange(row, 5).setValue(party);     // E partij
  sh.getRange(row, 6).setValue(expected);  // F verwacht
  sh.getRange(row, 15).setValue(channel);  // O kanaal

  SpreadsheetApp.flush();

  // zoek-index invalidaten
  // zoek-index gedeeltelijk bijwerken
  _voorraadUpdateItemInCache_({
    tab: tab,
    row: row,
    desc: desc,
    buy: buy,
    expected: expected,
    party: party,
    channel: channel
  });


  return { ok:true };
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

function _initTellingSpreadsheet_(ss){
  const sheets = ss.getSheets();

  // 1️⃣ Gebruik eerste sheet als Tellingen
  const shT = sheets[0];
  shT.setName('Tellingen');
  shT.clearContents();
  shT.getRange(1,1,1,9).setValues([[
    'TellingID','StartedAt','FinishedAt','ExpectedCount',
    'FoundCount','MissingCount','SoldScans','UnknownScans',
    'FileId'
  ]]);

  // 2️⃣ Verwijder eventuele extra sheets (veilig)
  for (let i = sheets.length - 1; i > 0; i--){
    ss.deleteSheet(sheets[i]);
  }

  // 3️⃣ Voeg overige tabs toe
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
}



function _getVoorraadTellingFolder_(){
  const ROOT_NAME = 'Golf Locker B.V.';
  const YEAR_NAME = '2026';
  const TELLING_NAME = 'Voorraadtellingen';

  function getOrCreate(parent, name){
    const it = parent.getFoldersByName(name);
    return it.hasNext() ? it.next() : parent.createFolder(name);
  }

  const root = DriveApp.getFoldersByName(ROOT_NAME).hasNext()
    ? DriveApp.getFoldersByName(ROOT_NAME).next()
    : DriveApp.createFolder(ROOT_NAME);

  const year = getOrCreate(root, YEAR_NAME);
  const telling = getOrCreate(year, TELLING_NAME);

  return telling;
}

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

  const folder = _getVoorraadTellingFolder_();

  const dateStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const fileName = `Golf Locker Voorraadtelling ${dateStr}`;

  const ssCount = SpreadsheetApp.create(fileName);
  DriveApp.getFileById(ssCount.getId()).moveTo(folder);

  // 👇 ESSENTIEEL
  _initTellingSpreadsheet_(ssCount);


  // Log telling
  const shT = ssCount.getSheetByName('Tellingen');
  shT.appendRow([id, now.toISOString(), '', expectedRows.length, 0, expectedRows.length, 0, 0, ssCount.getId() ]);

  // Save expected snapshot
  const shE = ssCount.getSheetByName('Expected');
  if (expectedRows.length){
    shE.getRange(shE.getLastRow()+1, 1, expectedRows.length, expectedRows[0].length).setValues(expectedRows);
  }

  // Zet active telling in user cache
  CacheService.getUserCache().put(STOCKCOUNT_USERCACHE_KEY, id, 60 * 60 * 6); // 6 uur
  CacheService.getUserCache().put(STOCKCOUNT_USERCACHE_KEY + '_FILE', ssCount.getId(), 60 * 60 * 6);

  const skuIndex = _buildSkuIndex_();
  CacheService.getUserCache().put(
    'STOCKCOUNT_SKU_INDEX',
    JSON.stringify(skuIndex),
    60 * 60 * 6
  );


  return { ok:true, tellingId:id, expectedCount: expectedRows.length };
}

function apiVoorraadTellingScan(sku){
  const id = CacheService.getUserCache().get(STOCKCOUNT_USERCACHE_KEY);
  _assert_(id, 'Geen actieve telling. Klik eerst op "Start telling".');

  const skuStr = String(sku || '').trim();
  _assert_(skuStr, 'Vul/scan een SKU.');

  // Zoek in alle tabs, pak de rij, bepaal sold status via kolom G
  // 🔥 SKU lookup via cache (8.1D)
  const skuIndex = _getSkuIndex_();
  _assert_(skuIndex, 'SKU-index niet gevonden. Start de telling opnieuw.');

  const found = skuIndex[skuStr] || null;

  const fileId = CacheService.getUserCache().get(STOCKCOUNT_USERCACHE_KEY + '_FILE');
  _assert_(fileId, 'Tellingbestand niet gevonden.');

  const ssCount = SpreadsheetApp.openById(fileId);
  const shS = ssCount.getSheetByName('Scans');


  const now = new Date();
  if (!found){
    shS.appendRow([id, now.toISOString(), skuStr, 'ONBEKEND', '', '', '', 'Niet gevonden in voorraad-tabs']);
    return { ok:true, tellingId:id, sku:skuStr, status:'ONBEKEND' };
  }

  const status = found.sold ? 'AL_VERKOCHT' : 'KLOPT';

  shS.appendRow([
    id,
    now.toISOString(),
    skuStr,
    status,
    found.tab,
    found.row,
    found.desc,
    found.sale || ''
  ]);

  return {
    ok: true,
    tellingId: id,
    sku: skuStr,
    status,
    tab: found.tab,
    row: found.row,
    desc: found.desc
  };
}


function apiVoorraadTellingFinish(opts){
  opts = opts || {};
  const id = CacheService.getUserCache().get(STOCKCOUNT_USERCACHE_KEY);
  _assert_(id, 'Geen actieve telling.');

  const dbg = { idFromCache: id };

  // Zorg voor consistente response-structuur
  const emptyArrays = {
    scannedRows: [],
    expectedRows: [],
    resultRows: []
  };

  const fileId = CacheService.getUserCache().get(STOCKCOUNT_USERCACHE_KEY + '_FILE');
  _assert_(fileId, 'Tellingbestand niet gevonden.');

  const ssCount = SpreadsheetApp.openById(fileId);

  const shE = ssCount.getSheetByName('Expected');
  const shS = ssCount.getSheetByName('Scans');
  const shR = ssCount.getSheetByName('Result');
  const shT = ssCount.getSheetByName('Tellingen');

  const expVals = shE.getDataRange().getValues();   // incl header
  const scanVals = shS.getDataRange().getValues();  // incl header
  dbg.expectedRowsTotal = expVals.length - 1;


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
      debug: dbg,


      // 👇 ALTIJD AANWEZIG
      scannedRows: [],
      expectedRows: [],
      resultRows: [],

      message: 'Er missen nog artikelen. Kies: verder tellen of toch afronden.'
    };
  }

  // afsluiten
  const uc = CacheService.getUserCache();
  uc.remove(STOCKCOUNT_USERCACHE_KEY);
  uc.remove(STOCKCOUNT_USERCACHE_KEY + '_FILE');
  uc.remove('STOCKCOUNT_SKU_INDEX');

  return {
    ok: true,
    tellingId: id,
    expectedCount: expected.length,
    foundCount: scannedOk.size,
    missingCount: missing.length,
    missingItems: missing,
    debug: dbg,
    scannedRows: scannedRows || [],
    expectedRows: expected || [],
    resultRows: resRows || [],
    soldScans,
    unknownScans
  };
}

function apiVoorraadTellingReceipt(tellingId){
  _assert_(tellingId, 'TellingID ontbreekt.');

  // Zoek bestand-ID via Tellingen sheet
  const ssMain = SpreadsheetApp.getActive();
  const tellingFolder = _getVoorraadTellingFolder_();

  // zoek alle telling-bestanden (we weten: 1 telling = 1 file)
  const files = tellingFolder.getFiles();
  let ss = null;

  while (files.hasNext()){
    const f = files.next();
    const tmp = SpreadsheetApp.openById(f.getId());
    const shT = tmp.getSheetByName('Tellingen');
    if (!shT) continue;

    const vals = shT.getRange(2,1,shT.getLastRow()-1,1).getValues();
    if (vals.flat().includes(tellingId)){
      ss = tmp;
      break;
    }
  }

  _assert_(ss, 'Tellingbestand niet gevonden.');

  const shR = ss.getSheetByName('Result');
  const shT = ss.getSheetByName('Tellingen');

  _assert_(shR && shT, 'Result/Tellingen sheet ontbreekt.');

  const resVals = shR.getDataRange().getValues();
  const tVals   = shT.getDataRange().getValues();

  // --- header info ---
  let header = null;
  for (let i = 1; i < tVals.length; i++){
    if (String(tVals[i][0]) === tellingId){
      header = {
        tellingId,
        startedAt: tVals[i][1],
        finishedAt: tVals[i][2],
        expectedCount: Number(tVals[i][3]) || 0,
        foundCount: Number(tVals[i][4]) || 0,
        missingCount: Number(tVals[i][5]) || 0
      };
      break;
    }
  }
  _assert_(header, 'Telling niet gevonden.');

  // --- result regels ---
  const seen = new Set();
  const counted = [];
  const missing = [];
  const other   = [];

  for (let i = 1; i < resVals.length; i++){
    const r = resVals[i];
    if (String(r[0]) !== tellingId) continue;

    const sku = String(r[2] || '').trim();
    if (seen.has(sku)) continue;   // 👈 DEDUPE
    seen.add(sku);

    const row = {
      sku,
      tab: r[3] || '',
      desc: r[5] || '',
      status: r[6] || '',
      extra: r[7] || ''
    };

    if (r[1] === 'GETELD' && row.status === 'KLOPT'){
      counted.push(row);
    } else if (r[1] === 'MISSEND'){
      missing.push(row);
    } else {
      other.push(row);
    }
  }

  return {
    ok: true,
    header,
    counted,
    missing,
    other
  };
}

/** Alleen voor testen/debug */
function apiVoorraadTellingGetActive(){
  const id = CacheService.getUserCache().get(STOCKCOUNT_USERCACHE_KEY);
  return { ok:true, tellingId: id || '' };
}