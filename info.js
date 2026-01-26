/*********************************
 * INFO DB CONFIG
 *********************************/

const INFO_DB_SPREADSHEET_ID = '1Y0MAN_a69IhFI7RGm6geVaVI21u1YOl7UPGen5F-IOM';

let INFO_CACHE = {
  shafts: null,
  lengthsStandard: null,
  lengthsByBrand: null,
  lofts: null,
};

let INFO_RUNTIME_CACHE = {
  shaftsIndex: null,
};

function getInfoSheetData_(sheetName) {
  const ss = SpreadsheetApp.openById(INFO_DB_SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return { headers: [], rows: [] };

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return { headers: [], rows: [] };

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return { headers, rows };
}

function getShaftsIndex_() {
  try {
    if (INFO_CACHE && INFO_CACHE.shafts) {
      return INFO_CACHE.shafts;
    }

    const master = getInfoSheetData_('Shafts_Master');
    const variants = getInfoSheetData_('Shafts_Variants');

    const masterIdx = {};
    master.rows.forEach(r => {
      const obj = {};
      master.headers.forEach((h, i) => obj[h] = r[i]);
      if (obj.active === false) return;
      if (!obj.shaft_id) return;
      masterIdx[obj.shaft_id] = obj;
    });

    const variantIdx = {};
    variants.rows.forEach(r => {
      const obj = {};
      variants.headers.forEach((h, i) => obj[h] = r[i]);
      if (!obj.shaft_id) return;
      if (!variantIdx[obj.shaft_id]) variantIdx[obj.shaft_id] = [];
      variantIdx[obj.shaft_id].push(obj);
    });

    INFO_CACHE = INFO_CACHE || {};
    INFO_CACHE.shafts = { masterIdx, variantIdx };

    return INFO_CACHE.shafts;
  } catch (e) {
    // HARD fallback → nooit null teruggeven
    return { masterIdx: {}, variantIdx: {} };
  }
}



function apiInfoSearchShafts(query, filters = {}) {
  const q = (query || '').toLowerCase();
  const { masterIdx } = getShaftsIndex_();

  const results = [];

  Object.values(masterIdx).forEach(s => {
    if (filters.category && s.category !== filters.category) return;
    if (filters.material && s.material !== filters.material) return;

    const haystack = (
      s.brand + ' ' +
      s.model + ' ' +
      (s.keywords || '')
    ).toLowerCase();

    if (q && !haystack.includes(q)) return;

    results.push({
      shaft_id: s.shaft_id,
      brand: s.brand,
      model: s.model,
      material: s.material,
      category: s.category
    });
  });

  return { ok: true, items: results };
}

function apiInfoGetShaftDetail(shaftId) {
  try {
    const { masterIdx, variantIdx } = getShaftsIndex_();

    const master = masterIdx[shaftId];
    if (!master) {
      return JSON.parse(JSON.stringify({
        ok: false,
        error: 'Shaft niet gevonden'
      }));
    }

    const variants = variantIdx[shaftId] || [];

    // 🔑 JSON-safe return (belangrijk!)
    return JSON.parse(JSON.stringify({
      ok: true,
      master,
      variants
    }));

  } catch (e) {
    return JSON.parse(JSON.stringify({
      ok: false,
      error: String(e.message || e)
    }));
  }
}

function apiInfoGetAllShaftsWithDetails() {
  try {
    const { masterIdx, variantIdx } = getShaftsIndex_();

    const items = Object.values(masterIdx).map(m => ({
      master: m,
      variants: variantIdx[m.shaft_id] || []
    }));

    return JSON.parse(JSON.stringify({
      ok: true,
      items
    }));
  } catch (e) {
    return JSON.parse(JSON.stringify({
      ok: false,
      error: String(e.message || e)
    }));
  }
}


function debug_Info_SearchShafts() {
  const res = apiInfoSearchShafts('dynamic', {});
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function debug_Info_GetShaftDetail() {
  const shaftId = 'TT_DG_STEEL_IRON';
  const res = apiInfoGetShaftDetail(shaftId);
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

function apiInfoGetStandardLengths() {
  const { headers, rows } = getInfoSheetData_('Lengths_Standards');
  const items = rows.map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
  return { ok: true, items };
}

function apiInfoGetBrandLengths(brand) {
  const { headers, rows } = getInfoSheetData_('Lengths_By_Brand');
  const items = [];

  rows.forEach(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    if (brand && o.brand !== brand) return;
    items.push(o);
  });

  return { ok: true, items };
}

function apiInfoCalculateFit(inputType, value) {
  const calc = getInfoSheetData_('Lengths_Fit_Calculator');
  let adjust = 0;

  calc.rows.forEach(r => {
    const o = {};
    calc.headers.forEach((h, i) => o[h] = r[i]);

    if (o.input_type !== inputType) return;
    if (value >= o.range_min && value <= o.range_max) {
      adjust = Number(o.recommended_adjust_in) || 0;
    }
  });

  return { ok: true, adjust };
}

function apiInfoGetLofts(brand) {
  const { headers, rows } = getInfoSheetData_('Lofts');
  const items = [];

  rows.forEach(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);

    if (brand && o.brand !== brand) return;
    items.push(o);
  });

  return { ok: true, items };
}

function apiInfoAddShaftMaster(data) {
  const ss = SpreadsheetApp.openById(INFO_DB_SPREADSHEET_ID);
  const sh = ss.getSheetByName('Shafts_Master');
  if (!sh) return { ok: false, error: 'Shafts_Master not found' };

  const existing = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat();
  if (existing.includes(data.shaft_id)) {
    return { ok: false, error: 'shaft_id bestaat al' };
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    if (h === 'created_at') return new Date();
    return data[h] ?? '';
  });

  sh.appendRow(row);
  resetInfoCache_();
  return { ok: true };
}


function apiInfoAddShaftVariant(data) {
  const ss = SpreadsheetApp.openById(INFO_DB_SPREADSHEET_ID);
  const sh = ss.getSheetByName('Shafts_Variants');
  if (!sh) return { ok: false, error: 'Shafts_Variants not found' };

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const idxShaft = headers.indexOf('shaft_id');
  const idxFlex = headers.indexOf('flex_label');

  for (const r of rows) {
    if (r[idxShaft] === data.shaft_id && r[idxFlex] === data.flex_label) {
      return { ok: false, error: 'Variant bestaat al' };
    }
  }

  const row = headers.map(h => data[h] ?? '');
  sh.appendRow(row);
  resetInfoCache_();
  return { ok: true };
}


function resetInfoCache_() {
  INFO_CACHE = {
    shafts: null,
    lengthsStandard: null,
    lengthsByBrand: null,
    lofts: null,
  };
}

function apiInfoGenerateShaftSummary(shaftId) {
  try {
    const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!key) {
      return JSON.parse(JSON.stringify({
        ok: false,
        error: 'OPENAI_API_KEY ontbreekt'
      }));
    }

    // 1️⃣ Haal master + varianten op (zelfde veilige manier)
    const masterData = getInfoSheetData_('Shafts_Master');
    const variantData = getInfoSheetData_('Shafts_Variants');

    let master = null;
    masterData.rows.forEach(r => {
      const o = {};
      masterData.headers.forEach((h, i) => o[h] = r[i]);
      if (o.shaft_id === shaftId) master = o;
    });

    if (!master) {
      return JSON.parse(JSON.stringify({
        ok: false,
        error: 'Shaft niet gevonden'
      }));
    }

    const variants = [];
    variantData.rows.forEach(r => {
      const o = {};
      variantData.headers.forEach((h, i) => o[h] = r[i]);
      if (o.shaft_id === shaftId) variants.push(o);
    });

    // 2️⃣ Prompt bouwen (simpel & kort)
    const prompt = `
    Je bent een golfclubfitter die klanten in de winkel neutraal uitlegt wat een shaft bijzonder maakt.
    Schrijf in het Nederlands EXACT 2–3 zinnen, zonder verkooppraat, zonder vage claims.

    Doel: Leg uit wat het model/nummer betekent (bijv. 105/120 of Red/Black), en wat je praktisch merkt t.o.v. een nabije variant binnen dezelfde familie.
    Als je een detail NIET zeker weet, zoek dan op de website van de fabrikant naar het model. Als je het daar ook niet kunt vinden, schrijf dan letterlijk: "Onbekend" voor dat detail (niet gokken, niet opvullen).

    Vereisten:
    - Noem 1x wat het getal/label in de modelnaam betekent (bijv. gewichtsklasse in gram: ja/nee, en hoe dat uitwerkt).
    - Noem 1x het verwachte effect op gevoel/tempo en ballflight (launch/spin) in simpele woorden.
    - Vermijd woorden als "hoogwaardig", "optimale", "perfect", "uitstekend".

    Context (alleen dit gebruiken):
    Merk: ${master.brand}
    Model: ${master.model}
    Materiaal: ${master.material}
    Categorie: ${master.category}
    Varianten/flexes: ${variants.map(v => v.flex_label).join(', ')}

    Output: alleen de 2–3 zinnen, geen opsomming, geen titel.
    `;


    // 3️⃣ OpenAI call
    const response = UrlFetchApp.fetch(
      'https://api.openai.com/v1/chat/completions',
      {
        method: 'post',
        contentType: 'application/json',
        headers: {
          Authorization: 'Bearer ' + key
        },
        payload: JSON.stringify({
          model: 'gpt-4o-mini',
          messages: [
            { role: 'user', content: prompt }
          ],
          max_tokens: 120
        })
      }
    );

    const json = JSON.parse(response.getContentText());
    const summary = json.choices?.[0]?.message?.content || '';

    // 4️⃣ Opslaan in sheet
    const ss = SpreadsheetApp.openById(INFO_DB_SPREADSHEET_ID);
    const sh = ss.getSheetByName('Shafts_Master');
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    const idCol = headers.indexOf('shaft_id') + 1;
    const sumCol = headers.indexOf('ai_summary') + 1;
    const updCol = headers.indexOf('ai_summary_updated_at') + 1;

    const ids = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getValues().flat();
    const idx = ids.indexOf(shaftId);
    if (idx >= 0) {
      const row = idx + 2;
      sh.getRange(row, sumCol).setValue(summary);
      sh.getRange(row, updCol).setValue(new Date());
    }

    return JSON.parse(JSON.stringify({
      ok: true,
      summary
    }));

  } catch (e) {
    return JSON.parse(JSON.stringify({
      ok: false,
      error: String(e.message || e)
    }));
  }
}


