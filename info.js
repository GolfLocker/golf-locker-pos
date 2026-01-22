/*********************************
 * INFO DB CONFIG
 *********************************/

const INFO_DB_SPREADSHEET_ID = '1Y0MAN_a69IhFI7RGm6geVaVI21u1YOl7UPGen5F-IOM';

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
  const master = getInfoSheetData_('Shafts_Master');
  const variants = getInfoSheetData_('Shafts_Variants');

  const masterIdx = {};
  master.rows.forEach(r => {
    const obj = {};
    master.headers.forEach((h, i) => obj[h] = r[i]);
    if (obj.active === false) return;
    masterIdx[obj.shaft_id] = obj;
  });

  const variantIdx = {};
  variants.rows.forEach(r => {
    const obj = {};
    variants.headers.forEach((h, i) => obj[h] = r[i]);
    if (!variantIdx[obj.shaft_id]) variantIdx[obj.shaft_id] = [];
    variantIdx[obj.shaft_id].push(obj);
  });

  return { masterIdx, variantIdx };
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
  const { masterIdx, variantIdx } = getShaftsIndex_();

  const master = masterIdx[shaftId];
  if (!master) return { ok: false, error: 'Shaft not found' };

  return {
    ok: true,
    master,
    variants: variantIdx[shaftId] || []
  };
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

