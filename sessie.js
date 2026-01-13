/***********************
 *  SESSIE / KAS LOGICA
 *  Sheet: Kas
 *  Sheet: Sales
 ***********************/

const KAS_SHEET   = 'Kas';
const SALES_SHEET = 'Sales';
const DASHBOARD_SHEET = 'Dashboard';

/**
 * Geeft sessie-id van vandaag: KYYYYMMDD
 */
function getTodaySessionId_() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `K${y}${m}${day}`;
}

/**
 * Haalt sessie van vandaag op (bestaat of niet)
 */
function getTodaySession() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(KAS_SHEET);
  const sid = getTodaySessionId_();

  if (!sh || sh.getLastRow() < 2) {
    return { exists: false };
  }

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === sid) {
      return {
        exists: true,
        row: i + 2,
        sessionId: sid
      };
    }
  }

  return { exists: false };
}

/**
 * Opent sessie van vandaag (maakt aan als die niet bestaat)
 * startCash mag null zijn
 */
function openSession(startCash) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(KAS_SHEET);
  const sid = getTodaySessionId_();
  const now = new Date();

  if (!sh) throw new Error('Kas-sheet niet gevonden');

  const res = getTodaySession();

  // Bestaat al → alleen starttijd/start cash zetten als leeg
  if (res.exists) {
    const row = res.row;

    const startTimeCell = sh.getRange(row, 2);
    const startCashCell = sh.getRange(row, 4);

    if (!startTimeCell.getValue()) {
      startTimeCell.setValue(now);
    }

    if (startCash !== null && startCash !== '' && startCashCell.getValue() === '') {
      startCashCell.setValue(Number(startCash));
    }

    return { ok: true, reused: true };
  }

  // Nieuwe sessie aanmaken
  sh.appendRow([
    sid,                 // Sessie
    now,                 // Starttijd
    '',                  // Eindtijd
    startCash !== null && startCash !== '' ? Number(startCash) : '', // Start Cash
    '',                  // Eind Cash
    '',                  // Verschil
    '',                  // Contante verkopen
    ''                   // Logisch verschil?
  ]);

  return { ok: true, reused: false };
}

/**
 * Sluit sessie van vandaag
 * endCash mag null zijn
 */
function closeSession(endCash) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(KAS_SHEET);
  const sid = getTodaySessionId_();
  const now = new Date();

  if (!sh) throw new Error('Kas-sheet niet gevonden');

  const res = getTodaySession();
  if (!res.exists) {
    throw new Error('Geen actieve sessie voor vandaag');
  }

  const row = res.row;

  const startCash = Number(sh.getRange(row, 4).getValue() || 0);

  // Eindtijd + eind cash
  sh.getRange(row, 3).setValue(now);
  if (endCash !== null && endCash !== '') {
    sh.getRange(row, 5).setValue(Number(endCash));
  }

  // Contante verkopen bepalen (zelfde bron als dashboard)
  let cashSales = 0;

  const salesSh = ss.getSheetByName(SALES_SHEET);
  if (salesSh && salesSh.getLastRow() > 1) {
    const today = new Date();
    today.setHours(0,0,0,0);

    const sales = salesSh.getRange(2,1,salesSh.getLastRow()-1,4).getValues();

    sales.forEach(r => {
      const d   = r[1]; // datum (B)
      const pay = String(r[2] || '').trim().toLowerCase(); // betaalwijze (C)
      const tot = Number(r[3] || 0); // totaal (D)

      if (!(d instanceof Date)) return;

      const day = new Date(d);
      day.setHours(0,0,0,0);

      if (day.getTime() === today.getTime() && pay === 'contant') {
        cashSales += tot;
      }
    });
  }

  sh.getRange(row, 7).setValue(cashSales);

  // Verschil + logische check
  const endCashVal = Number(sh.getRange(row, 5).getValue() || 0);
  const diff = endCashVal - startCash;

  sh.getRange(row, 6).setValue(diff);

  const logical = Math.abs(diff - cashSales) < 0.01 ? 'Logisch' : 'Nee, opnieuw tellen';
  sh.getRange(row, 8).setValue(logical);

  return { ok: true };
}

/**
 * Berekent contante verkopen van vandaag uit Sales
 * Sales:
 *   B = datum
 *   C = betaalwijze
 *   D = totaal
 */
function calculateCashSalesForToday_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SALES_SHEET);

  if (!sh || sh.getLastRow() < 2) return 0;

  const today = new Date();
  today.setHours(0,0,0,0);

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  let sum = 0;

  for (let i = 0; i < data.length; i++) {
    const date = data[i][1]; // kolom B
    const pay  = String(data[i][2] || '').toLowerCase(); // kolom C
    const tot  = Number(data[i][3] || 0); // kolom D

    if (!(date instanceof Date)) continue;

    const d = new Date(date);
    d.setHours(0,0,0,0);

    if (d.getTime() === today.getTime() && pay === 'Contant') {
      sum += tot;
    }
  }

  return sum;
}

function _normText_(v){
  return String(v || '').trim().toLowerCase();
}

/**
 * Zoekt in een sheet naar een cel met exact de tekst (case-insensitive).
 * Return: {row, col} of null
 */
function _findCellByText_(sh, text){
  const target = _normText_(text);
  const rng = sh.getDataRange();
  const vals = rng.getValues();
  for (let r = 0; r < vals.length; r++){
    for (let c = 0; c < vals[r].length; c++){
      if (_normText_(vals[r][c]) === target){
        return { row: r + 1, col: c + 1 };
      }
    }
  }
  return null;
}

/**
 * Leest een 2-koloms tabel (Datum/Label in kol1, bedrag in kol2) onder een blok-titel.
 * Verwacht structuur zoals in je screenshot: titelcel, daaronder header-rij, daaronder data.
 * Return: [{date: Date, value: Number}]
 */
function _readMonthValueTable_(sh, titleText){
  const pos = _findCellByText_(sh, titleText);
  if (!pos) return [];

  // In je screenshot staat de tabel direct onder de titel, met 2 kolommen.
  // We starten 2 rijen onder titel: (titel op rij X), headers op X+1, data vanaf X+2
  const startRow = pos.row + 2;
  const col1 = pos.col;     // "Maand"
  const col2 = pos.col + 1; // "Omzet" / "Bruto winst" etc.

  const lastRow = sh.getLastRow();
  const numRows = Math.max(0, lastRow - startRow + 1);
  if (numRows <= 0) return [];

  const vals = sh.getRange(startRow, col1, numRows, 2).getValues();

  const out = [];
  for (let i = 0; i < vals.length; i++){
    const d = vals[i][0];
    const v = vals[i][1];

    // stop als maand leeg is
    if (!d) break;

    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt)) continue;

    const n = Number(v);
    out.push({ date: dt, value: isNaN(n) ? 0 : n });
  }
  return out;
}

function _ymKey_(dt){
  const y = dt.getFullYear();
  const m = String(dt.getMonth()+1).padStart(2,'0');
  return `${y}-${m}`;
}

function _quarterKey_(dt){
  const y = dt.getFullYear();
  const q = Math.floor(dt.getMonth()/3) + 1;
  return { y, q, key: `${y}-Q${q}` };
}

/**
 * Dashboard data voor Sessie-tab
 */
function getSessionDashboardData(monthKey) {
  try {
    const ss = SpreadsheetApp.getActive();
    const kasSh   = ss.getSheetByName(KAS_SHEET);
    const salesSh = ss.getSheetByName(SALES_SHEET);
    // fallback: huidige maand
    if (!monthKey) {
      const now = new Date();
      monthKey = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
    }

    const sid = getTodaySessionId_();

    const today = new Date();
    today.setHours(0,0,0,0);
    const selectedMonthKey = monthKey ? String(monthKey) : _ymKey_(today);

    // Helpers: altijd JSON-safe
    const isoOrEmpty = (v) => (v instanceof Date) ? v.toISOString() : (v ? String(v) : '');
    const numOr0 = (v) => {
      const n = Number(v);
      return isNaN(n) ? 0 : n;
    };
    const keepEmptyOrNumber = (v) => {
      if (v === '' || v === null || v === undefined) return '';
      const n = Number(v);
      return isNaN(n) ? '' : n;
    };

    // =====================
    // SESSIE + KAS (Kas)
    // =====================
    let session = { id: sid, startTime: '', endTime: '' };
    let cash = { startCash: '', endCash: '', diff: '', cashSales: 0, logical: '' };

    if (kasSh && kasSh.getLastRow() > 1) {
      const kas = kasSh.getRange(2,1,kasSh.getLastRow()-1,8).getValues();
      for (let i = 0; i < kas.length; i++) {
        if (String(kas[i][0]) === sid) {
          session.startTime = isoOrEmpty(kas[i][1]);
          session.endTime   = isoOrEmpty(kas[i][2]);

          cash.startCash = keepEmptyOrNumber(kas[i][3]);
          cash.endCash   = keepEmptyOrNumber(kas[i][4]);
          cash.diff      = keepEmptyOrNumber(kas[i][5]);

          // cashSales kan leeg of nummer zijn
          cash.cashSales = numOr0(kas[i][6]);
          cash.logical   = kas[i][7] ? String(kas[i][7]) : '';
          break;
        }
      }
    }

    // =====================
    // SALES (Sales)
    // =====================
    let todayTotal = 0, todayCount = 0;
    let payTotals = { pin:0, contant:0, tikkie:0, marktplaats:0, website:0, overig:0, derving:0 };
    let weekTotal = 0, monthTotal = 0;
    let monthMap = {};

    if (salesSh && salesSh.getLastRow() > 1) {
      const sales = salesSh.getRange(2,1,salesSh.getLastRow()-1,4).getValues();

      sales.forEach(r => {
        const d   = r[1]; // B datum
        const pay = String(r[2] || '').trim().toLowerCase(); // C betaalwijze
        const tot = numOr0(r[3]); // D totaal

        if (!(d instanceof Date)) return;

        const day = new Date(d);
        day.setHours(0,0,0,0);

        // vandaag
        if (day.getTime() === today.getTime()) {
          todayTotal += tot;
          todayCount++;
          if (payTotals[pay] !== undefined) payTotals[pay] += tot;
        }

        // week (laatste 7 dagen incl vandaag)
        const diffDays = Math.floor((today - day) / 86400000);
        if (diffDays >= 0 && diffDays < 7) weekTotal += tot;

        // maand (geselecteerde maand)
        if (_ymKey_(day) === selectedMonthKey) {
          monthTotal += tot;
          const key = day.toISOString().slice(0,10);
          monthMap[key] = (monthMap[key] || 0) + tot;
        }
      });
    }

    const monthChart = Object.keys(monthMap).sort().map(d => ({ date: d, total: monthMap[d] }));
    // =====================
    // DASHBOARD KPI's (Dashboard-tab)
    // =====================
    let kpis = {
      monthRevenue: 0,
      monthProfit: 0,
      quarterRevenue: 0,
      quarterRevenuePrevYear: 0,
      ytdRevenue: 0,
      ytdRevenuePrevYear: 0
    };

    let availableMonths = []; // voor dropdown

    try {
      const dashSh = ss.getSheetByName(DASHBOARD_SHEET);
      if (dashSh) {
        // Verkoop (omzet per maand)
        const verkoopRows = _readMonthValueTable_(dashSh, 'Verkoop'); // {date,value}
        // Bruto Winst (winst per maand)
        const winstRows   = _readMonthValueTable_(dashSh, 'Bruto Winst'); // {date,value}

        // Maak maps
        const omzetByMonth = {};
        verkoopRows.forEach(r => { omzetByMonth[_ymKey_(r.date)] = (omzetByMonth[_ymKey_(r.date)] || 0) + r.value; });

        const winstByMonth = {};
        winstRows.forEach(r => { winstByMonth[_ymKey_(r.date)] = (winstByMonth[_ymKey_(r.date)] || 0) + r.value; });

        // months voor dropdown: uit omzet (verkoop)
        const now = new Date();
        const currentYm = _ymKey_(now);

        // alleen maanden t/m huidige maand
        availableMonths = Object.keys(omzetByMonth)
          .filter(k => k <= currentYm)
          .sort();

        // laatste 24 tonen (maar altijd incl huidige)
        if (availableMonths.length > 24) {
          availableMonths = availableMonths.slice(-24);
        }

        // maand omzet/winst
        kpis.monthRevenue = Number(omzetByMonth[selectedMonthKey] || 0);
        kpis.monthProfit  = Number(winstByMonth[selectedMonthKey] || 0);

        // quarter omzet en vorige jaar quarter omzet
        const y = Number(selectedMonthKey.slice(0,4));
        const m = Number(selectedMonthKey.slice(5,7)) - 1;
        const selDate = new Date(y, m, 1);
        const { q } = _quarterKey_(selDate);

        const quarterMonths = [(q-1)*3, (q-1)*3 + 1, (q-1)*3 + 2].map(mm => `${y}-${String(mm+1).padStart(2,'0')}`);
        kpis.quarterRevenue = quarterMonths.reduce((s, k) => s + Number(omzetByMonth[k] || 0), 0);

        const yPrev = y - 1;
        const quarterMonthsPrev = [(q-1)*3, (q-1)*3 + 1, (q-1)*3 + 2].map(mm => `${yPrev}-${String(mm+1).padStart(2,'0')}`);
        kpis.quarterRevenuePrevYear = quarterMonthsPrev.reduce((s, k) => s + Number(omzetByMonth[k] || 0), 0);

        // YTD omzet (jan t/m selected month in selected year)
        const ytdMonths = [];
        for (let mm = 0; mm <= m; mm++){
          ytdMonths.push(`${y}-${String(mm+1).padStart(2,'0')}`);
        }
        kpis.ytdRevenue = ytdMonths.reduce((s, k) => s + Number(omzetByMonth[k] || 0), 0);

        // YTD omzet vorig jaar (jan t/m dezelfde maand vorig jaar)
        const ytdPrevMonths = [];

        for (let mm = 0; mm <= m; mm++){
          ytdPrevMonths.push(`${yPrev}-${String(mm+1).padStart(2,'0')}`);
        }

        kpis.ytdRevenuePrevYear = ytdPrevMonths.reduce(
          (s, k) => s + Number(omzetByMonth[k] || 0),
          0
        );
      }
    } catch(e) {
      // niet hard falen; frontend kan kpis leeg tonen
    }
    const paymentCounts = getPaymentCountsForMonth_(salesSh, selectedMonthKey);

    // RETURN (altijd volledig + build)
    return {
      ok: true,
      session,
      today: { total: todayTotal, count: todayCount, payments: payTotals },
      totals: { day: todayTotal, week: weekTotal, month: monthTotal },
      cash,
      monthChart,
      paymentCounts,
      selectedMonth: selectedMonthKey,
      availableMonths,
      kpis
    };

  } catch (e) {
    // Ook dit is JSON-saf
  }
}

function getPaymentCountsForMonth_(salesSh, monthKey) {
  const counts = {
    pin: 0,
    contant: 0,
    tikkie: 0,
    marktplaats: 0,
    website: 0,
    overig: 0,
    derving: 0
  };

  if (!salesSh || salesSh.getLastRow() < 2) return counts;

  // monthKey = 'YYYY-MM'
  const year  = Number(monthKey.slice(0,4));
  const month = Number(monthKey.slice(5,7)) - 1;

  const start = new Date(year, month, 1);
  const end   = new Date(year, month + 1, 1);

  const rows = salesSh
    .getRange(2, 1, salesSh.getLastRow() - 1, 4)
    .getValues();

  rows.forEach(r => {
    const date = r[1]; // kolom B
    const pay  = String(r[2] || '').trim().toLowerCase();

    if (!(date instanceof Date)) return;
    if (date < start || date >= end) return;

    if (counts.hasOwnProperty(pay)) {
      counts[pay]++;
    } else {
      counts.overig++;
    }
  });

  return counts;
}


/**
 * Activeert sessie van vandaag opnieuw
 * - Als sessie bestaat: eindtijd leegmaken
 * - Als niet bestaat: nieuwe sessie openen (zonder start cash)
 */
function apiActivateSession() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(KAS_SHEET);
  const sid = getTodaySessionId_();

  if (!sh) throw new Error('Kas-sheet niet gevonden');

  const res = getTodaySession();

  // Sessie bestaat → heropenen
  if (res.exists) {
    const row = res.row;
    sh.getRange(row, 3).setValue(''); // eindtijd leeg
    return { ok: true, reused: true };
  }

  // Geen sessie → nieuwe aanmaken zonder start cash
  openSession(null);
  return { ok: true, reused: false };
}



