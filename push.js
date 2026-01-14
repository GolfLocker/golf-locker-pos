const PUSHOVER_USERS = [
  'uitgihsqoui9hjppdm54491bu1howh',   ////Danilo
  'u5eappra662knm6j9ngdvbv3a9e5ru',   ////Menno
  'uyanzyyvcd3t7rjprg9n5rgmg1i1n5'   ////Werk telefoon
];

const PUSHOVER_APP_TOKEN = 'a9ojzf33y6msiuf491xfz75ma268b6';

function sendPush(title, message) {
  PUSHOVER_USERS.forEach(userKey => {
    UrlFetchApp.fetch('https://api.pushover.net/1/messages.json', {
      method: 'post',
      payload: {
        token: PUSHOVER_APP_TOKEN,
        user: userKey,
        title,
        message,
        sound: 'cashregister'
      },
      muteHttpExceptions: true
    });
  });
}


function sendDailySummaryPush() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Sales');
  if (!sh) return;

  const tz = Session.getScriptTimeZone() || 'Europe/Amsterdam';
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  let total = 0;
  let count = 0;
  const pay = {};

  rows.forEach(r => {
    const date = r[1];      // datum
    const method = r[2];    // betaalwijze
    const amount = Number(r[3]) || 0;

    if (!(date instanceof Date)) return;

    const rowDate = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
    if (rowDate !== todayStr) return;

    total += amount;
    count++;

    pay[method] = (pay[method] || 0) + amount;
  });

  if (count === 0) return; // niks verkocht → geen push

  let lines = [
    `📊 Dagoverzicht (${Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy')})`,
    ``,
    `Omzet: € ${total.toFixed(2)}`,
    `Transacties: ${count}`,
    ``
  ];

  Object.keys(pay).forEach(k => {
    lines.push(`${k}: € ${pay[k].toFixed(2)}`);
  });

  sendPush(
    'Golf Locker POS',
    lines.join('\n')
  );
}

function sendWeeklySummaryPush() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Sales');
  if (!sh) return;

  const tz = Session.getScriptTimeZone() || 'Europe/Amsterdam';
  const now = new Date();

  // Maandag van deze week
  const day = now.getDay(); // 0 = zondag
  const diffToMonday = (day === 0 ? -6 : 1) - day;
  const monday = new Date(now);
  monday.setDate(now.getDate() + diffToMonday);
  monday.setHours(0,0,0,0);

  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  sunday.setHours(23,59,59,999);

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  let total = 0;
  let count = 0;
  const pay = {};

  rows.forEach(r => {
    const date = r[1];
    const method = r[2];
    const amount = Number(r[3]) || 0;

    if (!(date instanceof Date)) return;
    if (date < monday || date > sunday) return;

    total += amount;
    count++;
    pay[method] = (pay[method] || 0) + amount;
  });

  if (count === 0) return;

  const fmt = d => Utilities.formatDate(d, tz, 'dd-MM-yyyy');

  let lines = [
    `📅 Weekoverzicht`,
    `${fmt(monday)} t/m ${fmt(sunday)}`,
    ``,
    `Omzet: € ${total.toFixed(2)}`,
    `Transacties: ${count}`,
    ``
  ];

  Object.keys(pay).forEach(k => {
    lines.push(`${k}: € ${pay[k].toFixed(2)}`);
  });

  sendPush('Golf Locker POS', lines.join('\n'));
}

function sendMonthlySummaryPush() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Sales');
  if (!sh) return;

  const tz = Session.getScriptTimeZone() || 'Europe/Amsterdam';
  const now = new Date();

  const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastDay  = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  lastDay.setHours(23,59,59,999);

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  let total = 0;
  let count = 0;
  const pay = {};

  rows.forEach(r => {
    const date = r[1];
    const method = r[2];
    const amount = Number(r[3]) || 0;

    if (!(date instanceof Date)) return;
    if (date < firstDay || date > lastDay) return;

    total += amount;
    count++;
    pay[method] = (pay[method] || 0) + amount;
  });

  if (count === 0) return;

  const monthName = Utilities.formatDate(now, tz, 'MMMM yyyy');

  let lines = [
    `📊 Maandoverzicht`,
    `${monthName}`,
    ``,
    `Omzet: € ${total.toFixed(2)}`,
    `Transacties: ${count}`,
    ``
  ];

  Object.keys(pay).forEach(k => {
    lines.push(`${k}: € ${pay[k].toFixed(2)}`);
  });

  sendPush('Golf Locker POS', lines.join('\n'));
}
