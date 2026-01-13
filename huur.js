const VERHUUR_SHEET = 'Verhuur';

function _nextHuurNr_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(VERHUUR_SHEET);
  if (!sh) throw new Error('Tab Verhuur ontbreekt');

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd'
  );

  let count = 0;
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    vals.forEach(v => {
      if (String(v[0]).includes(today)) count++;
    });
  }

  const seq = String(count + 1).padStart(3, '0');
  return `HU-${today}-${seq}`;
}

function apiStartHuur(items, startdatum, einddatum, klant, betaalwijze) {
  if (!Array.isArray(items) || !items.length) {
    throw new Error('Geen huurartikelen ontvangen');
  }

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(VERHUUR_SHEET);
  if (!sh) throw new Error('Tab Verhuur ontbreekt');

  const huurNr = _nextHuurNr_();
  const now = new Date();

  items.forEach(it => {
    if (!it.sku || !it.desc || !it.verkoopprijs) {
      throw new Error('Ongeldig huurartikel');
    }

    const calc = _calcHuurPrijs_(
      it.verkoopprijs,
      startdatum,
      einddatum
    );

    sh.appendRow([
      huurNr,                    // HuurNr
      new Date(startdatum),      // Datum_start
      new Date(einddatum),       // Datum_eind
      '',                        // Datum_ingeleverd
      it.sku,                    // SKU
      it.desc,                   // Omschrijving
      it.verkoopprijs,           // Verkoopprijs
      calc.initiele,             // Initiele_huur
      calc.dagprijs,             // Dagprijs
      calc.dagen,                // Aantal_dagen
      calc.extraDagen,           // Extra_dagen
      calc.weekKorting,          // Week_korting
      calc.huurTotaal,           // Huur_totaal
      calc.borg,                 // Borg
      calc.teBetalen,            // Te_betalen
      'ACTIEF',                  // Status
      klant || '',               // Klant
      betaalwijze || '',         // Betaalwijze
      ''                          // Opmerking
    ]);
  });

  let huurTotaal = 0;
  let borgTotaal = 0;

  items.forEach(it => {
    const calc = _calcHuurPrijs_(it.verkoopprijs, startdatum, einddatum);
    huurTotaal += calc.huurTotaal;
    borgTotaal += calc.borg;
  });

  return {
    ok: true,
    huurNr,
    huurTotaal: Math.round(huurTotaal * 100) / 100,
    borgTotaal: Math.round(borgTotaal * 100) / 100,
    eindTotaal: Math.round((huurTotaal + borgTotaal) * 100) / 100,
    startdatum,
    einddatum,
    betaalwijze
  };

}

function apiEindigHuur(huurNr) {
  if (!huurNr) throw new Error('Geen huurnummer');

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(VERHUUR_SHEET);
  if (!sh) throw new Error('Tab Verhuur ontbreekt');

  const data = sh.getDataRange().getValues();
  const now = new Date();
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(huurNr) && data[i][15] === 'ACTIEF') {
      sh.getRange(i + 1, 4).setValue(now);      // Datum_ingeleverd
      sh.getRange(i + 1, 16).setValue('INGELEVERD'); // Status
      found = true;
    }
  }

  if (!found) {
    throw new Error('Geen actieve huur gevonden voor dit huurnummer');
  }

  return { ok:true };
}

function _calcHuurPrijs_(verkoopprijs, start, eind) {
  verkoopprijs = Number(verkoopprijs) || 0;
  if (verkoopprijs <= 0) throw new Error('Ongeldige verkoopprijs');

  const s = new Date(start);
  const e = new Date(eind);
  s.setHours(0,0,0,0);
  e.setHours(0,0,0,0);

  const dagen = Math.round((e - s) / 86400000) + 1;
  if (dagen <= 0) throw new Error('Ongeldige huurperiode');

  // 1️⃣ Initiële huur
  const initiele = Math.min(verkoopprijs * 0.15, 50);

  // 2️⃣ Dagprijs
  const dagprijs = verkoopprijs > 400 ? 10 : 5;

  // 3️⃣ Extra dagen (vanaf dag 2)
  const extraDagen = Math.max(0, dagen - 1);

  // 4️⃣ Weekkorting: dag 6 & 7 gratis per week
  const volleWeken = Math.floor(extraDagen / 7);
  const gratisDagen = volleWeken * 2;
  const betaaldeDagen = Math.max(0, extraDagen - gratisDagen);

  const weekKorting = gratisDagen * dagprijs;

  const huurTotaal =
    Math.round((initiele + (betaaldeDagen * dagprijs)) * 100) / 100;

  const borg = Math.round((verkoopprijs / 2) * 100) / 100;

  return {
    dagen,
    extraDagen,
    dagprijs,
    initiele,
    weekKorting,
    huurTotaal,
    borg,
    teBetalen: huurTotaal
  };
}

function apiGetActieveHuren() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(VERHUUR_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];

  const rows = sh.getRange(2,1,sh.getLastRow()-1,19).getValues();

  const map = {};
  rows.forEach(r => {
    if (r[15] !== 'ACTIEF') return;
    if (!map[r[0]]) map[r[0]] = [];
    map[r[0]].push({
      sku: r[4],
      omschrijving: r[5],
      huurTotaal: r[12],
      borg: r[13],
      klant: r[16]
    });
  });

  return Object.keys(map).map(k => ({
    huurNr: k,
    items: map[k]
  }));
}

function apiLookupHuurSku(sku) {
  if (!sku) throw new Error('Geen SKU');

  const ss = SpreadsheetApp.getActive();
  const sheets = ['Clubs','Sets','Tassen',"Trolley's",'Overig','Diensten'];

  for (const name of sheets) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(sku).trim()) {
        return {
          sku: data[i][0],
          desc: data[i][1] || '',
          verkoopprijs: Number(data[i][5]) || 0
        };
      }
    }
  }

  throw new Error('SKU niet gevonden in voorraad');
}

function apiCalcHuurPreview(verkoopprijs, startdatum, einddatum) {
  if (!startdatum || !einddatum) {
    return { ok:false, error:'Geen datums' };
  }

  const calc = _calcHuurPrijs_(verkoopprijs, startdatum, einddatum);
  return {
    ok: true,
    huurTotaal: calc.huurTotaal,
    borg: calc.borg,
    teBetalen: calc.teBetalen,
    dagen: calc.dagen
  };
}

function buildHuurBonHtml_(data) {
  const {
    huurNr,
    startdatum,
    einddatum,
    items,
    huurTotaal,
    borgTotaal,
    eindTotaal,
    betaalwijze
  } = data;

  const rows = items.map(it => `
    <tr>
      <td colspan="2">${it.desc}</td>
    </tr>
    <tr>
      <td class="small">${it.sku}</td>
      <td class="right">€ ${it.huurTotaal.toFixed(2)}</td>
    </tr>
  `).join('');

  return `
  <html>
    <head>
      <style>
        @page {
          size: 80mm auto;
          margin: 4mm;
        }

        body {
          font-family: monospace;
          font-size: 12px;
          width: 72mm;
          margin: 0;
          padding: 0;
        }

        h2 {
          text-align: center;
          margin: 0 0 6px 0;
          font-size: 14px;
        }

        .center { text-align: center; }
        .right { text-align: right; }

        table {
          width: 100%;
          border-collapse: collapse;
        }

        td {
          padding: 2px 0;
          vertical-align: top;
        }

        .line {
          border-top: 1px dashed #000;
          margin: 6px 0;
        }

        .total {
          font-weight: bold;
          font-size: 14px;
        }

        .small {
          font-size: 11px;
        }
      </style>
    </head>

    <body>
      <h2>GOLF LOCKER</h2>
      <div class="center small">Huurbon</div>

      <div class="line"></div>

      <div class="small">
        HuurNr: ${huurNr}<br>
        Start: ${startdatum}<br>
        Eind: ${einddatum}<br>
        Betaalwijze: ${betaalwijze || '-'}
      </div>

      <div class="line"></div>

      <table>
        ${rows}
      </table>

      <div class="line"></div>

      <table>
        <tr>
          <td>Huur</td>
          <td class="right">€ ${huurTotaal.toFixed(2)}</td>
        </tr>
        <tr>
          <td>Borg</td>
          <td class="right">€ ${borgTotaal.toFixed(2)}</td>
        </tr>
        <tr class="total">
          <td>Totaal</td>
          <td class="right">€ ${eindTotaal.toFixed(2)}</td>
        </tr>
      </table>

      <div class="line"></div>

      <div class="center small">
        Dank voor het huren bij Golf Locker
      </div>

    </body>

  </html>
  `;
}

function apiPrintHuurBon(payload) {
  if (!payload || !payload.items || !payload.huurNr) {
    throw new Error('Ongeldige huurbon data');
  }

  const html = _build80mmHuurTicketHtml_(payload);
  return { ok:true, html };
}

function apiOmzettenNaarVerkoop(huurNr) {
  if (!huurNr) throw new Error('Geen huurnummer');

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(VERHUUR_SHEET);
  if (!sh) throw new Error('Tab Verhuur ontbreekt');

  const data = sh.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === huurNr && data[i][15] === 'ACTIEF') {
      sh.getRange(i + 1, 16).setValue('OMGEZET_NAAR_VERKOOP');
      found = true;
    }
  }

  if (!found) throw new Error('Geen actieve huur gevonden');

  return { ok:true };
}

//huurbon//

function _build80mmHuurTicketHtml_(opts) {

  const {
    huurNr,
    betaalwijze,
    startdatum,
    einddatum,
    huurTotaal,
    borgTotaal,
    eindTotaal,
    items
  } = opts;

  const enc = encodeURIComponent;
  const fmt = n => Utilities.formatString(
    "€ %s",
    Number(n || 0).toFixed(2).replace('.', ',')
  );

  const qrUrl = `https://quickchart.io/qr?text=${enc(BRAND.webshopUrl || '')}&size=180&margin=1&format=png`;
  const code128Url = `https://bwipjs-api.metafloor.com/?bcid=code128&text=${enc(huurNr || '')}&scale=3&height=12&includetext&textxalign=center`;

  // let op: </script> escapen
  const autoPrintScript = `
    <script>
      window.onload = function() {
        try { window.print(); } catch(e){}
      };
    <\\/script>
  `;

  const rowsHtml = (items || []).map(it => {
    const skuShort = String(it.sku || '').slice(0, 6).padEnd(6, ' ');
    const descSafe = String(it.desc || '').replace(/</g, '&lt;');
    const price    = Number(it.huurTotaal || 0);

    return `
      <tr>
        <td class="sku mono">${skuShort}</td>
        <td class="desc">${descSafe}</td>
        <td class="price">${fmt(price)}</td>
      </tr>
    `;
  }).join('');

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Huurbon ${huurNr || ''}</title>
  <style>
    @page { size: 80mm auto; margin: 4mm 4mm; }
    * { box-sizing:border-box; }
    body {
      font:11px/1.25 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
      color:#000;
      margin:0;
    }
    .ticket { width:72mm; max-width:72mm; margin:0 auto; }
    .center { text-align:center; }
    .right { text-align:right; }
    .muted { color:#555; }
    .logo { display:block; width:54mm; margin:0 auto 8px; filter:grayscale(100%); }
    h1 { font-size:14px; margin:0 0 6px; text-align:center; }
    .small { font-size:10px; }
    .mono { font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace; }
    .row { display:flex; justify-content:space-between; }
    hr { border:0; border-top:1px dashed #000; margin:6px 0; }
    table { width:100%; border-collapse:collapse; }
    th,td { padding:2px 0; vertical-align:top; }
    th { text-align:left; font-weight:700; }
    .sku   { width:15mm; font-size:9px; }
    .desc  { width:40mm; padding-right:1.5mm; }
    .price { width:18mm; text-align:right; }
    .total { font-size:13px; font-weight:700; }
    .qr { width:36mm; margin:6px auto 0; }
    .barcode { width:60mm; margin:6px auto 0; display:block; }
    @media print { .noprint { display:none!important; } }
  </style>
</head>
<body>
  <div class="ticket">
    ${BRAND.logoUrl ? `<img class="logo" src="${BRAND.logoUrl}" alt="logo">` : ''}
    <h1>${BRAND.name || ''}</h1>
    <div class="center small muted">HUURBON</div>
    <div class="center small muted">${BRAND.line1 || ''} • ${BRAND.line2 || ''}</div>
    <div class="center small muted">${BRAND.phone || ''} • ${BRAND.email || ''}</div>
    <div class="center small muted">BTW: ${BRAND.vat || '-'} • KvK: ${BRAND.kvk || '-'}</div>

    <hr>

    <div class="row small">
      <div>Betaalwijze: ${betaalwijze || '-'}</div>
      <div class="right">${startdatum || ''}</div>
    </div>
    <div class="small muted">Periode: ${startdatum} t/m ${einddatum}</div>
    <div class="small muted">HuurNr: ${huurNr}</div>

    <hr>

    <table>
      <thead>
        <tr>
          <th class="sku">SKU</th>
          <th class="desc">Artikel</th>
          <th class="price">Huur</th>
        </tr>
      </thead>
      <tbody>
        ${rowsHtml}
      </tbody>
    </table>

    <hr>

    <table>
      <tr>
        <td class="right" colspan="5">Huur: ${fmt(huurTotaal)}</td>
      </tr>
      <tr>
        <td class="right" colspan="5">Borg: ${fmt(borgTotaal)}</td>
      </tr>
      <tr>
        <td class="right total" colspan="5">Totaal: ${fmt(eindTotaal)}</td>
      </tr>
    </table>

    ${BRAND.extra ? `<div class="right small muted" style="margin-top:3px">${BRAND.extra}</div>` : ''}

    <div class="center">
      <img class="qr" src="${qrUrl}" alt="qr">
    </div>
    <div class="center small muted" style="margin-top:4px">
      ${BRAND.webshopUrl || ''}
    </div>

    <img class="barcode" src="${code128Url}" alt="Huurbon ${huurNr || ''}">

    <div class="center small muted" style="margin-top:6px">
      Bedankt voor het huren bij Golf Locker!
    </div>

    <div class="noprint center" style="margin-top:8px">
      <button onclick="window.print()">Print</button>
    </div>
  </div>

  ${autoPrintScript}
</body>
</html>`;
}
