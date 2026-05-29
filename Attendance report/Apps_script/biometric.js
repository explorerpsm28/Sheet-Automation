function syncAttendanceBySheetDates() {
  const AK  = 'LlQ4aRikl7b6YSzFFamluZ5VDxqUEkg7';
  const SK  = '6YSgpCPZKiq2hTBAZo7NGq6MvEkFMurm';
  const FALLBACK_BASE = 'https://isgp.hikcentralconnect.com';
  const TZ   = 'Asia/Kolkata';
  const SHEET_NAME = 'Weekreport';
  const PAGE_SIZE = 200;

  const fmtDate = d => Utilities.formatDate(d, TZ, 'yyyy/MM/dd');
  const fmtISO  = d => Utilities.formatDate(d, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
  const sod = d => (d = new Date(d), d.setHours(0,0,0,0), d);
  const eod = d => (d = new Date(d), d.setHours(23,59,59,999), d);

  // find min & max valid dates in the Date column
  const sheetDateRangeIST = (rows, dateColIndex) => {
    let min = null, max = null;
    for (let i = 0; i < rows.length; i++) {
      let d = rows[i][dateColIndex];
      if (!d) continue;
      if (!(d instanceof Date)) d = new Date(d);
      if (isNaN(+d)) continue;
      d = sod(d);
      if (!min || d < min) min = d;
      if (!max || d > max) max = d;
    }
    if (!min || !max) throw new Error('No valid dates found in "Date" column.');
    return { start: min, end: max };
  };

  const httpPost = (url, body, headers) => {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body || {}),
      headers: headers || {},
      muteHttpExceptions: true
    });
    const text = res.getContentText();
    let json;
    try { json = JSON.parse(text); }
    catch (e) { throw new Error(`HTTP ${res.getResponseCode()} at ${url}: ${text}`); }
    if (json.errorCode !== '0') throw new Error(`API error at ${url}: ${text}`);
    return json.data || {};
  };

  // open sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows found.');
    return;
  }

  const header = sheet.getRange(1,1,1,lastCol).getValues()[0].map(x => String(x).trim());
  const idx = t => {
    const i = header.indexOf(t);
    if (i < 0) throw new Error(`Missing header: ${t}`);
    return i;
  };

  const C_FULL = idx('Full Name');
  const C_DATE = idx('Date');
  const C_OS   = idx('Office Start Time');
  const C_CIN  = idx('Clock-In Time');
  const C_OE   = idx('Office End Time');
  const C_COUT = idx('Clock-Out Time');
  const C_H    = idx('Office Hours');

  const data = sheet.getRange(2,1,lastRow-1,lastCol).getValues();

  // build date window from sheet (NOT last week)
  const { start, end } = sheetDateRangeIST(data, C_DATE);
  const beginISO = fmtISO(sod(start));
  const endISO   = fmtISO(eod(end));

  // 1) token
  const tokenResp = httpPost(
    `${FALLBACK_BASE}/api/hccgw/platform/v1/token/get`,
    { appKey: AK, secretKey: SK }
  );
  const token = tokenResp.accessToken;
  if (!token) throw new Error('Token missing in response from token/get.');
  const BASE = (tokenResp.areaDomain || FALLBACK_BASE).replace(/\/+$/,'');

  // 2) get attendance for that whole range
  const headers = { 'Token': String(token) };
  let pageIndex = 1;
  const all = [];

  while (true) {
    const body = {
      pageIndex,
      pageSize: PAGE_SIZE,
      beginTime: beginISO,
      endTime: endISO,
      dateFormat: 'yyyy/MM/dd',
      timeFormat: 'HH:mm',
      durationFormat: 'HH:MM'
    };
    const d = httpPost(`${BASE}/api/hccgw/attendance/v1/report/totaltimecard/list`, body, headers);
    const list = d.reportDataList || [];
    all.push.apply(all, list);
    if (d.moreData === 1) pageIndex++; else break;
  }

  // 3) index by (fullNameLower|date)
  const map = new Map();
  all.forEach(it => {
    const key = `${String(it.fullName || '').trim().toLowerCase()}|${String(it.date || '').trim()}`;
    map.set(key, {
      os:   it.checkInTime   || '',
      cin:  it.clockInTime   || '',
      oe:   it.checkOutTime  || '',
      cout: it.clockOutTime  || '',
      h:    it.workDuration  || ''
    });
  });

  // 4) for EACH ROW: use that row's Full Name + Date
  const outOS   = [];
  const outCIN  = [];
  const outOE   = [];
  const outCOUT = [];
  const outH    = [];

  for (let r = 0; r < data.length; r++) {
    const row  = data[r];
    const name = String(row[C_FULL] || '').trim();
    let d      = row[C_DATE];
    if (!(d instanceof Date) && d) d = new Date(d);

    let os = '', cin = '', oe = '', cout = '', h = '';

    if (name && d instanceof Date && !isNaN(+d)) {
      const keyDate = fmtDate(d); // yyyy/MM/dd
      const hit = map.get(`${name.toLowerCase()}|${keyDate}`);
      if (hit) {
        os   = hit.os;
        cin  = hit.cin;
        oe   = hit.oe;
        cout = hit.cout;
        h    = hit.h;
      }
    }

    outOS.push([os]);
    outCIN.push([cin]);
    outOE.push([oe]);
    outCOUT.push([cout]);
    outH.push([h]);
  }

  sheet.getRange(2, C_OS+1,   outOS.length,   1).setValues(outOS);
  sheet.getRange(2, C_CIN+1,  outCIN.length,  1).setValues(outCIN);
  sheet.getRange(2, C_OE+1,   outOE.length,   1).setValues(outOE);
  sheet.getRange(2, C_COUT+1, outCOUT.length, 1).setValues(outCOUT);
  sheet.getRange(2, C_H+1,    outH.length,    1).setValues(outH);

  Logger.log(`Synced ${fmtDate(start)} → ${fmtDate(end)}, BASE=${BASE}`);
}
