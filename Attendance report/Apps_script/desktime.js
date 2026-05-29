/***** CONFIG *****/
const TZ = 'Asia/Kolkata';
const SHEET_NAME = 'Weekreport';   // your tab holding the headers

// Always choose names in this order (adjust if your sheet prefers another field)
const NAME_FIELD_PRIORITY  = ['name','full_name','fullName','user.name','Full Name'];

// Try these keys in this order to find the employee email:
const EMAIL_FIELD_PRIORITY = ['email', 'user.email', 'userEmail', 'Email'];

/***** RUN THIS *****/
// Appends last (completed) week Mon..Sun, skipping existing rows.
// Dedupe order: (1) Date+EmployeeID in-run, (2) Date+ExactName, (3) Date+LooseName
function syncDeskTimeAppendLastWeek() {
  const sh   = getSheetStrict_(SHEET_NAME);
  const cols = findColumnMap_(sh);                   // locate headers (no changes to your sheet)

  // Existing (from sheet)
  const existingExact = buildExistingKeys_(sh, cols);        // Date|exact name
  const existingLoose = buildExistingLooseKeys_(sh, cols);   // Date|loose name (sorted tokens)

  // In-run dedupe
  const seenByDateId = new Set();  // Date|employeeId
  const newLoose     = new Set();  // Date|loose name

  const {start, end} = lastWeekRange_();
  const rowsToAppend = [];

  for (const d of enumerateDays_(start, end)) {
    const ymd = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');

    const employees = retry3_(() => fetchDeskTimeEmployeesDay_(getApiKey_(), ymd)); // array

    for (const e of employees) {
      const rec = employeeToObj_(e, ymd); // builds fields + _Id (for dedupe)
      if (!rec.Name) continue;            // need a display name for the sheet

      const exactKey = dtMakeKey_(rec.Date, rec.Name);                // Date|exact name
      const looseKey = rec.Date + '|' + normalizeNameLoose_(rec.Name); // Date|loose name
      const idKey    = rec._Id ? (rec.Date + '|' + rec._Id) : null; // Date|employeeId

      // Skip if same employeeId already seen this date,
      // or if sheet already has same exact/loose name for this date,
      // or if this run already queued a row with same loose name for this date.
      if ((idKey && seenByDateId.has(idKey)) ||
          existingExact.has(exactKey) ||
          existingLoose.has(looseKey) ||
          newLoose.has(looseKey)) {
        continue;
      }

      rowsToAppend.push(rec);

      if (idKey) seenByDateId.add(idKey);
      existingExact.add(exactKey);
      existingLoose.add(looseKey);
      newLoose.add(looseKey);
    }
  }

  appendRowsToColumns_(sh, cols, rowsToAppend);
  Logger.log('Appended %d rows for %s..%s',
    rowsToAppend.length,
    Utilities.formatDate(start, TZ, 'yyyy-MM-dd'),
    Utilities.formatDate(end,   TZ, 'yyyy-MM-dd'));
}

// Clears ONLY the attendance columns (rows 2..last). Leaves other columns untouched.
function clearAttendanceColumns_(sh, cols){
  const lr = sh.getLastRow();
  if (lr < 2) return;
  const targets = [cols.date, cols.name, cols.email, cols.total, cols.arrived, cols.left];
  for (const c of targets) {
    if (c) sh.getRange(2, c, lr - 1, 1).clearContent();
  }
}

function syncDeskTimeAppendYesterday() {
  const sh   = getSheetStrict_(SHEET_NAME);
  const cols = findColumnMap_(sh);

  // Existing rows in the sheet (for dedupe)
  const existingExact = buildExistingKeys_(sh, cols);        // Date|exact name
  const existingLoose = buildExistingLooseKeys_(sh, cols);   // Date|loose name

  // In-run dedupe
  const seenByDateId = new Set();  // Date|employeeId
  const newLoose     = new Set();  // Date|loose name

  const ymd = getYesterdayYmd_();
  const rowsToAppend = [];

  const employees = retry3_(() =>
    fetchDeskTimeEmployeesDay_(getApiKey_(), ymd)
  );

  for (const e of employees) {
    const rec = employeeToObj_(e, ymd);   // has Name, Email, Date, etc.
    if (!rec.Name) continue;

    const exactKey = dtMakeKey_(rec.Date, rec.Name);                 // Date|exact name
    const looseKey = rec.Date + '|' + normalizeNameLoose_(rec.Name); // Date|loose name
    const idKey    = rec._Id ? (rec.Date + '|' + rec._Id) : null;  // Date|employeeId

    if ((idKey && seenByDateId.has(idKey)) ||
        existingExact.has(exactKey) ||
        existingLoose.has(looseKey) ||
        newLoose.has(looseKey)) {
      continue;
    }

    rowsToAppend.push(rec);

    if (idKey) seenByDateId.add(idKey);
    existingExact.add(exactKey);
    existingLoose.add(looseKey);
    newLoose.add(looseKey);
  }

  appendRowsToColumns_(sh, cols, rowsToAppend);

  Logger.log('Appended %d rows for %s', rowsToAppend.length, ymd);
}

function writeRowsToColumnsAt_(sh, cols, rows, startRow){
  if (!rows || !rows.length) return;
  const n = rows.length;
  const col = f => rows.map(r => [r[f] ?? '']);

  sh.getRange(startRow, cols.date,    n, 1).setValues(col('Date'));
  sh.getRange(startRow, cols.name,    n, 1).setValues(col('Name'));

  if (cols.email) {
    sh.getRange(startRow, cols.email, n, 1).setValues(col('Email'));
  }

  sh.getRange(startRow, cols.total,   n, 1).setValues(col('TotalDeskTime'));
  sh.getRange(startRow, cols.arrived, n, 1).setValues(col('Arrived'));
  sh.getRange(startRow, cols.left,    n, 1).setValues(col('Left'));
}

function syncDeskTimeReplaceLastWeek() {
  const sh   = getSheetStrict_(SHEET_NAME);
  const cols = findColumnMap_(sh);

  // Collect last week's rows with in-run dedupe (ID + loose-name)
  const {start, end} = lastWeekRange_();
  const seenByDateId = new Set();
  const seenLoose    = new Set();
  const rows = [];

  for (const d of enumerateDays_(start, end)) {
    const ymd = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    const employees = retry3_(() => fetchDeskTimeEmployeesDay_(getApiKey_(), ymd));
    for (const e of employees) {
      const rec = employeeToObj_(e, ymd);
      if (!rec.Name) continue;

      const idKey    = rec._Id ? (rec.Date + '|' + rec._Id) : null;
      const looseKey = rec.Date + '|' + normalizeNameLoose_(rec.Name);
      if ((idKey && seenByDateId.has(idKey)) || seenLoose.has(looseKey)) continue;

      rows.push(rec);
      if (idKey) seenByDateId.add(idKey);
      seenLoose.add(looseKey);
    }
  }

  // Clear ONLY the attendance columns under the header (keep all other columns intact)
  clearAttendanceColumns_(sh, cols);

  // Write fresh rows into those attendance columns starting at row 2
  writeRowsToColumnsAt_(sh, cols, rows, 2);

  Logger.log('Replaced with %d rows for %s..%s',
    rows.length,
    Utilities.formatDate(start, TZ, 'yyyy-MM-dd'),
    Utilities.formatDate(end,   TZ, 'yyyy-MM-dd'));
}

/***** DeskTime API (returns an ARRAY of employees or throws with a clear error) *****/
function fetchDeskTimeEmployeesDay_(apiKey, dateStr) {
  const url = 'https://desktime.com/api/v2/json/employees'
            + '?apiKey=' + encodeURIComponent(apiKey)
            + '&date='   + encodeURIComponent(dateStr)
            + '&period=day';

  const res  = UrlFetchApp.fetch(url, {method:'get', muteHttpExceptions:true});
  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code >= 400) throw new Error('DeskTime /employees HTTP ' + code + ': ' + text);

  let data;
  try { data = JSON.parse(text); }
  catch (_) { throw new Error('DeskTime response is not JSON for ' + dateStr + ': ' + text); }

  if (data && (data.error || data.status === 'error' || data.success === false)) {
    throw new Error('DeskTime API error for ' + dateStr + ': ' + (data.error || data.message || text));
  }

  // --- Normalize to an array of employee objects ---
  let arr = [];

  // Case A: already an array at top level or under common keys
  if (Array.isArray(data)) arr = data;
  else if (Array.isArray(data.employees)) arr = data.employees;
  else if (data.data && Array.isArray(data.data)) arr = data.data;
  else if (data.result && Array.isArray(data.result)) arr = data.result;
  else if (data.result_data && Array.isArray(data.result_data)) arr = data.result_data;

  // Case B: nested maps (employees -> date -> employeeId -> obj)
  if (!arr.length) {
    const takeObjectValues = (obj) =>
      obj && typeof obj === 'object' && !Array.isArray(obj) ? Object.values(obj) : [];

    if (data && data.employees && typeof data.employees === 'object') {
      if (data.employees[dateStr]) {
        arr = takeObjectValues(data.employees[dateStr]);
      } else {
        const all = [];
        for (const k in data.employees) {
          const bucket = data.employees[k];
          if (Array.isArray(bucket)) all.push(...bucket);
          else all.push(...takeObjectValues(bucket));
        }
        arr = all;
      }
    } else if (data && data.data && typeof data.data === 'object') {
      arr = takeObjectValues(data.data);
    }
  }

  if (!arr.length || typeof arr[0] !== 'object') {
    throw new Error(
      'DeskTime API: no employees array for ' + dateStr +
      '. Keys: ' + Object.keys(data || {}).join(', ').slice(0,120) +
      '. Raw: ' + text.slice(0, 300)
    );
  }
  return arr;
}

/***** Build the fields you want (plus _Id for dedupe only) *****/
function employeeToObj_(e, ymd) {
  const id    = getEmployeeId_(e);
  const name  = pickName_(e);
  const email = pickEmail_(e);     // <— EMAIL picked here

  // total desk time (seconds)
  const deskSecs = pickNumber_(e, [
    'desktime_time','desktimeTime','desktime','desk_time','time_at_desk',
    'total_desktime','totalDeskTime'
  ]);

  // arrival / left (handle multiple field variants)
  const arrivalRaw  = pickTime_(e, [
    'arrival','arrival_time','arrived',
    'first_activity','firstActivity','first_active','start_time','startedAt'
  ]);
  const leftRaw     = pickTime_(e, [
    'left','left_time','departed','departure',
    'last_activity','lastActivity','last_active','end_time','endedAt'
  ]);

  return {
    Date: ymd,
    Name:   String(name  || '').trim(),
    Email:  String(email || '').trim(),   // <— used in sheet
    // Prefer H:MM like timesheets. If you want H:MM:SS, swap to secsToHHMMSS_(deskSecs).
    TotalDeskTime: secsToHhmm_(deskSecs),
    Arrived: fmtTimeOfDay_(arrivalRaw, TZ, true, ymd), // "h:mm AM/PM"
    Left:    fmtTimeOfDay_(leftRaw,    TZ, true, ymd),

    // used only for in-run dedupe, not written to the sheet
    _Id: id
  };
}

/***** Append ONLY the attendance columns (no header changes, no extra columns) *****/
function getLastAttendanceRow_(sh, cols){
  const lr = sh.getLastRow();
  if (lr < 2) return 1; // header only
  const targets = [cols.date, cols.name, cols.email, cols.total, cols.arrived, cols.left].filter(Boolean);
  let last = 1;
  for (const c of targets) {
    const vals = sh.getRange(2, c, lr - 1, 1).getDisplayValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0] || '').trim() !== '') {
        last = Math.max(last, i + 2); // convert to absolute row index
        break;
      }
    }
  }
  return last; // 1 means "only header used"
}

function appendRowsToColumns_(sh, cols, rows) {
  if (!rows || !rows.length) return;

  // Start just after the last-used row within the attendance columns
  const startRow = getLastAttendanceRow_(sh, cols) + 1;

  const n = rows.length;
  const col = f => rows.map(r => [r[f] ?? '']);

  sh.getRange(startRow, cols.date,    n, 1).setValues(col('Date'));
  sh.getRange(startRow, cols.name,    n, 1).setValues(col('Name'));

  if (cols.email) {
    sh.getRange(startRow, cols.email, n, 1).setValues(col('Email'));
  }

  sh.getRange(startRow, cols.total,   n, 1).setValues(col('TotalDeskTime'));
  sh.getRange(startRow, cols.arrived, n, 1).setValues(col('Arrived'));
  sh.getRange(startRow, cols.left,    n, 1).setValues(col('Left'));
}

/***** Locate your existing headers (row 1) – no reordering/renaming *****/
function getSheetStrict_(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found.');
  return sh;
}

function findColumnMap_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error('Header row missing.');
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim());

  const find = (patterns) => {
    for (let c = 0; c < headers.length; c++) {
      const h  = headers[c];
      const hn = h.toLowerCase().replace(/\s+/g,' ').trim();
      for (const re of patterns) if (re.test(h) || re.test(hn)) return c + 1; // 1-based
    }
    return null;
  };

  // Accept slight variations like "Total ,DeskTime"
  const colDate    = find([/^date$/i]);
  const colName    = find([/^name$/i, /^full\s*name$/i]);
  const colEmail   = find([/^email$/i]); // optional email column
  const colTotal   = find([/^total\s*,?\s*desktime$/i, /^total\s*desk\s*time$/i, /^totaldesktime$/i, /^total$/i]);
  const colArrived = find([/^arrived$/i, /^arrival(\s*time)?$/i]);
  const colLeft    = find([/^left$/i, /^departure$/i, /^left(\s*time)?$/i]);

  const miss = [];
  if (!colDate)    miss.push('Date');
  if (!colName)    miss.push('Name');
  if (!colTotal)   miss.push('Total DeskTime');
  if (!colArrived) miss.push('Arrived');
  if (!colLeft)    miss.push('Left');
  if (miss.length) throw new Error('Could not find header(s): ' + miss.join(', ') + ' in row 1.');

  return {
    date:    colDate,
    name:    colName,
    email:   colEmail,   // may be null if you remove the column
    total:   colTotal,
    arrived: colArrived,
    left:    colLeft
  };
}

/***** Dedupe (existing exact + existing loose) *****/
function buildExistingKeys_(sh, cols) {
  const set = new Set();
  const lr = sh.getLastRow();
  if (lr < 2) return set;
  const dates = sh.getRange(2, cols.date, lr - 1, 1).getValues();
  const names = sh.getRange(2, cols.name, lr - 1, 1).getValues();
  for (let i = 0; i < dates.length; i++) {
    const ymd = toYmdGuess_(dates[i][0]);
    const nm  = String(names[i][0] || '').trim();
    if (ymd && nm) set.add(dtMakeKey_(ymd, nm)); // exact key
  }
  return set;
}

function buildExistingLooseKeys_(sh, cols) {
  const set = new Set();
  const lr = sh.getLastRow();
  if (lr < 2) return set;
  const dates = sh.getRange(2, cols.date, lr - 1, 1).getValues();
  const names = sh.getRange(2, cols.name, lr - 1, 1).getValues();
  for (let i = 0; i < dates.length; i++) {
    const ymd = toYmdGuess_(dates[i][0]);
    const nm  = String(names[i][0] || '').trim();
    if (ymd && nm) set.add(ymd + '|' + normalizeNameLoose_(nm));
  }
  return set;
}

// *** NEW NAME to avoid collision with other files ***
function dtMakeKey_(ymd, name) {
  return ymd + '|' + String(name || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/***** Name, Email & ID helpers for canonical display + dedupe *****/
function pickName_(e){
  for (const p of NAME_FIELD_PRIORITY) {
    let v;
    if (p.indexOf('.') > -1) {
      const [a,b] = p.split('.');
      v = e && e[a] && e[a][b];
    } else {
      v = e && e[p];
    }
    if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function pickEmail_(e){
  for (const p of EMAIL_FIELD_PRIORITY) {
    let v;
    if (p.indexOf('.') > -1) {
      const [a,b] = p.split('.');
      v = e && e[a] && e[a][b];
    } else {
      v = e && e[p];
    }
    if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

function getEmployeeId_(e){
  const cands = ['id','employee_id','employeeId','user_id','userId'];
  for (const k of cands) {
    const v = e && e[k];
    if (typeof v === 'number' && v > 0) return v;
    if (typeof v === 'string' && /^\d+$/.test(v)) return Number(v);
  }
  if (e && e.user) {
    const v = e.user.id;
    if (typeof v === 'number' && v > 0) return v;
    if (typeof v === 'string' && /^\d+$/.test(v)) return Number(v);
  }
  return null;
}

function normalizeNameLoose_(name){
  return String(name || '')
    .toLowerCase()
    .replace(/\s+/g,' ')
    .trim()
    .split(' ')
    .sort()
    .join(' ');
}

/***** Helpers: time, parsing, retries, keys, API key *****/

function getYesterdayYmd_() {
  const now   = new Date();
  const local = new Date(Utilities.formatDate(now, TZ, "yyyy-MM-dd'T'HH:mm:ss"));
  local.setDate(local.getDate() - 1); // move back one day in Asia/Kolkata time
  return Utilities.formatDate(local, TZ, 'yyyy-MM-dd');
}

function secsToHhmm_(s) {
  s = Math.max(0, Number(s || 0));
  const h = Math.floor(s/3600);
  const m = Math.floor((s%3600)/60);
  return Utilities.formatString('%d:%02d', h, m);
}

function secsToHHMMSS_(s) { // keep available if you prefer H:MM:SS
  s = Math.max(0, Number(s || 0));
  const h   = Math.floor(s/3600);
  const m   = Math.floor((s%3600)/60);
  const sec = Math.floor((s%60));
  return Utilities.formatString('%d:%02d:%02d', h, m, sec);
}

// NOTE: uses the specific day's midnight (ymd) for "seconds from midnight"
function fmtTimeOfDay_(val, tz, twelveHour, ymdOpt) {
  if (val === undefined || val === null || val === '') return '';

  if (typeof val === 'string') {
    const s = val.trim();
    if (/^\d{1,2}:\d{2}$/.test(s)) return twelveHour ? HHmmTo12_(s) : normalizeHHmm_(s);
    const n = Number(s);
    if (!isNaN(n)) return fmtTimeOfDay_(n, tz, twelveHour, ymdOpt);
    const d = new Date(s);
    if (d instanceof Date && !isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz, twelveHour ? 'h:mm a' : 'HH:mm');
    }
    return ''; // unknown string → blank
  }

  if (typeof val === 'number') {
    if (val <= 0) return '';                        // treat 0/negatives as "no time"
    let ms;
    if (val > 1e12) ms = val;                       // millis
    else if (val > 1e6) ms = val * 1000;            // epoch seconds
    else ms = secondsFromMidnightToMillis_(val, tz, ymdOpt); // seconds since midnight
    const d = new Date(ms);
    return Utilities.formatDate(d, tz, twelveHour ? 'h:mm a' : 'HH:mm');
  }

  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, tz, twelveHour ? 'h:mm a' : 'HH:mm');
  }
  return '';
}

function HHmmTo12_(s){
  const parts = (s+'').split(':');
  const H     = Number(parts[0] || 0);
  const M     = Number(parts[1] || 0);
  const am    = H < 12;
  const h12   = ((H + 11) % 12) + 1;
  return Utilities.formatString('%d:%02d %s', h12, M, am ? 'AM' : 'PM');
}

function normalizeHHmm_(x){
  if (x instanceof Date && !isNaN(x.getTime())) return Utilities.formatDate(x, TZ, 'HH:mm');
  const s = String(x || '').trim();
  if (/^\d{1,2}:\d{2}$/.test(s)) {
    const parts = s.split(':');
    const H     = Number(parts[0] || 0);
    const M     = Number(parts[1] || 0);
    return Utilities.formatString('%02d:%02d', H, M);
  }
  const d = new Date(s);
  if (d instanceof Date && !isNaN(d.getTime())) return Utilities.formatDate(d, TZ, 'HH:mm');
  return '';
}

// Use the given ymd's midnight in tz (better than "today")
function secondsFromMidnightToMillis_(sec, tz, ymdOpt) {
  const ymd = (ymdOpt && /^\d{4}-\d{2}-\d{2}$/.test(ymdOpt)) ? ymdOpt
            : Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const base = new Date(Utilities.formatDate(new Date(ymd + 'T00:00:00'), tz, "yyyy-MM-dd'T'HH:mm:ss"));
  return base.getTime() + (Number(sec)||0) * 1000;
}

function toYmdGuess_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  return (d instanceof Date && !isNaN(d.getTime())) ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : '';
}

function lastWeekRange_(){
  const now   = new Date();
  const local = new Date(Utilities.formatDate(now, TZ, "yyyy-MM-dd'T'HH:mm:ss"));
  const day   = local.getDay();                 // 0 Sun..6 Sat
  const diffToMonday = (day + 6) % 7;
  const thisMon = new Date(local); thisMon.setDate(local.getDate() - diffToMonday);
  const lastMon = new Date(thisMon); lastMon.setDate(thisMon.getDate() - 7);
  const lastSat = new Date(lastMon); lastSat.setDate(lastMon.getDate() + 5);
  setToStartOfDay_(lastMon); setToEndOfDay_(lastSat);
  return {start:lastMon, end:lastSat};
}

function enumerateDays_(start, end){
  const out=[]; const d=new Date(start);
  while (d <= end) { out.push(new Date(d)); d.setDate(d.getDate()+1); }
  return out;
}

function setToStartOfDay_(d){ d.setHours(0,0,0,0); }
function setToEndOfDay_(d){ d.setHours(23,59,59,999); }

function coalesce_(...xs){
  for (const x of xs) if (x !== undefined && x !== null && x !== '') return x;
  return undefined;
}

function pickNumber_(obj, keys){
  for (const k of keys){
    if (obj && typeof obj[k] === 'number') return obj[k];
    if (obj && typeof obj[k] === 'string' && obj[k].trim() !== '' && !isNaN(Number(obj[k])))
      return Number(obj[k]);
    if (obj && obj.stats && typeof obj.stats[k] !== 'undefined') {
      const v = obj.stats[k];
      if (typeof v === 'number') return v;
      if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) return Number(v);
    }
  }
  return 0;
}

function pickTime_(obj, keys){
  for (const k of keys){
    if (obj && typeof obj[k] !== 'undefined' && obj[k] !== null && obj[k] !== '') return obj[k];
    if (obj && obj.stats && typeof obj.stats[k] !== 'undefined' && obj.stats[k] !== null && obj.stats[k] !== '')
      return obj.stats[k];
  }
  return '';
}

function retry3_(fn){
  const waits = [0, 400, 1200];
  let err;
  for (const w of waits){ try { if (w) Utilities.sleep(w); return fn(); } catch (e){ err = e; } }
  throw err;
}

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('DESKTIME_API_KEY');
  if (!key) throw new Error('Missing DESKTIME_API_KEY in Script properties.');
  return key;
}



