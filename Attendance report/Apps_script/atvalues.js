/************** CONFIG – CHANGE IF NEEDED **************/

const ATTENDANCE_SHEET_NAME = 'Weekreport'; // biometric + DeskTime sheet
const WEEK_OFF_DAYS = [0];                  // still here but we no longer rely on it

/******************************************************/

/**
 * Run this – it will auto-pick the correct month sheet
 * and fill codes ONLY for yesterday, then update totals.
 */
function updateCurrentMonth() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  // Yesterday
  const today = new Date();
  const yesterday = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate() - 1
  );

  // Month sheet name like "November"
  const monthName = Utilities.formatDate(yesterday, tz, 'MMMM');
  const monthSheet = ss.getSheetByName(monthName);

  if (!monthSheet) {
    throw new Error('Month sheet "' + monthName + '" not found.');
  }

  fillCodesForMonthSheet_(monthSheet);
}

/**
 * Core logic for one month sheet.
 * Fills abbreviations only for yesterday’s date column,
 * then updates the PTO / P-HD / Leave / UP-HD / WO / R/NJ / P / Total Days
 * columns for all employees.
 */
function fillCodesForMonthSheet_(monthSheet) {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const attendanceSheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
  if (!attendanceSheet) {
    throw new Error('Sheet "' + ATTENDANCE_SHEET_NAME + '" not found.');
  }

  // "Yesterday" in spreadsheet timezone
  const today = new Date();
  const yesterday = new Date(
    today.getFullYear(),
    today.getMonth(),
    today.getDate() - 1
  );

  const HEADER_ROW = 3;             // row that has: S.No. | Name | Status | 01 Nov 25 | ...
  const FIRST_EMP_ROW = HEADER_ROW + 1;

  /************ 1. Build lookup from Attendance sheet ************/

  const attValues = attendanceSheet.getDataRange().getValues();
  if (attValues.length < 2) return;

  const attHeader = attValues[0];

  const idxName   = findHeaderIndex_(attHeader, 'full name');
  const idxDate   = findHeaderIndex_(attHeader, 'date');
  const idxOffSh  = findHeaderIndex_(attHeader, 'office short hours');
  const idxDeskSh = findHeaderIndex_(attHeader, 'desktime short hours');

  if (idxName === -1 || idxDate === -1 || idxOffSh === -1 || idxDeskSh === -1) {
    throw new Error(
      'Check headers on sheet "' + ATTENDANCE_SHEET_NAME +
      '". Need "Full Name", "Date", "Office Short Hours", "Desktime Short Hours".'
    );
  }

  // key: "name|yyyy-MM-dd" -> {officeShort, deskShort}
  const lookup = {};
  for (let i = 1; i < attValues.length; i++) {
    const row = attValues[i];
    const name = row[idxName];
    const date = row[idxDate];
    if (!name || !(date instanceof Date)) continue;

    const key = makeKey_(name, date, tz);
    lookup[key] = {
      officeShort: row[idxOffSh],
      deskShort:   row[idxDeskSh]
    };
  }

  /************ 2. Read month sheet structure ************/

  const lastRow = monthSheet.getLastRow();
  const lastCol = monthSheet.getLastColumn();
  if (lastRow < FIRST_EMP_ROW || lastCol < 4) return;

  // Row 3 has headers
  const headerRow = monthSheet
    .getRange(HEADER_ROW, 1, 1, lastCol)
    .getValues()[0];

  // Find columns "Name" and "Status" in row 3 (exact text, ignoring spaces/case)
  const nameColIndex   = findHeaderIndex_(headerRow, 'name') + 1;   // 1-based
  const statusColIndex = findHeaderIndex_(headerRow, 'status') + 1; // 1-based

  if (nameColIndex === 0 || statusColIndex === 0) {
    throw new Error(
      'Could not find "Name" or "Status" in row 3 of sheet "' +
      monthSheet.getName() + '". Current headers: ' + JSON.stringify(headerRow)
    );
  }

  const firstDateCol = statusColIndex + 1;

  // Find last date column (stop when header is no longer a Date → PTO, P-HD, etc.)
  let lastDateCol = firstDateCol;
  for (let col = firstDateCol; col <= lastCol; col++) {
    const v = headerRow[col - 1];
    if (v instanceof Date) {
      lastDateCol = col;
    } else {
      break;
    }
  }

  const numDays = lastDateCol - firstDateCol + 1;

  // Date headers (row 3, only real dates)
  const dateHeaders = monthSheet
    .getRange(HEADER_ROW, firstDateCol, 1, numDays)
    .getValues()[0];

  // Employee rows (row 4 downwards)
  const numEmployees = lastRow - HEADER_ROW;
  const names  = monthSheet
    .getRange(FIRST_EMP_ROW, nameColIndex,   numEmployees, 1)
    .getValues();
  const status = monthSheet
    .getRange(FIRST_EMP_ROW, statusColIndex, numEmployees, 1)
    .getValues();

  // --- find the column for *yesterday* in this month sheet ---
  const yesterdayIdx = dateHeaders.findIndex(d =>
    d instanceof Date && isSameYMD_(d, yesterday, tz)
  );

  if (yesterdayIdx === -1) {
    // Yesterday is not in this month tab – nothing to do
    Logger.log(
      'Yesterday (' +
      Utilities.formatDate(yesterday, tz, 'dd MMM yyyy') +
      ') not found in sheet "' + monthSheet.getName() + '".'
    );
  } else {
    const targetColIndex = firstDateCol + yesterdayIdx;
    const targetDateCell = dateHeaders[yesterdayIdx];

    // Only one column (yesterday)
    const output = Array(numEmployees).fill().map(() => ['']);

    /************ 3. Fill codes for yesterday only ************/

    for (let r = 0; r < numEmployees; r++) {
      const empName   = names[r][0];
      const empStatus = (status[r][0] || '').toString().trim().toLowerCase();

      if (!empName || !(targetDateCell instanceof Date)) {
        output[r][0] = '';
        continue;
      }

      let code = '';

      // Status-based codes (Not Joined / Resigned)
      if (empStatus === 'not joined') {
        code = 'NJ';
      } else if (empStatus === 'resigned') {
        code = 'R';
      } else if (isWeekOffDate_(targetDateCell)) {
        // <<< NEW: Sundays + 2nd & 4th Saturdays are ALWAYS WO
        code = 'WO';
      } else {
        const key = makeKey_(empName, targetDateCell, tz);
        const rec = lookup[key];

        if (rec) {
          code = getAttendanceCode(rec.officeShort, rec.deskShort);
        } else {
          // Working day but no record for this emp+date
          code = 'NCNS'; // No cause no show
        }
      }

      output[r][0] = code;
    }

    /************ 4. Write back yesterday’s column ************/
    monthSheet
      .getRange(FIRST_EMP_ROW, targetColIndex, numEmployees, 1)
      .setValues(output);
  }

  /************ 5. Update summary columns for the whole month ************/
  updateSummaryColumns_(
    monthSheet,
    headerRow,
    firstDateCol,
    numDays,
    FIRST_EMP_ROW,
    numEmployees
  );
}

/***************** SUMMARY HELPER *****************/

/**
 * Count PTO, P-HD(HD), Leave(L), UP-HD(UPHD), WO, R/NJ, P
 * and fill the summary columns + Total Days.
 */
function updateSummaryColumns_(
  monthSheet,
  headerRow,
  firstDateCol,
  numDays,
  firstEmpRow,
  numEmployees
) {
  // Read all date cells for all employees
  const data = monthSheet
    .getRange(firstEmpRow, firstDateCol, numEmployees, numDays)
    .getValues();

  // Summary column indices (1-based). 0 means "not found / skip".
  const ptoCol        = findHeaderIndex_(headerRow, 'pto')        + 1;
  const phdCol        = findHeaderIndex_(headerRow, 'p-hd')       + 1; // counts "HD"
  const leaveCol      = findHeaderIndex_(headerRow, 'leave')      + 1; // counts "L"
  const uphdCol       = findHeaderIndex_(headerRow, 'up-hd')      + 1; // counts "UPHD"
  const woCol         = findHeaderIndex_(headerRow, 'wo')         + 1; // counts "WO"
  const rnjCol        = findHeaderIndex_(headerRow, 'r/nj')       + 1; // counts "R" + "NJ"
  const pCol          = findHeaderIndex_(headerRow, 'p')          + 1; // counts "P"
  const totalDaysCol  = findHeaderIndex_(headerRow, 'total days') + 1;

  // Prepare output arrays (each is Nx1)
  const ptoVals       = Array(numEmployees).fill().map(() => [0]);
  const phdVals       = Array(numEmployees).fill().map(() => [0]);
  const leaveVals     = Array(numEmployees).fill().map(() => [0]);
  const uphdVals      = Array(numEmployees).fill().map(() => [0]);
  const woVals        = Array(numEmployees).fill().map(() => [0]);
  const rnjVals       = Array(numEmployees).fill().map(() => [0]);
  const pVals         = Array(numEmployees).fill().map(() => [0]);
  const totalDaysVals = Array(numEmployees).fill().map(() => [numDays]); // same for all

  for (let r = 0; r < numEmployees; r++) {
    let cPTO = 0, cPHD = 0, cLeave = 0, cUPHD = 0, cWO = 0, cRNJ = 0, cP = 0;

    for (let c = 0; c < numDays; c++) {
      const v = String(data[r][c] || '').toUpperCase().trim();
      if (!v) continue;

      if (v === 'PTO') cPTO++;
      else if (v === 'HD') cPHD++;
      else if (v === 'L') cLeave++;
      else if (v === 'UPHD') cUPHD++;
      else if (v === 'WO') cWO++;
      else if (v === 'R' || v === 'NJ') cRNJ++;

      if (v === 'P') cP++;
    }

    ptoVals[r][0]   = cPTO;
    phdVals[r][0]   = cPHD;
    leaveVals[r][0] = cLeave;
    uphdVals[r][0]  = cUPHD;
    woVals[r][0]    = cWO;
    rnjVals[r][0]   = cRNJ;
    pVals[r][0]     = cP;
  }

  if (ptoCol)       monthSheet.getRange(firstEmpRow, ptoCol,       numEmployees, 1).setValues(ptoVals);
  if (phdCol)       monthSheet.getRange(firstEmpRow, phdCol,       numEmployees, 1).setValues(phdVals);
  if (leaveCol)     monthSheet.getRange(firstEmpRow, leaveCol,     numEmployees, 1).setValues(leaveVals);
  if (uphdCol)      monthSheet.getRange(firstEmpRow, uphdCol,      numEmployees, 1).setValues(uphdVals);
  if (woCol)        monthSheet.getRange(firstEmpRow, woCol,        numEmployees, 1).setValues(woVals);
  if (rnjCol)       monthSheet.getRange(firstEmpRow, rnjCol,       numEmployees, 1).setValues(rnjVals);
  if (pCol)         monthSheet.getRange(firstEmpRow, pCol,         numEmployees, 1).setValues(pVals);
  if (totalDaysCol) monthSheet.getRange(firstEmpRow, totalDaysCol, numEmployees, 1).setValues(totalDaysVals);
}

/***************** OTHER HELPERS *****************/

// case-insensitive + trimmed header lookup in a 1-D array
function findHeaderIndex_(headerRow, targetText) {
  const target = String(targetText).toLowerCase().trim();
  for (let i = 0; i < headerRow.length; i++) {
    const h = String(headerRow[i]).toLowerCase().trim();
    if (h === target) return i; // 0-based
  }
  return -1;
}

// Compare dates by yyyy-MM-dd in given timezone
function isSameYMD_(d1, d2, tz) {
  return Utilities.formatDate(d1, tz, 'yyyy-MM-dd') ===
         Utilities.formatDate(d2, tz, 'yyyy-MM-dd');
}

// key: "name|yyyy-MM-dd"
function makeKey_(name, date, tz) {
  const namePart = String(name || '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
  const datePart = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  return namePart + '|' + datePart;
}

/**
 * NEW helper: is this date a weekly off?
 * - All Sundays
 * - 2nd and 4th Saturdays
 */
function isWeekOffDate_(date) {
  if (!(date instanceof Date)) return false;

  const dow = date.getDay(); // 0=Sun, 6=Sat
  if (dow === 0) return true; // every Sunday

  if (dow === 6) { // Saturday
    const dom = date.getDate(); // day of month
    const weekOfMonth = Math.floor((dom - 1) / 7) + 1; // 1..5
    if (weekOfMonth === 2 || weekOfMonth === 4) {
      return true; // 2nd or 4th Saturday
    }
  }

  return false;
}

/**
 * Decide code from Office Short Hours + DeskTime Short Hours.
 * Currently only DeskTime (second argument) is used.
 */
function getAttendanceCode(officeShort, deskShort) {
  var d = 0;

  // If the cell is a time (Date object), convert HH:MM:SS → hours
  if (deskShort instanceof Date) {
    d = deskShort.getHours()
        + deskShort.getMinutes() / 60
        + deskShort.getSeconds() / 3600;
  } else {
    // If it's already numeric (hours), use it directly
    d = Number(deskShort);
    if (isNaN(d)) d = 0;
  }

  // Your rules in HOURS:
  // <= 15 min (0.25 hr)              -> P
  // > 15 min and <= 3h 15m (3.25 hr) -> HD
  // > 3h 15m and <= 5h               -> UPHD
  // > 5h                             -> L

  var FULL_PRESENT_LIMIT = 0.25; // 15 minutes
  var HALF_DAY_LIMIT     = 3.25; // 3 hours 15 minutes
  var UNPAID_HALF_LIMIT  = 5;    // 5 hours

  if (d <= FULL_PRESENT_LIMIT) {
    return 'P';    // Present
  }
  if (d <= HALF_DAY_LIMIT) {
    return 'HD';   // Half Day
  }
  if (d <= UNPAID_HALF_LIMIT) {
    return 'UPHD'; // Unpaid Half Day
  }
  return 'L';      // Leave
}