/********** CONFIG **********/
const LT_SHEET_NAME = 'Leavetracker';   // your sheet tab name
const LT_FIRST_EMP_ROW = 2;             // first row with an employee
const LT_START_YEAR = 2025;             // fiscal year start (April 2025)
const LT_START_MONTH = 4;               // 4 = April
const LT_START_BAL_COL = 3;             // column C  ("As on 01.04.2025")
const LT_START_ALLOT_COL = 4;           // column D  ("Leave Allotment - Apr 2025")
const LT_MONTHS_IN_YEAR = 12;           // Apr → Mar
const LT_MONTHLY_LEAVE = 1.5;           // 1.5 days each month
/********************************/

/**
 * Unified function:
 *  - For every month from April 2025 up to the current month
 *    * fills Leave Allotment with 1.5 if blank/0
 *    * calculates "Leaves On 01.xx" from:
 *        start balance + allotment - paid leaves availed
 *
 * Run this manually, or connect a monthly trigger to it.
 */
function ltUpdateLeaves() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LT_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + LT_SHEET_NAME + '" not found.');
  }

  const today = new Date();
  const currentMonth = today.getMonth() + 1;   // 1–12
  const currentYear = today.getFullYear();

  // 0 = Apr, 1 = May, ... 11 = Mar
  let lastIndex = (currentYear - LT_START_YEAR) * 12 + (currentMonth - LT_START_MONTH);
  if (lastIndex < 0) return; // before April 2025 -> nothing to do
  if (lastIndex >= LT_MONTHS_IN_YEAR) lastIndex = LT_MONTHS_IN_YEAR - 1;

  const lastRow = sheet.getLastRow();
  if (lastRow < LT_FIRST_EMP_ROW) return;
  const numRows = lastRow - LT_FIRST_EMP_ROW + 1;

  // Get all employee names once (column B)
  const names = sheet
    .getRange(LT_FIRST_EMP_ROW, 2, numRows, 1)
    .getValues();

  // Loop through each month from April up to current month
  for (let idx = 0; idx <= lastIndex; idx++) {
    const startBalCol = ltGetStartBalanceCol(idx); // starting balance for this month
    const allotCol = ltGetAllotCol(idx);           // Leave Allotment for this month
    const availedCol = ltGetAvailedCol(idx);       // Paid leaves availed this month
    const asOnNextCol = ltGetAsOnNextCol(idx);     // Leaves On 01.(next month)

    const startBalances = sheet
      .getRange(LT_FIRST_EMP_ROW, startBalCol, numRows, 1)
      .getValues();
    const allotments = sheet
      .getRange(LT_FIRST_EMP_ROW, allotCol, numRows, 1)
      .getValues();
    const availed = sheet
      .getRange(LT_FIRST_EMP_ROW, availedCol, numRows, 1)
      .getValues();

    const newAllotments = [];
    const closingBalances = [];

    for (let r = 0; r < numRows; r++) {
      const name = names[r][0];
      if (!name) {
        // Empty line → keep row blank
        newAllotments.push(['']);
        closingBalances.push(['']);
        continue;
      }

      const start = Number(startBalances[r][0]) || 0;

      // Leave Allotment: if blank or 0, default to LT_MONTHLY_LEAVE
      let allot = Number(allotments[r][0]);
      if (!allot) allot = LT_MONTHLY_LEAVE;
      newAllotments.push([allot]);

      // Paid leaves availed for this month (manual entry)
      const used = Number(availed[r][0]) || 0;

      // Closing balance = start balance + allotment - used
      const closeBal = start + allot - used;
      closingBalances.push([closeBal]);
    }

    // Write updated allotments and new "Leaves On 01.xx" balances
    sheet
      .getRange(LT_FIRST_EMP_ROW, allotCol, numRows, 1)
      .setValues(newAllotments);
    sheet
      .getRange(LT_FIRST_EMP_ROW, asOnNextCol, numRows, 1)
      .setValues(closingBalances);
  }
}

/********** Helper functions (for column positions) **********/

// For month index 0..11 (0 = Apr 2025)
function ltGetAllotCol(idx) {
  return LT_START_ALLOT_COL + idx * 3;      // D, G, J, ...
}
function ltGetAvailedCol(idx) {
  return LT_START_ALLOT_COL + idx * 3 + 1;  // E, H, K, ...
}
function ltGetAsOnNextCol(idx) {
  return LT_START_ALLOT_COL + idx * 3 + 2;  // F, I, L, ...
}
function ltGetStartBalanceCol(idx) {
  if (idx === 0) return LT_START_BAL_COL;   // April uses initial "As on 01.04.2025" (C)
  // From May onwards, start balance is previous month’s "Leaves On 01.xx"
  return ltGetAsOnNextCol(idx - 1);
}

/********** Optional: trigger that runs this once a month **********/

// Run this ONCE from the Script Editor to create a monthly trigger
function ltCreateMonthlyTrigger() {
  ScriptApp.newTrigger('ltUpdateLeaves')
    .timeBased()
    .onMonthDay(1)   // 1st of every month
    .atHour(1)       // at 01:00 (change if you like)
    .create();
}
