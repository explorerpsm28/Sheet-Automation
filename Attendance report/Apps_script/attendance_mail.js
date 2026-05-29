/**
 * Per-employee weekly attendance mailer (Mon–Sat, most recent, IST):
 * - Reads "Attendance" sheet.
 * - Groups rows by Email for the latest Mon–Sat week.
 * - Sends one email per employee with their own table.
 * - CC rules:
 *     @offshoreally.com -> translation@offshoreally.com, s2@corptive.com, sonalirana@nlink.tech
 *     @nlink.tech      -> s2@corptive.com, sonalirana@nlink.tech
 * - Uses same banner / footer / disclaimer layout as before.
 */
function sendAttendanceReportEmails() {
  // ---- CONFIG ----
  var SHEET_NAME = 'Weekreport';
  var IST_TZ = 'Asia/Kolkata';

  // === Compute most recent Mon–Sat in IST ===
  var weekInfo = getMostRecentWeekIST();
  var PERIOD_TEXT = weekInfo.text;        // e.g. "10 November '25 - 15 November '25"
  var WEEK_START = weekInfo.monday;       // Date (IST)
  var WEEK_END   = weekInfo.saturday;     // Date (IST)

  var SUBJECT = 'Weekly Report — Biometric Attendance & DeskTime (' + PERIOD_TEXT + ')';
  
  // Emails to exclude from processing/sending
  const SKIP_EMAILS = new Set(['varun@corptive.com']);

  // Inline image Drive IDs
  const BANNER_FILE_ID = '1GLvDYAB8w_K81MvzWQvhpeB7zWrGOKLb';
  const TECH_FILE_ID   = '1_8T73h0rCtitBRLfgQ4ajgqVEaNiRiGh';

  // CID keys
  const BANNER_CID    = 'banner';
  const TECH_LOGO_CID = 'techlogo';

  // Preload blobs once
  const inlineImagesBase = {
    [BANNER_CID]:    DriveApp.getFileById(BANNER_FILE_ID).getBlob(),
    [TECH_LOGO_CID]: DriveApp.getFileById(TECH_FILE_ID).getBlob(),
  };

  // Sheet link (if you want to add it in the email later)
  var SPREADSHEET_URL = '';

  // ---- HTML Blocks (same look as your previous mail) ----

  const BODY_OPEN = `
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
      <tr><td align="center" style="padding:0 20px;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="max-width:560px;">
          <tr><td style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:20px;color:#111;">
  `;
  const BODY_CLOSE = `
          </td></tr>
        </table>
      </td></tr>
    </table>
  `;

  // ---- Read Attendance sheet ----
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    throw new Error('No data rows in sheet "' + SHEET_NAME + '".');
  }

  // Read values AND background colors
  var dataRange   = sheet.getRange(1, 1, lastRow, lastCol);
  var values      = dataRange.getValues();
  var backgrounds = dataRange.getBackgrounds();

  var headers = values[0];
  var data    = values.slice(1);        // from row 2
  var bgData  = backgrounds.slice(1);   // backgrounds for row 2+

  var nameColIndex  = headers.indexOf('Full Name');
  var emailColIndex = headers.indexOf('Email');
  var dateColIndex  = headers.indexOf('Date');

  if (nameColIndex === -1 || emailColIndex === -1 || dateColIndex === -1) {
    throw new Error('Headers "Full Name", "Email", and "Date" are required in the first row.');
  }

  // ---- Group rows by employee email for this Mon–Sat week ----
  var groups = {};  // emailLower -> { email, name, rows[] }
  var skippedRows = 0;

  data.forEach(function (row, idx) {
    var email = (row[emailColIndex] || '').toString().trim();
    if (!email) {
      skippedRows++;
      return;
    }

    // ✅ Skip specific email(s)
    if (SKIP_EMAILS.has(email.toLowerCase())) {
      skippedRows++;
      return;
    }

    var d = row[dateColIndex];
    if (!(d instanceof Date) || isNaN(d)) {
      skippedRows++;
      return;
    }

    // Normalize date to date-only (ignore time part)
    var dNorm = new Date(d.getFullYear(), d.getMonth(), d.getDate());

    if (dNorm < WEEK_START || dNorm > WEEK_END) {
      // Not in the current Mon–Sat window
      skippedRows++;
      return;
    }

    var key = email.toLowerCase();
    if (!groups[key]) {
      groups[key] = {
        email: email,
        name: row[nameColIndex] || '',
        rows: []   // each item: { values: [...], backgrounds: [...] }
      };
    }

    groups[key].rows.push({
      values: row,
      backgrounds: bgData[idx] || []
    });
  });

  // ---- Build and send one email per employee ----
  var sentCount = 0;

  Object.keys(groups).forEach(function (key) {
    var group = groups[key];
    if (!group.rows.length) return;

    // Sort by date ascending
    group.rows.sort(function (a, b) {
      var da = a.values[dateColIndex];
      var db = b.values[dateColIndex];
      if (da && db) return da - db;
      return 0;
    });

    var toAddress = group.email;
    var ccAddress = getManagerCcForEmail(toAddress);

    var tableHtml = buildAttendanceTableHtml(headers, group.rows, dateColIndex, IST_TZ);

    var greetingName = group.name ? escapeHtml(group.name) : 'Team';

    var body =
      BODY_OPEN +
      `<p>Dear ${greetingName},</p>
       <p>I hope this email finds you well.</p>
       <p>We are sharing your weekly report that combines insights from both Biometric Attendance and DeskTime Analysis for the period <strong>${PERIOD_TEXT}</strong>.</p>
       <p>Below is your detailed attendance and DeskTime summary for this period. Yellow highlights in the sheet indicate approved buffer times.</p>
       ${tableHtml}
       <p>You can also review the complete report if needed.</p>
       <p>If you have any questions or suggestions, feel free to reach out.</p>
       <p>Thank you.</p>` +
      BODY_CLOSE;

    MailApp.sendEmail({
      to: toAddress,
      cc: ccAddress,
      subject: SUBJECT,
      htmlBody: body,
      name: 'Nlink Tech'
      // If you actually embed the banner/logo in HTML, add:
      // ,inlineImages: inlineImagesBase
    });

    sentCount++;
  });

  // ---- ONE-LINE EXECUTION LOG TOAST ----
  var tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || "UTC";
  var stamp = Utilities.formatDate(new Date(), tz, "h:mm:ss a");
  ss.toast(
    stamp + '  Info  Done. Sent emails: ' + sentCount + ', Skipped rows (no email / out of week): ' + skippedRows,
    'Execution Log',
    10
  );

  Logger.log('Per-employee attendance emails sent: ' + sentCount);
}

/**
 * Manager CC selection based on employee email domain.
 *  - @offshoreally.com -> translation@offshoreally.com, s2@corptive.com, sonalirana@nlink.tech
 *  - @nlink.tech      -> s2@corptive.com, sonalirana@nlink.tech
 */
function getManagerCcForEmail(email) {
  if (!email) return '';
  var e = email.toLowerCase().trim();

  if (e.endsWith('@offshoreally.com')) {
    return [
      's2@corptive.com',
      'sonalirana@nlink.tech'
    ].join(',');
  }

  if (e.endsWith('@nlink.tech')) {
    return [
      's2@corptive.com',
      'sonalirana@nlink.tech'
    ].join(',');
  }

  // For any other domain, no manager CC defined
  return '';
}

/**
 * Build HTML table for the employee's rows.
 * This version carries over background colors from the sheet.
 */
function buildAttendanceTableHtml(headers, rowObjs, dateColIndex, tz) {
  var html = ''
    + '<table role="presentation" cellspacing="0" cellpadding="4" border="0"'
    + ' style="border-collapse:collapse;width:100%;max-width:700px;margin:16px 0;'
    + ' font-family:Arial,Helvetica,sans-serif;font-size:12px;line-height:18px;">';

  // Header row
  html += '<tr>';
  for (var c = 0; c < headers.length; c++) {
    html += '<th style="border:1px solid #c0c0c0;padding:6px 8px;background-color:#dde5fb;text-align:center;">'
      + escapeHtml(headers[c])
      + '</th>';
  }
  html += '</tr>';

  // Data rows (with background colors)
  rowObjs.forEach(function (rowObj) {
    var rowValues = rowObj.values;
    var rowBgs    = rowObj.backgrounds || [];

    html += '<tr>';

    for (var c = 0; c < headers.length; c++) {
      var val       = rowValues[c];
      var isDateCol = (c === dateColIndex);
      var text      = formatCellForHtml(val, tz, isDateCol);

      var cellStyle = 'border:1px solid #e0e0e0;padding:4px 6px;text-align:center;white-space:nowrap;';

      // apply background from sheet if not plain white
      var bg = rowBgs[c];
      if (bg && bg !== '#ffffff' && bg !== '#fff' && bg !== '') {
        cellStyle += 'background-color:' + bg + ';';
      }

      html += '<td style="' + cellStyle + '">'
        + escapeHtml(text)
        + '</td>';
    }

    html += '</tr>';
  });

  html += '</table>';
  return html;
}

/**
 * Format a sheet cell value for display in HTML.
 */
function formatCellForHtml(value, tz, isDateCol) {
  if (value === null || value === '') return '';

  if (value instanceof Date && !isNaN(value)) {
    if (isDateCol) {
      // Date column: show full date
      return Utilities.formatDate(value, tz, 'dd-MMM-yy');
    } else {
      // Time columns: show time only
      return Utilities.formatDate(value, tz, 'HH:mm');
    }
  }

  return String(value);
}

/**
 * Escape HTML special characters.
 */
function escapeHtml(str) {
  return String(str || '').replace(/[&<>"']/g, function (c) {
    return ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    })[c];
  });
}

/**
 * Compute the most recent Mon–Sat week in IST, returning:
 *  { monday: Date, saturday: Date, text: "d MMMM 'yy - d MMMM 'yy" }
 * Logic is based on your original PERIOD_TEXT code.
 */
function getMostRecentWeekIST() {
  var IST_OFFSET_MS = 5.5 * 60 * 60 * 1000; // +05:30
  var nowUtcMs = new Date().getTime();
  var istNow = new Date(nowUtcMs + IST_OFFSET_MS); // "shifted" to IST

  // Day-of-week in IST using UTC getters (Sun=0..Sat=6)
  var dow = istNow.getUTCDay();

  // Most recent Saturday at/before today
  var daysSinceSat = (dow - 6 + 7) % 7;

  // Midnight (IST) for "today" in this shifted timeline
  var istY = istNow.getUTCFullYear();
  var istM = istNow.getUTCMonth();
  var istD = istNow.getUTCDate();
  var istTodayMidnightFakeUTC = Date.UTC(istY, istM, istD);

  var dayMs = 24 * 60 * 60 * 1000;
  var lastSatFakeUTC = istTodayMidnightFakeUTC - daysSinceSat * dayMs;
  var monFakeUTC = lastSatFakeUTC - 5 * dayMs;

  // Convert back to real epoch ms (undo the shift) and format in IST
  var mondayReal = new Date(monFakeUTC - IST_OFFSET_MS);
  var saturdayReal = new Date(lastSatFakeUTC - IST_OFFSET_MS);

  var fromStr = Utilities.formatDate(mondayReal, 'Asia/Kolkata', "d MMMM ''yy");
  var toStr   = Utilities.formatDate(saturdayReal, 'Asia/Kolkata', "d MMMM ''yy");

  // Strip time portion: normalize to date-only for comparisons
  var monNorm = new Date(mondayReal.getFullYear(), mondayReal.getMonth(), mondayReal.getDate());
  var satNorm = new Date(saturdayReal.getFullYear(), saturdayReal.getMonth(), saturdayReal.getDate());

  return {
    monday: monNorm,
    saturday: satNorm,
    text: fromStr + ' - ' + toStr
  };
}
