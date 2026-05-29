function sendMonthlyIncentivesUnified() {
  /***** CONFIG *****/
  const YEAR = null; // set null if you want current year automatically
  const SENDER_NAME = "Finance Team";
  const SUBJECT_PREFIX = "Incentive Announcement – ";
  const CC_LIST = "translation@offshoreally.com,nitin@offshoreally.com";
  const DRY_RUN = false;

  const USE_PENALTY_COLUMN = true;

  // Fallback only if Penalty column is not present
  const PENALTY_MAP = {
    "Sakshi Nautiyal": 500,
    "Rahul Dhiman": 500,
    "Jyoti Koteri": 500
  };

  const MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || "Asia/Kolkata";
  const now = new Date();

  const runMonthIndex = Number(Utilities.formatDate(now, tz, "M")) - 1;
  const runYear = YEAR || Number(Utilities.formatDate(now, tz, "yyyy"));

  // Pick previous completed quarter
  const sendInfo = getQuarterToSend_(runMonthIndex, runYear);

  if (!sendInfo) {
    ss.toast(
      "This is not a quarterly send month. Run this only in January, April, July, or October.",
      "Execution Log",
      8
    );
    return;
  }

  const sheet = ss.getSheetByName(sendInfo.sheetName);
  if (!sheet) {
    throw new Error(`Quarter sheet "${sendInfo.sheetName}" not found.`);
  }

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    ss.toast(`No data rows found in ${sendInfo.sheetName}.`, "Execution Log", 6);
    return;
  }

  const headers = values[0].map(h => String(h || "").trim());
  const idx = {
    name: findHeader_(headers, ["employee name", "name"]),
    email: findHeader_(headers, ["email"]),
    total: findHeaderRegex_(headers, /^total$/i),
    month: findHeader_(headers, ["month"]),
    penalty: findHeaderRegex_(headers, /^penalty$/i, true)
  };

  const grouped = {};

  for (let r = 1; r < values.length; r++) {
    const row = values[r];

    const empName = safe_(row[idx.name]);
    const toEmail = safe_(row[idx.email]);
    const monthName = normalizeMonthName_(row[idx.month]);

    if (!empName || !toEmail || !monthName) continue;

    const monthIndex = MONTHS.indexOf(monthName);
    if (monthIndex === -1) continue;

    // Only take rows from the full completed quarter
    if (monthIndex < sendInfo.startIndex || monthIndex > sendInfo.endIndex) continue;

    const key = `${toEmail}__${empName}`;

    if (!grouped[key]) {
      grouped[key] = {
        name: empName,
        email: toEmail,
        rows: [],
        penaltyTotal: 0
      };
    }

    grouped[key].rows.push({
      employeeName: empName,
      total: safe_(row[idx.total]),
      month: monthName,
      monthIndex: monthIndex
    });

    let penaltyValue = "";
    if (USE_PENALTY_COLUMN && idx.penalty !== -1) {
      penaltyValue = safe_(row[idx.penalty]);
    } else if (Object.prototype.hasOwnProperty.call(PENALTY_MAP, empName)) {
      penaltyValue = String(PENALTY_MAP[empName]);
    }

    grouped[key].penaltyTotal += toNumber_(penaltyValue);
  }

  const periodShort = buildPeriodLabel_(sendInfo.startIndex, sendInfo.endIndex, sendInfo.dataYear, true);
  const periodLong = buildPeriodLabel_(sendInfo.startIndex, sendInfo.endIndex, sendInfo.dataYear, false);
  const subject = SUBJECT_PREFIX + periodShort;

  const employeeKeys = Object.keys(grouped);
  if (!employeeKeys.length) {
    ss.toast(`No matching employee rows found in ${sendInfo.sheetName} for ${periodShort}.`, "Execution Log", 8);
    return;
  }

  let sent = 0;
  let skipped = 0;

  ss.toast(`Starting ${subject}${DRY_RUN ? " (DRY RUN)" : ""}…`, "Execution Log", 4);

  employeeKeys.forEach(key => {
    const emp = grouped[key];

    if (!emp.rows.length) {
      skipped++;
      return;
    }

    emp.rows.sort((a, b) => a.monthIndex - b.monthIndex);

    const htmlBody = buildEmailHtml_({
      employeeName: emp.name,
      rows: emp.rows,
      periodLabel: periodLong,
      penaltyTotal: emp.penaltyTotal
    });

    if (DRY_RUN) {
      Logger.log(`[DRY RUN] To: ${emp.email} | Subject: ${subject}`);
      sent++;
    } else {
      MailApp.sendEmail({
        to: emp.email,
        cc: CC_LIST,
        subject: subject,
        htmlBody: htmlBody,
        name: SENDER_NAME
      });
      Utilities.sleep(150);
      sent++;
    }
  });

  const stamp = Utilities.formatDate(new Date(), tz, "h:mm:ss a");
  ss.toast(`${stamp}  Info  Done. Sent: ${sent}, Skipped: ${skipped}`, "Execution Log", 10);

  /***** HELPERS *****/

  function getQuarterToSend_(runMonthIndex, runYear) {
    // January -> send Oct-Dec of previous year
    if (runMonthIndex === 0) {
      return { sheetName: "Oct-Dec", startIndex: 9, endIndex: 11, dataYear: runYear - 1 };
    }

    // April -> send Jan-Mar
    if (runMonthIndex === 3) {
      return { sheetName: "Jan-Mar", startIndex: 0, endIndex: 2, dataYear: runYear };
    }

    // July -> send Apr-Jun
    if (runMonthIndex === 6) {
      return { sheetName: "Apr-Jun", startIndex: 3, endIndex: 5, dataYear: runYear };
    }

    // October -> send July-Sept
    if (runMonthIndex === 9) {
      return { sheetName: "July-Sept", startIndex: 6, endIndex: 8, dataYear: runYear };
    }

    return null;
  }

  function buildPeriodLabel_(startIndex, endIndex, year, shortName) {
    const shortMonths = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sept", "Oct", "Nov", "Dec"];
    const names = shortName ? shortMonths : MONTHS;
    const apostrophe = "’";
    return `${names[startIndex]}-${names[endIndex]} ${apostrophe}${String(year).slice(-2)}`;
  }

  function normalizeMonthName_(value) {
    const raw = String(value || "").trim().toLowerCase();
    if (!raw) return "";

    const map = {
      jan: "January", january: "January",
      feb: "February", february: "February",
      mar: "March", march: "March",
      apr: "April", april: "April",
      may: "May",
      jun: "June", june: "June",
      jul: "July", july: "July",
      aug: "August", august: "August",
      sep: "September", sept: "September", september: "September",
      oct: "October", october: "October",
      nov: "November", november: "November",
      dec: "December", december: "December"
    };

    return map[raw] || map[raw.slice(0, 3)] || "";
  }

  function findHeader_(headerRow, names, optional) {
    const wanted = names.map(s => String(s).toLowerCase().trim());
    for (let i = 0; i < headerRow.length; i++) {
      const header = String(headerRow[i] || "").toLowerCase().trim();
      if (wanted.indexOf(header) !== -1) return i;
    }
    if (optional) return -1;
    throw new Error(`Missing required header. Need one of: ${names.join(", ")}`);
  }

  function findHeaderRegex_(headerRow, regex, optional) {
    for (let i = 0; i < headerRow.length; i++) {
      const header = String(headerRow[i] || "").trim();
      if (regex.test(header)) return i;
    }
    if (optional) return -1;
    throw new Error(`Missing required header matching ${regex}`);
  }

  function safe_(value) {
    return value == null ? "" : String(value).trim();
  }

  function toNumber_(value) {
    const cleaned = String(value || "")
      .replace(/,/g, "")
      .replace(/[^\d.-]/g, "");
    const num = Number(cleaned);
    return isFinite(num) ? num : 0;
  }

  function formatAmount_(num) {
    return Number(num || 0).toLocaleString("en-IN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });
  }

  function esc_(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function buildEmailHtml_({ employeeName, rows, periodLabel, penaltyTotal }) {
    const baseCss = "font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #111; line-height: 1.6;";
    const tableCss = "border-collapse: collapse; min-width: 650px; border: 1px solid #e5e5e5;";
    const thCss = "padding: 10px 14px; border: 1px solid #d9d9d9; text-align: left; background:#f7f7f7; font-weight:700;";
    const tdCss = "padding: 10px 14px; border: 1px solid #e5e5e5; text-align: left;";

    const tableRows = rows.map(item => `
      <tr>
        <td style="${tdCss}">${esc_(item.employeeName)}</td>
        <td style="${tdCss}">${esc_(item.total || "-")}</td>
        <td style="${tdCss}">${esc_(item.month)}</td>
      </tr>
    `).join("");

    const showPenalty = penaltyTotal > 0;

    return `
      <div style="${baseCss}">
        <p>Dear ${esc_(employeeName)},</p>

        <p>We have fantastic news! We are thrilled to tell you that you have earned incentives as a token of appreciation for your exceptional performance.</p>

        <p>Your hard work, dedication and outstanding contributions have not gone unnoticed. As a gesture of gratitude, we are pleased to offer you the incentives based on the number of projects completed by you in the quarter of <b>${esc_(periodLabel)}</b>. The details are as under:</p>

        <table style="${tableCss}">
          <thead>
            <tr>
              <th style="${thCss}">Employee Name</th>
              <th style="${thCss}">Month</th>
              <th style="${thCss}">Total</th>
              
            </tr>
          </thead>
          <tbody>
            ${tableRows}
          </tbody>
        </table>

        ${showPenalty ? `
          <div style="margin-top:16px;">
            <p><b>Penalty details are as follows:</b><br>${esc_(employeeName)} : INR ${esc_(formatAmount_(penaltyTotal))}</p>
            
          </div>
        ` : ""}

        <p>The incentive is our way of recognizing your valuable contributions to the team’s and the company’s success. We truly appreciate your efforts and your hard work.</p>

        <p>Thank You.</p>
      </div>`;
  }
}