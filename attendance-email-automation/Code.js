/**
 * @fileoverview Google Apps Script to automate attendance email reporting.
 * It groups attendance data by department, generates an HTML table and a CSV
 * attachment for each, and sends a single email to the specified department
 * recipients.
 */

// --- Global Constants ---
const ATTENDANCE_SHEET_NAME = "Attendance";
const DEPT_EMAILS_SHEET_NAME = "dept_emails";
// const CC_EMAILS = "okkar@marathonmyanmar.com, aungkoswe@marathonmyanmar.com, hr@marathonmyanmar.com"; // Optional: Add a constant CC list here. Leave empty if not needed.
const CC_EMAILS = "soeyarzar@marathonmyanmar.com";
const TIMEZONE = "Asia/Yangon";

// --- Main Functions ---

/**
 * On spreadsheet open, creates a custom menu to trigger the script.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Attendance")
    .addItem("Send per Department", "sendAttendancePerDepartment")
    .addToUi();
}

/**
 * The main function to send attendance reports.
 * It validates sheets, reads data, groups by department, and sends
 * a separate email for each department with a valid email mapping.
 */
function sendAttendancePerDepartment() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);
    const deptEmailsSheet = spreadsheet.getSheetByName(DEPT_EMAILS_SHEET_NAME);

    // 1. Validate sheets and headers
    if (!attendanceSheet || !deptEmailsSheet) {
      ui.alert("Error: Missing required sheets!", `Please ensure both "${ATTENDANCE_SHEET_NAME}" and "${DEPT_EMAILS_SHEET_NAME}" sheets exist.`, ui.ButtonSet.OK);
      return;
    }

    const attendanceData = getAttendanceRows_(attendanceSheet);
    if (attendanceData.length === 0) {
      ui.alert("No attendance data found in the Attendance sheet. Exiting.");
      return;
    }

    // 2. Group data and get email mappings
    const groupedData = groupRowsByDepartment_(attendanceData);
    const deptEmailsMap = getDeptEmails_(deptEmailsSheet);

    // 3. Process and send emails
    let departmentsProcessed = 0;
    let emailsSent = 0;
    let departmentsSkippedNoData = 0;
    let departmentsSkippedNoEmail = 0;

    // Get all unique departments from the attendance data
    const departments = Object.keys(groupedData);

    for (const department of departments) {
      departmentsProcessed++;
      const departmentRows = groupedData[department];
      const recipients = deptEmailsMap[department];

      if (departmentRows.length === 0) {
        departmentsSkippedNoData++;
        Logger.log(`Skipped department '${department}': No attendance rows.`);
        continue;
      }

      if (!recipients || recipients.trim() === "") {
        departmentsSkippedNoEmail++;
        Logger.log(`Skipped department '${department}': No email mapping found.`);
        continue;
      }

      try {
        const htmlTable = buildHtmlTable_(departmentRows);
        const dateString = Utilities.formatDate(new Date(), TIMEZONE, "YYYYMMdd");
        // const dateDisplay = Utilities.formatDate(new Date(), TIMEZONE, "YYYY-MM-dd");
        // const dateDisplay = spreadsheet.getRange("E2").getValue();
        const csvBlob = toCsvBlob_(departmentRows, `Attendance_${department}_${dateString}.csv`);

        sendDepartmentEmail_(department, recipients, htmlTable, csvBlob);
        emailsSent++;
        Logger.log(`Successfully sent email for department '${department}' to ${recipients}.`);

      } catch (e) {
        Logger.log(`Failed to send email for department '${department}'. Error: ${e.message}`);
        // Continue to the next department even if one fails
      }
    }

    // 4. Log summary
    Logger.log("--- Email Sending Summary ---");
    Logger.log(`Total departments processed: ${departmentsProcessed}`);
    Logger.log(`Emails successfully sent: ${emailsSent}`);
    Logger.log(`Departments skipped (no attendance data): ${departmentsSkippedNoData}`);
    Logger.log(`Departments skipped (no email mapping): ${departmentsSkippedNoEmail}`);

    ui.alert("Script complete!", "Check the logs for a detailed summary.", ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(`A fatal error occurred: ${e.message}`);
    ui.alert("An unexpected error occurred.", e.message, ui.ButtonSet.OK);
  }
}

// --- Helper Functions ---

/**
 * Retrieves and formats attendance data from the specified sheet.
 * This version is updated to find the header row dynamically.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Attendance sheet.
 * @returns {Object[]} An array of objects, where each object represents a row.
 * @private
 */
function getAttendanceRows_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return []; // No data rows

  // Find the header row by looking for a key header like "Department"
  let headerRowIndex = -1;
  for (let i = 0; i < values.length; i++) {
    const rowValues = values[i];
    if (rowValues.some(cell => cell && cell.toString().trim() === "Department")) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) {
    throw new Error("Missing required headers in the Attendance sheet. Please ensure 'Department' is present in row 1.");
  }

  const headers = values[headerRowIndex].map(h => h.toString().trim());
  const data = values.slice(headerRowIndex + 1).filter(row => row.some(cell => cell !== "")); // Filter out completely blank rows

  const headerMap = {
    "Department": headers.indexOf("Department"),
    "Name": headers.indexOf("Name"),
    "Employee ID": headers.indexOf("Employee ID"),
    "AM Face Scan": headers.indexOf("AM Face Scan"),
    "PM Face Scan": headers.indexOf("PM Face Scan"),
    "Attendance": headers.indexOf("Attendance"),
    "HR - Adj": headers.indexOf("HR - Adj"),
    "KPI Leave": headers.indexOf("KPI Leave"),
    "KPI (%)": headers.indexOf("KPI (%)")
  };

  // Check for -1 which indicates a missing header
  if (Object.values(headerMap).some(idx => idx === -1)) {
    throw new Error("Missing one or more required headers in the Attendance sheet.");
  }

  const records = [];
  for (const row of data) {
    records.push({
      "Department": row[headerMap["Department"]] ? row[headerMap["Department"]].toString().trim() : "",
      "Name": row[headerMap["Name"]] ? row[headerMap["Name"]].toString().trim() : "",
      "Employee ID": row[headerMap["Employee ID"]] ? row[headerMap["Employee ID"]].toString().trim() : "",
      "AM Face Scan": row[headerMap["AM Face Scan"]] ? row[headerMap["AM Face Scan"]].toString().trim() : "",
      "PM Face Scan": row[headerMap["PM Face Scan"]] ? row[headerMap["PM Face Scan"]].toString().trim() : "",
      "Attendance": row[headerMap["Attendance"]] ? row[headerMap["Attendance"]].toString().trim() : "",
      "HR - Adj": row[headerMap["HR - Adj"]] ? row[headerMap["HR - Adj"]].toString().trim() : "",
      "KPI Leave": row[headerMap["KPI Leave"]] ? row[headerMap["KPI Leave"]].toString().trim() : "",
      "KPI (%)": row[headerMap["KPI (%)"]] ? row[headerMap["KPI (%)"]].toString().trim() : "",
    });
  }
  return records;
}

/**
 * Groups an array of attendance row objects by department.
 * @param {Object[]} rows The attendance rows.
 * @returns {{ [dept: string]: Object[] }} A map of departments to their attendance rows.
 * @private
 */
function groupRowsByDepartment_(rows) {
  return rows.reduce((acc, row) => {
    const department = row.Department;
    if (department) {
      if (!acc[department]) {
        acc[department] = [];
      }
      acc[department].push(row);
    }
    return acc;
  }, {});
}

/**
 * Retrieves department email mappings from the "DeptEmails" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The DeptEmails sheet.
 * @returns {{ [dept: string]: string }} A map of departments to their email recipients.
 * @private
 */
function getDeptEmails_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return {};

  const headers = values.shift().map(h => h.toString().trim());
  const headerMap = {
    "Department": headers.indexOf("Department"),
    "Emails": headers.indexOf("Emails")
  };

  if (Object.values(headerMap).some(idx => idx === -1)) {
    throw new Error("Missing required headers in the DeptEmails sheet. Please ensure 'Department' and 'Emails' are present in row 1.");
  }

  const emailMap = {};
  for (const row of values) {
    const department = row[headerMap.Department] ? row[headerMap.Department].toString().trim() : "";
    const emails = row[headerMap.Emails] ? row[headerMap.Emails].toString().trim() : "";
    if (department) {
      emailMap[department] = emails;
    }
  }
  return emailMap;
}

/**
 * Builds an HTML table from an array of attendance row objects.
 * @param {Object[]} rows The attendance rows for one department.
 * @returns {string} An HTML string representing the table.
 * @private
 */

function buildHtmlTable_(rows) {
  if (rows.length === 0) return "";
  const headers = Object.keys(rows[0]);

  let html = `
    <table style="width: 849px; table-layout: fixed; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px;">
    <thead>
      <tr style="background-color: #343a40; color: #ffffff;">`;

  // Filter out the 'Department' header for the table
  const displayHeaders = headers.filter(header => header !== "Department");
  // for (const header of displayHeaders) {
  //   // Left-align 'Name', center-align everything else
  //   const textAlign = header === "Name" ? "left" : "center";
  //   html += `<th style="border: 1px solid #dee2e6; width: 80px; padding: 8px; text-align: ${textAlign};">${header}</th>`;
  // }
  for (let i = 0; i < displayHeaders.length; i++) {
    const header = displayHeaders[i];
    // Check if the current column is between index 2 and 6 (3rd to 7th column)
    let columnWidth = '';

    switch (i) {
      case 0:
        columnWidth = 'width: 166px;';
        break;
      case 1:
        columnWidth = 'width: 133px;';
        break;
      case 7: // This corresponds to the 9th column
        columnWidth = 'width: 100px;';
        break;
      default:
        // This applies to columns 3 through 8 (indices 2 through 7)
        if (i >= 2 && i <= 6) {
          columnWidth = 'width: 90px;';
        }
        break;
    }
    
    // Left-align 'Name', center-align everything else
    const textAlign = header === "Name" ? "left" : "center";
    
    html += `<th style="border: 1px solid #dee2e6; padding: 8px; text-align: ${textAlign}; ${columnWidth}">${header}</th>`;
  }


  html += `
      </tr>
    </thead>
    <tbody>`;

  rows.forEach(function(row, index) {
    const rowBackgroundColor = index % 2 === 0 ? '#f8f9fa' : '#ffffff';
    html += `<tr style="background-color: ${rowBackgroundColor};">`;
    // Filter out the 'Department' cell for each row
    for (const header of displayHeaders) {
      let cellValue = row[header] ? row[header] : "";
      let textAlign = "center";

      if (header === "Name") {
        textAlign = "left";
      } else if (header === "KPI (%)") {
        const numericValue = parseFloat(cellValue);
        // Multiply by 100 and format to 2 decimal places.
        // This handles cases where the value is a decimal (e.g., 0.95)
        // or already a percentage string (e.g., "95.00%").
        cellValue = isNaN(numericValue) || numericValue === 0 ? '0%' : `${Math.round(numericValue * 100)}%`;
      }
      if (cellValue === null || cellValue === '' || cellValue === 0) {
        cellValue = 0;
      }

      html += `<td style="border: 1px solid #dee2e6; padding: 8px; text-align: ${textAlign};">${cellValue}</td>`;
    }
    html += `</tr>`;
  });

  html += `
    </tbody>
  </table>`;

  return html;
}



/**
 * Creates a CSV blob from an array of attendance row objects.
 * @param {Object[]} rows The attendance rows for one department.
 * @param {string} filename The desired filename for the CSV.
 * @returns {GoogleAppsScript.Base.Blob} A blob representing the CSV file.
 * @private
 */
function toCsvBlob_(rows, filename) {
  if (rows.length === 0) return null;
  const headers = Object.keys(rows[0]);
  let csvContent = headers.join(",") + "\n";

  for (const row of rows) {
    const rowValues = headers.map(header => {
      let value = row[header] ? row[header].toString() : "";
      // Enclose values with commas or newlines in quotes
      if (value.includes(",") || value.includes("\n")) {
        value = `"${value.replace(/"/g, '""')}"`;
      }
      return value;
    });
    csvContent += rowValues.join(",") + "\n";
  }

  return Utilities.newBlob(csvContent, "text/csv", filename);
}

/**
 * Sends an email with an HTML body and a CSV attachment.
 * @param {string} department The department name.
 * @param {string} recipients The comma-separated email addresses.
 * @param {string} htmlTable The HTML table body.
 * @param {GoogleAppsScript.Base.Blob} csvBlob The CSV file blob.
 * @param {string} dateDisplay The formatted date string for the subject.
 * @private
 */
function sendDepartmentEmail_(department, recipients, htmlTable, csvBlob) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);

  const weekNum = attendanceSheet.getRange("E2").getValue();
  const fromDate = attendanceSheet.getRange("F2").getValue();
  const upToDate = attendanceSheet.getRange("G2").getValue();


  const DEPT_MAP = {
    "1.Admin & HR": "Admin and HR",
    "13.Shwe_Zay": "Shwe Zay",
    "3. M_BI": "M BI",
    "4.Finance": "Finance",
    "6.M_Tech": "M Tech",
    "7.Marathon_Express": "Marathon Express",
    "8.M_Kitchen": "M Kitchen",
    "9.M_Oil": "M Oil",
  };
  department = DEPT_MAP[department];

  const FMT = "dd-MMM-yyyy";
  const fromDate_F = Utilities.formatDate(new Date(fromDate), TIMEZONE, FMT);
  const upToDate_F = Utilities.formatDate(new Date(upToDate), TIMEZONE, FMT);
  // const isoWeekNum = isoWeekLabel(dateDisplay, TIMEZONE);

  const subject = `Attendance (${department}) ${weekNum}`;
  const htmlBody = `
    <p>Dear ${department}'s HOD,</p>
    <p>Please review the following table for your department's attendance KPI.</p>
    <p>If you notice any discrepancies in the report, kindly reach out to hr@marathonmyanmar.com  (Thu Zar Swe).</p>
    <p>Any adjustments will be reflected in the next weekly report.</p>
    <p>Date Form: <strong>${fromDate_F}</strong></p>
    <p>Up To Date: <strong>${upToDate_F}</strong></p>
    ${htmlTable}
    <br>
    <p>Best Regards,<br>
    Marathon BI</p>
  `;

  const options = {
    htmlBody: htmlBody,
    // attachments: [csvBlob],
    cc: CC_EMAILS || undefined, // Add CC if the constant is not empty
    name: "Marathon BI"
  };

  MailApp.sendEmail(recipients, subject, '', options);
}
