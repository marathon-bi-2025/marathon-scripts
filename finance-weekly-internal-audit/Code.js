/**
 * Flattens the source data where 15 audit columns repeat 20 times per row.
 * Creates 20 new rows (19 columns each) in the target sheet for every source row.
 */
function flattenTransactionsAudit() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = sheet.getSheetByName("source_02"); 
  const targetSheetName = "flattern_audit";
  const targetSheet = sheet.getSheetByName(targetSheetName) || sheet.insertSheet(targetSheetName);

  const lastTimestampCell = "Z1";
  const lastTimestampValue = targetSheet.getRange(lastTimestampCell).getValue();
  // If the sheet is empty, lastDate is Jan 1, 1970 00:00:00 UTC (0 milliseconds).
  const lastDate = lastTimestampValue ? new Date(lastTimestampValue).getTime() : 0; 

  // --- Configuration ---
  const STATIC_COLS = 4;        // Timestamp, Email, Request, Date-Monday (A, B, C, D)
  const AUDIT_GROUP_SIZE = 17;  // T Ref No. - FIN01 to FIN15 - T Remark
  const NUMBER_OF_GROUPS = 20;  // How many times the 15 FIN columns repeat
  const EXPECTED_TOTAL_COLS = STATIC_COLS + (AUDIT_GROUP_SIZE * NUMBER_OF_GROUPS); // 
  // ---------------------

  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  // Define headers for the 21-column flattened output
  const headers = [
    "Timestamp", "Email Address", "Choose Request", "Date - Monday", "T Ref No.",
    "FIN-1 HOD signature Amend Amount & Description", "FIN-2 Checked by, Approved by, Prepared by", "FIN-3 Xero System bills", "FIN-4 Budget, Approval Request", 
    "FIN-5 Debit Voucher, Business Unit (To tick)", "FIN-6 Bank Balance - Daily Update", "FIN-7 Cash Balance", "FIN-8 Daily Entry update in System",
    "FIN-9 Bank A/C, Bank Information for payment (Correct)", "FIN-10 Avoid of duplicating transfer (Accounting Bill only)", "FIN-11 Daily update Xero",
    "FIN-12 Description Accuracy", "FIN-13 Completion of Source Document", "FIN-14 Advance Form Compliance", "FIN-15 Accuracy of Amount", "T Remark"
  ];

  if (targetSheet.getLastRow() === 0) {
    targetSheet.appendRow(headers);
  }

  const newRows = [];
  let newLastTimestamp = lastTimestampValue;

  // 1. Loop through source data rows (skipping header row i=0)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Validate row length
    if (row.length < EXPECTED_TOTAL_COLS) continue;

    const timestampValue = row[0];
    if (!timestampValue) continue; // Skip if timestamp is empty
    
    const currentDate = new Date(timestampValue).getTime();

    // Check if the row has already been processed: current must be strictly GREATER than last.
    if (currentDate <= lastDate) continue;

    // Extract the 4 static columns once
    const staticData = row.slice(0, STATIC_COLS);

    // 2. Inner loop: Iterate through the 20 repeating groups
    for (let group = 0; group < NUMBER_OF_GROUPS; group++) {
      
      const startIndex = STATIC_COLS + (group * AUDIT_GROUP_SIZE);
      const endIndex = startIndex + AUDIT_GROUP_SIZE;

      // Extract the 15 audit columns for this group
      const auditGroupData = row.slice(startIndex, endIndex);

      // Skip this group if all 15 audit cells are empty (optional, but good practice)
      if (auditGroupData.every(cell => cell === "" || cell === null)) {
          continue; 
      }
      
      // Combine static data (4 columns) + audit group data (15 columns)
      const newRow = [...staticData, ...auditGroupData];
      newRows.push(newRow);
    }
    
    // 3. Update the latest timestamp processed, only if it's new
    if (currentDate > new Date(newLastTimestamp).getTime() || newLastTimestamp === "") {
        newLastTimestamp = timestampValue;
    }
  }

  // 4. Batch write all new data to the target sheet
  if (newRows.length > 0) {
    const startRow = targetSheet.getLastRow() + 1;
    const numRows = newRows.length;
    const numCols = newRows[0].length; // Should be 19
    
    targetSheet.getRange(startRow, 1, numRows, numCols).setValues(newRows);
  }

  // 5. Store the last processed timestamp
  if (newLastTimestamp !== lastTimestampValue) {
      targetSheet.getRange(lastTimestampCell).setValue(newLastTimestamp);
  }
}

/**
 * Global configuration variables
 */
const SHEET_NAME = 'last_audit'; // CHANGE THIS to the actual name of your sheet
const DATA_RANGE = 'A3:U19';          // The range of data to include in the email table


function sendFormattedEmail() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  const WEEK_CELL = sheet.getRange('C1').getValue();
  const timestampCell = sheet.getRange("A1").getValue(); // change cell if needed
  const EMAIL_TO = sheet.getRange('B20').getValue();
  const EMAIL_CC = sheet.getRange('B21').getValue();
  const TREF01 = sheet.getRange('B2').getValue();
  const TREF02 = sheet.getRange('C2').getValue();
  const TREF03 = sheet.getRange('D2').getValue();
  const TREF04 = sheet.getRange('E2').getValue();
  const TREF05 = sheet.getRange('F2').getValue();
  const TREF06 = sheet.getRange('G2').getValue();
  const TREF07 = sheet.getRange('H2').getValue();
  const TREF08 = sheet.getRange('I2').getValue();
  const TREF09 = sheet.getRange('J2').getValue();
  const TREF10 = sheet.getRange('K2').getValue();
  const TREF11 = sheet.getRange('L2').getValue();
  const TREF12 = sheet.getRange('M2').getValue();
  const TREF13 = sheet.getRange('N2').getValue();
  const TREF14 = sheet.getRange('O2').getValue();
  const TREF15 = sheet.getRange('P2').getValue();
  const TREF16 = sheet.getRange('Q2').getValue();
  const TREF17 = sheet.getRange('R2').getValue();
  const TREF18 = sheet.getRange('S2').getValue();
  const TREF19 = sheet.getRange('T2').getValue();
  const TREF20 = sheet.getRange('U2').getValue();

  if (!sheet) {
    Logger.log(`Sheet named '${SHEET_NAME}' not found.`);
    return;
  }

  // Validate timestamp
  if (!timestampCell) {
    Logger.log("No timestamp found, email not sent.");
    return;
  }

  const currentTimestamp = (timestampCell instanceof Date)
    ? Utilities.formatDate(timestampCell, Session.getScriptTimeZone(), "MM-dd-yyyy HH:mm:ss")
    : timestampCell.toString();

  const scriptProps = PropertiesService.getScriptProperties();
  const lastTimestamp = scriptProps.getProperty("lastSentTicketTimestamp");

  // Fetch data from the specified range (e.g., A1:D)
  const data = sheet.getRange(DATA_RANGE).getValues();
  const htmlTable = convertDataToHtmlTable(data);

  const subject = `Weekly Internal Audit for ${WEEK_CELL}`;
  const body = `Dear Management, <br><br>
    Please find attached the weekly internal audit results table for your review.<br><br>
    This report summarizes the key findings from our internal audit activities for the week, ${WEEK_CELL}.<br><br>
    ${htmlTable}
    <br><br>
    <table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 10pt;">
    <tbody>
        <tr><td style="padding-bottom: 5px;"></td></tr>
        <tr>
            <td colspan="2" style="padding-bottom: 5px;">
                <span style="font-weight: bold; font-size: 14pt; text-decoration: underline;">Transaction Ref No.</span>
            </td>
            
            <td style="padding-right: 40px;"></td> 
            
            <td colspan="2" style="padding-bottom: 5px;">
                <span style="font-weight: bold; font-size: 14pt; text-decoration: underline;">Audit Criteria Definition</span>
            </td>
        </tr>
        <tr>
            <td colspan="5" style="padding: 0; line-height: 5px; height: 5px; background-color: #FFFFFF;">
                &nbsp;
            </td>
        </tr
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR1</td>
            <td style="padding: 2px 0;"> ${TREF01}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td> <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-1</td>
            <td style="padding: 2px 0;"> HOD signature Amend Amount & Description</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR2</td>
            <td style="padding: 2px 0;"> ${TREF02}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-2</td>
            <td style="padding: 2px 0;"> Checked by, Approved by, Prepared by</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR3</td>
            <td style="padding: 2px 0;"> ${TREF03}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-3</td>
            <td style="padding: 2px 0;"> Xero System bills</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR4</td>
            <td style="padding: 2px 0;"> ${TREF04}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-4</td>
            <td style="padding: 2px 0;"> Budget, Approval Request</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR5</td>
            <td style="padding: 2px 0;"> ${TREF05}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-5</td>
            <td style="padding: 2px 0;"> Debit Voucher, Business Unit (To tick)</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR6</td>
            <td style="padding: 2px 0;"> ${TREF06}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-6</td>
            <td style="padding: 2px 0;"> Bank Balance - Daily Update</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR7</td>
            <td style="padding: 2px 0;"> ${TREF07}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-7</td>
            <td style="padding: 2px 0;"> Cash Balance</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR8</td>
            <td style="padding: 2px 0;"> ${TREF08}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-8</td>
            <td style="padding: 2px 0;"> Daily Entry update in System</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR9</td>
            <td style="padding: 2px 0;"> ${TREF09}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-9</td>
            <td style="padding: 2px 0;"> Bank A/C, Bank Information for payment (Correct)</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR10</td>
            <td style="padding: 2px 0;"> ${TREF10}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-10</td>
            <td style="padding: 2px 0;"> Avoid of duplicating transfer (Accounting Bill only)</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR11</td>
            <td style="padding: 2px 0;"> ${TREF11}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-11</td>
            <td style="padding: 2px 0;"> Daily update Xero</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR12</td>
            <td style="padding: 2px 0;"> ${TREF12}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-12</td>
            <td style="padding: 2px 0;"> Description Accuracy</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR13</td>
            <td style="padding: 2px 0;"> ${TREF13}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-13</td>
            <td style="padding: 2px 0;"> Completion of Source Document</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR14</td>
            <td style="padding: 2px 0;"> ${TREF14}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-14</td>
            <td style="padding: 2px 0;"> Advance Form Compliance</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR15</td>
            <td style="padding: 2px 0;"> ${TREF15}</td>
            <td style="padding: 2px 40px 2px 0;">&nbsp;</td>
            <td style="padding: 2px 10px 2px 0; text-align: left;">FIN-15</td>
            <td style="padding: 2px 0;"> Accuracy of Amount</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR16</td>
            <td style="padding: 2px 0;"> ${TREF16}</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR17</td>
            <td style="padding: 2px 0;"> ${TREF17}</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR18</td>
            <td style="padding: 2px 0;"> ${TREF18}</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR19</td>
            <td style="padding: 2px 0;"> ${TREF19}</td>
        </tr>
        <tr>
            <td style="padding: 2px 10px 2px 0; text-align: left;">TR20</td>
            <td style="padding: 2px 0;"> ${TREF20}</td>
        </tr>
        <tr><td style="padding-bottom: 20px;"></td></tr>
      </tbody>
    </table>

    <p>Best regards,<br>Marathon BI</p>
  `;

  // --- Send only if new ---
  if (currentTimestamp && currentTimestamp !== lastTimestamp) {

    try {
      // Send the email
      MailApp.sendEmail({
        to: EMAIL_TO,
        cc: EMAIL_CC,
        subject: subject,
        htmlBody: body,
        name: 'Marathon BI'
      });
      Logger.log('Email sent successfully.');
      // Save timestamp
      scriptProps.setProperty("lastSentTicketTimestamp", currentTimestamp);
    } catch (e) {
      Logger.log("!Error sending weekly audit email: " + e.toString());
    }
  } else {
    Logger.log("No new timestamp or timestamp already processed!");
  }
}


function convertDataToHtmlTable(data) {
  let html = '<table style="border-collapse: collapse; width: auto; font-family: Arial, sans-serif;">';

  data.forEach((row, rowIndex) => {
    html += '<tr style="border: 1px solid #ccc;">';
    
    // Use the first row as the header with a dark background
    const isHeader = (rowIndex === 0);
    const cellTag = isHeader ? 'th' : 'td';
    const headerStyle = 'background-color: #333; color: white; font-weight: bold; padding: 8px; text-align: left;';

    row.forEach(cellValue => {
      let cellStyle = 'padding: 8px; border: 1px solid #ccc;';
      const cellText = String(cellValue).trim().toUpperCase();

      if (!isHeader) {
        if (cellText === 'PASS') {
          cellStyle += 'background-color: #D9EAD3; color: #10630D; font-weight: bold;'; // Light Green
        } else if (cellText === 'FAIL') {
          cellStyle += 'background-color: #F4CCCC; color: #CC0000; font-weight: bold;'; // Light Red
        } else {
          cellStyle += 'font-weight: bold; text-align: center';
        }
      }

      html += `<${cellTag} style="${isHeader ? headerStyle : cellStyle}">${cellValue}</${cellTag}>`;
    });
    
    html += '</tr>';
  });

  html += '</table>';
  return html;
}