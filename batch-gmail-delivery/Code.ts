/**
 * This script performs a mail merge, sending emails from Gmail using a
 * Google Doc as a template and a Google Sheet as the data source.
 *
 * Configuration is managed in a separate sheet named "Settings".
 */

// --- SETTINGS KEYS ---
// These are the names of the settings the script looks for in the "Settings" sheet.
const DOC_ID_KEY = 'Google Doc Template ID';
const SHEET_NAME_KEY = 'Data Sheet Name';
const EMAIL_COL_KEY = 'Email Column Name';
// --- END SETTINGS KEYS ---


/**
 * Creates a custom menu in the Google Sheet UI when the file is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('批量寄送Gmail')
    .addItem('寄送Gmail', 'sendTemplatedEmails')
    .addSeparator()
    .addItem('建立/重置設定表單', 'createSettingsSheet')
    .addToUi();
}

/**
 * Creates or resets the "Settings" sheet with default values.
 */
function createSettingsSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Settings');

  if (!sheet) {
    sheet = ss.insertSheet('Settings');
  } else {
    // Optional: Ask for confirmation before clearing
    const response = ui.alert('重置設定？', '您確定要重置「設定」工作表嗎？這將會清除所有現有值。', ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) {
      return;
    }
    sheet.clear();
  }

  // Set headers and default values
  const defaultSettings = [
    ['設定項目', '數值'],
    [DOC_ID_KEY, 'YOUR_DOC_ID_GOES_HERE'],
    [SHEET_NAME_KEY, '客戶資料'],
    [EMAIL_COL_KEY, '客戶信箱']
  ];

  sheet.getRange(1, 1, defaultSettings.length, 2).setValues(defaultSettings);

  // Formatting
  sheet.setFrozenRows(1);
  sheet.getRange("A1:B1").setFontWeight("bold");
  sheet.autoResizeColumn(1);
  sheet.getRange("B:B").setNumberFormat('@'); // Set values as plain text

  // Activate the cell the user needs to edit first
  sheet.getRange("B2").activate();
  ui.alert('設定工作表已準備就緒', '「設定」工作表已建立。請在儲存格 B2 中填入您的 Google 文件 ID。', ui.ButtonSet.OK);
}

/**
 * Reads the configuration from the "Settings" sheet and returns it as an object.
 * @return {object | null} An object with setting keys and values, or null if setup fails.
 */
function getSettings() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');

  if (!sheet) {
    ui.alert(
      '找不到設定工作表',
      '找不到「設定」工作表。請執行「批量寄送Gmail > 建立/重置設定表單」來建立它。',
      ui.ButtonSet.OK
    );
    return null;
  }

  const data = sheet.getDataRange().getValues();
  const settings = {};

  // Start from row 1 (skip header)
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];   // Column A
    const value = data[i][1]; // Column B
    if (key) {
      settings[key.trim()] = value.trim();
    }
  }

  // Validate that all required settings are present
  const requiredKeys = [DOC_ID_KEY, SHEET_NAME_KEY, EMAIL_COL_KEY];
  const missingKeys = [];
  for (const key of requiredKeys) {
    if (!settings[key]) {
      missingKeys.push(key);
    }
  }

  if (missingKeys.length > 0) {
    ui.alert(
      '設定不完整',
      `您的「設定」工作表缺少以下項目的數值：\n\n${missingKeys.join('\n')}\n\n請更新「設定」工作表。`,
      ui.ButtonSet.OK
    );
    return null;
  }

  return settings;
}


/**
 * Main function to process the sheet and send emails.
 * Triggered by the "Send Billing Emails" menu item.
 */
function sendTemplatedEmails() {
  const ui = SpreadsheetApp.getUi();

  // 1. Get configuration from the "Settings" sheet
  const settings = getSettings();
  if (!settings) {
    return; // getSettings() already showed an alert
  }

  // Assign settings to variables for easier use
  const docTemplateId = settings[DOC_ID_KEY];
  const sheetName = settings[SHEET_NAME_KEY];
  const emailColumnName = settings[EMAIL_COL_KEY];
  const sentDateColumnName = '已寄送？';
  const statusColumnName = '狀態';

  // 2. Check if the Doc ID has been filled in
  if (docTemplateId === 'YOUR_DOC_ID_GOES_HERE' || !docTemplateId) {
    ui.alert(
      '腳本尚未設定',
      '請前往「設定」工作表，並在儲存格 B2 中輸入您的 Google 文件範本 ID。',
      ui.ButtonSet.OK
    );
    return;
  }

  // 3. Get the Google Doc template
  let doc, templateSubject, templateBody;
  try {
    doc = DocumentApp.openById(docTemplateId);
    const fullTemplateText = doc.getBody().getText();

    // The first line of the Doc is the Subject, the rest is the Body
    const firstNewline = fullTemplateText.indexOf('\n');
    if (firstNewline === -1) {
      templateSubject = fullTemplateText;
      templateBody = ''; // No body, just subject
    } else {
      templateSubject = fullTemplateText.substring(0, firstNewline).trim();
      templateBody = fullTemplateText.substring(firstNewline + 1).trim();
    }

    if (!templateSubject || !templateBody) {
      ui.alert(
        '範本錯誤',
        '您的 Google 文件範本似乎是空的，或缺少主旨（第一行）和內文（文件的其餘部分）。',
        ui.ButtonSet.OK
      );
      return;
    }

  } catch (e) {
    Logger.log(e);
    ui.alert(
      '範本錯誤',
      `無法開啟 Google 文件範本。請檢查「設定」工作表中的 ID，並確認您有權限檢視該文件。錯誤：${e.message}`,
      ui.ButtonSet.OK
    );
    return;
  }

  // 4. Get the Google Sheet data
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    ui.alert(
      '工作表錯誤',
      `找不到名稱為「${sheetName}」的工作表。請檢查「設定」工作表中的「資料工作表名稱」數值。`,
      ui.ButtonSet.OK
    );
    return;
  }

  const dataRange = sheet.getDataRange();
  // Get all values, including headers
  const values = dataRange.getDisplayValues();

  // The first row (index 0) contains the headers
  const headers = values[0];

  // 5. Find the column indices by name (this is the key to your request)
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[header.trim()] = index;
  });

  const emailColIndex = headerMap[emailColumnName];
  const sentColIndex = headerMap[sentDateColumnName];
  const statusColIndex = headerMap[statusColumnName];

  // 6. Validate that we found the required columns
  if (emailColIndex === undefined) {
    ui.alert(
      '找不到欄位',
      `在「${sheetName}」工作表中找不到必要的欄位「${emailColumnName}」。請檢查您的「設定」。`,
      ui.ButtonSet.OK
    );
    return;
  }

  if (sentColIndex === undefined) {
    ui.alert(
      '找不到欄位',
      `在「${sheetName}」工作表中找不到必要的欄位「已寄送？」。`,
      ui.ButtonSet.OK
    );
    return;
  }

  if (statusColIndex === undefined) {
    ui.alert(
      '找不到欄位',
      `在「${sheetName}」工作表中找不到必要的欄位「狀態」。`,
      ui.ButtonSet.OK
    );
    return;
  }

  // 7. Process each row
  let emailsSentCount = 0;
  // Start from row 1 (the first data row), skip headers (row 0)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sentStatus = row[sentColIndex];
    const recipientEmail = row[emailColIndex];
    const status = row[statusColIndex];

    // Only send if checkbox is unchecked, status is "Ready", and email is not empty
    // Checkbox values come as strings "TRUE" or "FALSE" from getDisplayValues()
    if (sentStatus !== 'TRUE' && status === 'Ready' && recipientEmail) {
      // Build the data object for this row
      const rowData = {};
      headers.forEach((header, index) => {
        // Use the header name as the key
        rowData[header.trim()] = row[index];
      });

      // Fill the template with this row's data
      const subject = fillTemplate(templateSubject, rowData);
      const body = fillTemplate(templateBody, rowData);

      try {
        // 8. Send the email
        GmailApp.sendEmail(recipientEmail, subject, body);

        // 9. Update the checkbox column to checked (true)
        // We add 1 to 'i' because row index is 0-based, but sheet rows are 1-based
        // We add 1 to 'sentColIndex' because col index is 0-based, but sheet cols are 1-based
        sheet.getRange(i + 1, sentColIndex + 1).setValue(true);
        emailsSentCount++;

      } catch (e) {
        Logger.log(`Failed to send email to ${recipientEmail}: ${e.message}`);
        // Optional: Write an error message to a 'Status' column if you have one
        // For now, we just log it and don't update the 'Sent Date'
      }
    }
  }

  // 10. Show a summary
  if (emailsSentCount > 0) {
    ui.alert(
      '成功',
      `已成功寄送 ${emailsSentCount} 封郵件。`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      '完成',
      '沒有新的郵件需要寄送。（所有列都已勾選、狀態不是「Ready」或缺少「客戶信箱」）。',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Helper function to replace {{variable}} placeholders in a template string.
 * @param {string} template The template string (e.g., "Hello {{Customer Name}}")
 * @param {object} data A data object (e.g., { "Customer Name": "ABC" })
 * @return {string} The template with placeholders replaced.
 */
function fillTemplate(template, data) {
  let result = template;
  // Iterate over each key in our data object
  for (const key in data) {
    // Create a regular expression to match the placeholder globally (all instances)
    // We escape special regex characters in the key, just in case (e.g., "Amount (USD)")
    const placeholder = '{{' + key + '}}';
    const regex = new RegExp(escapeRegExp(placeholder), 'g');
    
    // Replace the placeholder with the corresponding value
    result = result.replace(regex, data[key]);
  }
  return result;
}

/**
 * Helper function to escape special characters for use in a RegExp.
 * @param {string} str The string to escape.
 * @return {string} The escaped string.
 */
function escapeRegExp(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}

