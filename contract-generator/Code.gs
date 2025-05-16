function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("租屋合約產生器")
      .addItem("生成合約", "generateContracts")
      .addToUi();
}

function generateContracts() {
  const docTemplateUrl = "https://docs.google.com/document/d/15EadpllRlf5hNutxIhM41rnlzH_k86MZSEYaAtdOKAM/edit?tab=t.0";
  const sheetHeaders = ["訂單編號", "審閱日期", "出租人", "地址", "租約開始", "租約結束", "租金", "合約連結"]; // Manually define the headers

  try {
    // Open the spreadsheet using the default method for bound scripts
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("合約資料");
    if (!sheet) {
      throw Error("找不到'合約資料'分頁！");
    }

    // Get the data range, skipping the header row
    const range = sheet.getDataRange();
    const values = range.getValues();

    // Assuming the first row is the header, get the data rows starting from the second row
    const dataRows = values.slice(1);

    // Find the index of the "合約連結" column
    const linkColumnIndex = sheetHeaders.indexOf("合約連結");
    if (linkColumnIndex === -1) {
      throw new Error("Could not find '合約連結' column in headers.");
    }

    // Extract the template document ID from the URL
    const templateDocId = docTemplateUrl.match(/document\/d\/([a-zA-Z0-9-_]+)/);
    if (!templateDocId || templateDocId.length < 2) {
        throw new Error("Could not extract document ID from template URL.");
    }
    const docId = templateDocId[1];

    // Get the indices for date and money columns
    const reviewDateIndex = sheetHeaders.indexOf("審閱日期");
    const startDateIndex = sheetHeaders.indexOf("租約開始");
    const endDateIndex = sheetHeaders.indexOf("租約結束");
    const rentIndex = sheetHeaders.indexOf("租金");

    // Loop through each data row
    dataRows.forEach((rowData, index) => {
      // Create a copy of the template document
      const templateFile = DriveApp.getFileById(docId);
      const newDocTitle = `Contract - ${rowData[sheetHeaders.indexOf("訂單編號")]}`; // Example title using the order number
      const newFile = templateFile.makeCopy(newDocTitle);
      const newDoc = DocumentApp.openById(newFile.getId());

      // Get the body of the new document
      const body = newDoc.getBody();

      // Replace placeholders with data from the current row
      sheetHeaders.forEach((header, headerIndex) => {
        const placeholder = `{{${header}}}`;
        let replacement = rowData[headerIndex];

        // Apply formatting based on the header
        if (headerIndex === reviewDateIndex || headerIndex === startDateIndex || headerIndex === endDateIndex) {
            // Check if the value is a Date object before formatting
            if (replacement instanceof Date) {
                replacement = displayDateTW(replacement);
            }
        } else if (headerIndex === rentIndex) {
             // Check if the value is a number before formatting
            if (typeof replacement === 'number') {
                replacement = formatMoney(replacement);
            }
        }

        // Replace text while preserving formatting
        // The replaceText method in DocumentApp generally preserves formatting
        // Ensure replacement is a string, in case the original value was not
        body.replaceText(placeholder, String(replacement));
      });

      // Save and close the new document
      newDoc.saveAndClose();

      // Get the URL of the newly generated document
      const newDocUrl = newFile.getUrl();

      // Write the new document URL back to the "合約連結" column in the sheet
      // Add 2 to the index: 1 for skipping the header row, 1 for 1-based indexing
      sheet.getRange(index + 2, linkColumnIndex + 1).setValue(newDocUrl);
    });

    Logger.log("Contracts generated successfully.");

  } catch (error) {
    Logger.log("Error generating contracts: " + error);
  }
}

/**
 * Formats a number as currency with comma thousands separators.
 * @param {number} amount The number to format.
 * @return {string} The formatted currency string.
 */
function formatMoney(amount) {
  if (typeof amount !== 'number') {
    return amount; // Return original value if not a number
  }
  // Use toLocaleString for locale-aware formatting
  return amount.toLocaleString('en-US'); // Using 'en-US' locale for comma separation
}

/**
 * Displays a Date object in 民國/月/日 format.
 * @param {Date} date The date object to format.
 * @return {string} The formatted date string.
 */
function displayDateTW(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return date; // Return original value if not a valid Date
  }
  const year = date.getFullYear();
  const month = date.getMonth() + 1; // getMonth() is 0-indexed
  const day = date.getDate();

  // Calculate Minguo year
  const minguoYear = year - 1911;

  // Pad month and day with leading zeros if necessary
  const formattedMonth = month < 10 ? '0' + month : month;
  const formattedDay = day < 10 ? '0' + day : day;

  return `民國${minguoYear}年${formattedMonth}月${formattedDay}日`;
}

