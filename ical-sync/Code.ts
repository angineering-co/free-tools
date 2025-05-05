/**
 * Configuration Section (設定區段)
 * -------------------------------
 * Names of the sheets used by the script. (指令碼使用的試算表工作表名稱。)
 */
const ICAL_CONFIG_SHEET_NAME = "日曆連結"; // Sheet containing the list of iCal URLs (包含 iCal 網址清單的工作表)
const BOOKINGS_SHEET_NAME = "預訂紀錄";      // Sheet where booking data is written (寫入預訂資料的工作表)

/**
 * Script Constants (指令碼常數) - Generally no need to edit (通常無需編輯)
 * ---------------------------------------------------------------------
 */
// Columns in the '日曆連結' sheet (1-based index) (「日曆連結」工作表中的欄位，索引從 1 開始)
const ICAL_PROP_NAME_COL = 1; // 房源名稱
const ICAL_URL_COL = 2;       // iCal 網址
const ICAL_ENABLED_COL = 3;   // 啟用

// Columns in the '預訂紀錄' sheet (1-based index) (「預訂紀錄」工作表中的欄位，索引從 1 開始)
const BOOKING_UID_COL = 1;        // A欄: 預訂編號 (UID)
const BOOKING_PROP_NAME_COL = 2;  // B欄: 房源名稱
const BOOKING_STATUS_COL = 3;     // C欄: 狀態
const BOOKING_GUEST_INFO_COL = 4; // D欄: 房客資訊 (摘要)
const BOOKING_CHECKIN_COL = 5;    // E欄: 入住日期
const BOOKING_CHECKOUT_COL = 6;   // F欄: 退房日期
const BOOKING_NIGHTS_COL = 7;     // G欄: 晚數
const BOOKING_LAST_UPDATED_COL = 8;// H欄: 最後更新

// Header row for the '預訂紀錄' sheet (「預訂紀錄」工作表的標頭列)
const BOOKINGS_HEADER_ROW = [
  "預訂編號 (UID)",
  "房源名稱",
  "狀態",
  "房客資訊 (摘要)",
  "入住日期",
  "退房日期",
  "晚數",
  "最後更新"
];
// Value in '啟用' column to process a URL (case-insensitive) (「啟用」欄中表示要處理的值，不區分大小寫)
const ENABLED_VALUE = "是";

function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("AirBnb iCal Sync")
        .addItem("同步預訂紀錄", "syncAllIcalLinks")
        .addToUi();
}

/**
 * Main function to synchronize ALL enabled iCal links from the config sheet
 * to the Google Sheet. Run this function manually or set up a time-driven trigger.
 * (主函數：從設定工作表同步所有已啟用的 iCal 連結到 Google 試算表。手動執行此函數或設定時間驅動觸發器。)
 */
function syncAllIcalLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(ICAL_CONFIG_SHEET_NAME);
  const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);

  if (!configSheet) {
    const msg = `錯誤：找不到設定工作表 "${ICAL_CONFIG_SHEET_NAME}"。請建立該工作表並包含欄位：房源名稱, iCal 網址, 啟用。`;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
    return;
  }
  if (!bookingsSheet) {
     const msg = `錯誤：找不到目標工作表 "${BOOKINGS_SHEET_NAME}"。請建立該工作表或檢查 BOOKINGS_SHEET_NAME 常數。`;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
    return;
  }

  const configData = configSheet.getDataRange().getValues();
  const existingBookingsData = getSheetData(bookingsSheet); // { uid: { data: [row values], rowIndex: number } }
  const scriptTimeZone = Session.getScriptTimeZone();
  const now = new Date();
  const allRowsToAdd = [];
  const allRowsToUpdate = []; // { rowIndex: number, data: [row values] }
  const allProcessedUIDs = new Set();

  // Iterate through configured iCal URLs (skip header row) (遍歷設定的 iCal 網址，跳過標頭列)
  for (let i = 1; i < configData.length; i++) {
    const row = configData[i];
    const propertyName = row[ICAL_PROP_NAME_COL - 1];
    const icalUrl = row[ICAL_URL_COL - 1];
    const isEnabled = String(row[ICAL_ENABLED_COL - 1]).trim(); // Don't convert to lower case here, compare directly

    // Basic validation for the row (該列的基本驗證)
    if (isEnabled !== ENABLED_VALUE) { // Direct comparison with '是'
      Logger.log(`跳過 ${ICAL_CONFIG_SHEET_NAME} 中的第 ${i + 1} 列：未啟用。`);
      continue;
    }
     if (!propertyName || typeof propertyName !== 'string' || propertyName.trim() === '') {
        Logger.log(`跳過 ${ICAL_CONFIG_SHEET_NAME} 中的第 ${i + 1} 列：缺少或無效的房源名稱。`);
        continue;
    }
    if (!icalUrl || typeof icalUrl !== 'string' || !icalUrl.toLowerCase().startsWith('http')) {
      Logger.log(`跳過 ${ICAL_CONFIG_SHEET_NAME} 中的第 ${i + 1} 列 (${propertyName})：缺少或無效的 iCal 網址。`);
      continue;
    }

    Logger.log(`正在處理房源： "${propertyName}"，來源網址： ${icalUrl}`);

    try {
      const icalText = fetchIcalData(icalUrl);
      if (!icalText) {
        Logger.log(`讀取 ${propertyName} 的資料失敗。跳過此來源。`);
        continue; // Skip to the next URL if fetch fails
      }

      const events = parseIcalData(icalText);
      Logger.log(`為 "${propertyName}" 解析了 ${events.length} 個事件。`);

      // Process events for *this specific* iCal feed (處理此特定 iCal Feed 的事件)
      events.forEach(event => {
        if (!event.uid || !event.dtstart || !event.dtend) {
          Logger.log(`為 "${propertyName}" 跳過事件，因缺少 UID、DTSTART 或 DTEND： ${JSON.stringify(event)}`);
          return;
        }

        allProcessedUIDs.add(event.uid);

        const checkInDate = event.dtstart;
        const checkOutDate = event.dtend;
        const nights = calculateNights(checkInDate, checkOutDate);
        const status = translateStatus(event.status || 'CONFIRMED'); // Translate status here

        const currentRowData = [
          event.uid,
          propertyName,
          status, // Use translated status
          event.summary || '',
          Utilities.formatDate(checkInDate, scriptTimeZone, 'yyyy-MM-dd'),
          Utilities.formatDate(checkOutDate, scriptTimeZone, 'yyyy-MM-dd'),
          nights,
          Utilities.formatDate(now, scriptTimeZone, 'yyyy-MM-dd HH:mm:ss')
        ];

        if (existingBookingsData[event.uid]) {
          // Existing event: Check if update needed (已存在事件：檢查是否需要更新)
          const existingRow = existingBookingsData[event.uid];
          const existingRowIndex = existingRow.rowIndex;
          const oldData = existingRow.data;
          const oldStatus = oldData[BOOKING_STATUS_COL - 1];

          // Only update if essential data changed OR if property name differs
          if (oldStatus !== currentRowData[BOOKING_STATUS_COL - 1] || // Compare translated status
              oldData[BOOKING_CHECKIN_COL - 1] !== currentRowData[BOOKING_CHECKIN_COL - 1] ||
              oldData[BOOKING_CHECKOUT_COL - 1] !== currentRowData[BOOKING_CHECKOUT_COL - 1] ||
              oldData[BOOKING_GUEST_INFO_COL - 1] !== currentRowData[BOOKING_GUEST_INFO_COL - 1] ||
              oldData[BOOKING_PROP_NAME_COL - 1] !== currentRowData[BOOKING_PROP_NAME_COL - 1])
          {
             const existingUpdateIndex = allRowsToUpdate.findIndex(update => update.rowIndex === existingRowIndex);
             if (existingUpdateIndex > -1) {
                 allRowsToUpdate[existingUpdateIndex].data = currentRowData;
                 Logger.log(`以來自 ${propertyName} 的資料覆寫第 ${existingRowIndex} 列 (UID: ${event.uid}) 的更新指令。`);
             } else {
                 allRowsToUpdate.push({ rowIndex: existingRowIndex, data: currentRowData });
                 Logger.log(`從 ${propertyName} 標記第 ${existingRowIndex} 列 (UID: ${event.uid}) 進行更新。`);
             }
          } else {
             // No changes detected, but update the "Last Updated" timestamp if it's old (未偵測到變更，但如果「最後更新」時間戳記過舊則更新)
             const lastUpdateDateInSheet = new Date(oldData[BOOKING_LAST_UPDATED_COL -1]);
             if (isNaN(lastUpdateDateInSheet.getTime()) || (now.getTime() - lastUpdateDateInSheet.getTime() > 60 * 60 * 1000 )) { // Update if older than 1 hour or invalid date
                 currentRowData[BOOKING_LAST_UPDATED_COL - 1] = Utilities.formatDate(now, scriptTimeZone, 'yyyy-MM-dd HH:mm:ss');
                 const existingUpdateIndex = allRowsToUpdate.findIndex(update => update.rowIndex === existingRowIndex);
                  if (existingUpdateIndex > -1) {
                       allRowsToUpdate[existingUpdateIndex].data[BOOKING_LAST_UPDATED_COL - 1] = currentRowData[BOOKING_LAST_UPDATED_COL - 1];
                   } else {
                       allRowsToUpdate.push({ rowIndex: existingRowIndex, data: currentRowData });
                       Logger.log(`更新第 ${existingRowIndex} 列 (UID: ${event.uid}) 的「最後更新」時間戳記。`);
                   }
             }
          }
        } else {
          // New event (新事件)
           const existingAddIndex = allRowsToAdd.findIndex(newRow => newRow[BOOKING_UID_COL - 1] === event.uid);
           if (existingAddIndex === -1) {
               allRowsToAdd.push(currentRowData);
               Logger.log(`從 ${propertyName} 標記新列以供添加 (UID: ${event.uid})。`);
           } else {
               Logger.log(`來自 ${propertyName} 的 UID ${event.uid} 已在此次執行中標記為新增。跳過重複新增。`);
           }
        }
      }); // End processing events for one feed

    } catch (error) {
       Logger.log(`處理 "${propertyName}" (網址: ${icalUrl}) 的 Feed 時發生錯誤： ${error.message}\n${error.stack}`);
       // Continue to the next feed
    }
  } // End loop through config sheet rows

  // --- Perform aggregated Sheet Updates (執行彙總的工作表更新) ---
  Logger.log("開始工作表更新程序...");

  // 1. Add new rows (新增列)
  if (allRowsToAdd.length > 0) {
    const startRow = bookingsSheet.getLastRow() + 1;
    bookingsSheet.getRange(startRow, 1, allRowsToAdd.length, BOOKINGS_HEADER_ROW.length).setValues(allRowsToAdd);
    Logger.log(`已新增 ${allRowsToAdd.length} 個新列到 ${BOOKINGS_SHEET_NAME}。`);
  } else {
     Logger.log("沒有新列需要新增。");
  }

  // 2. Update existing rows (更新現有列)
  if (allRowsToUpdate.length > 0) {
    allRowsToUpdate.forEach(update => {
      if (update.rowIndex > 0 && update.rowIndex <= bookingsSheet.getMaxRows()) {
         bookingsSheet.getRange(update.rowIndex, 1, 1, BOOKINGS_HEADER_ROW.length).setValues([update.data]);
      } else {
         Logger.log(`跳過無效列索引的更新： ${update.rowIndex} (UID: ${update.data ? update.data[BOOKING_UID_COL - 1] : 'N/A'})`);
      }
    });
    Logger.log(`已更新 ${allRowsToUpdate.length} 個列於 ${BOOKINGS_SHEET_NAME}。`);
  } else {
      Logger.log("沒有列需要更新。");
  }

  // 3. Handle potentially removed/cancelled events (處理可能已移除/取消的事件)
  let missingCount = 0;
  const sheetUIDs = Object.keys(existingBookingsData);
  sheetUIDs.forEach(uid => {
    if (!allProcessedUIDs.has(uid)) {
       const missingRowIndex = existingBookingsData[uid].rowIndex;
       const currentStatus = bookingsSheet.getRange(missingRowIndex, BOOKING_STATUS_COL).getValue();
       const cancelledStatusValue = translateStatus('CANCELLED'); // "已取消"
       const possiblyCancelledStatusValue = "可能已取消?";

       // Only mark as 'Possibly Cancelled?' if not already marked as cancelled or possibly cancelled.
       if (currentStatus !== cancelledStatusValue && currentStatus !== possiblyCancelledStatusValue) {
           bookingsSheet.getRange(missingRowIndex, BOOKING_STATUS_COL).setValue(possiblyCancelledStatusValue);
           bookingsSheet.getRange(missingRowIndex, BOOKING_LAST_UPDATED_COL).setValue(Utilities.formatDate(now, scriptTimeZone, 'yyyy-MM-dd HH:mm:ss'));
           Logger.log(`事件 UID ${uid} (列 ${missingRowIndex}) 存在於工作表中，但在任何已處理的 iCal Feed 中找不到。已標記為 '${possiblyCancelledStatusValue}'。`);
           missingCount++;
       } else {
            Logger.log(`事件 UID ${uid} (列 ${missingRowIndex}) 存在於工作表中，但在任何已處理的 iCal Feed 中找不到。狀態已是 '${currentStatus}'。未做更改。`);
       }
    }
  });
  if(missingCount > 0) {
      Logger.log(`標記了 ${missingCount} 個工作表中的事件為 '${translateStatus('Possibly Cancelled?')}'，因為在任何啟用的 iCal Feed 中都找不到它們。`);
  }


  // Ensure header row exists and is correct (確保標頭列存在且正確)
  ensureHeaderRow(bookingsSheet, BOOKINGS_HEADER_ROW, BOOKING_UID_COL);

  // Optional: Sort sheet by Check-out date (可選：依退房日期排序工作表)
  sortSheet(bookingsSheet, BOOKING_CHECKOUT_COL); // Sort by Column F (依 F 欄排序)

  Logger.log("所有已啟用 Feed 的同步完成。");
  SpreadsheetApp.getActiveSpreadsheet().toast("iCal 同步完成!"); // User feedback (使用者提示)
}


// ================================================
// Helper Functions (輔助函數) - Most are unchanged (多數未變)
// ================================================

/**
 * Translates standard iCal status values to Traditional Chinese.
 * @param {string} status The status string (e.g., CONFIRMED, TENTATIVE, CANCELLED). Case-insensitive.
 * @return {string} The translated status or the original if no translation exists.
 */
function translateStatus(status) {
    if (!status) return ''; // Handle null or undefined status
    const upperStatus = status.toUpperCase();
    switch (upperStatus) {
        case 'CONFIRMED': return '已確認';
        case 'TENTATIVE': return '暫定';
        case 'CANCELLED': return '已取消';
        case 'POSSIBLY CANCELLED?': return '可能已取消?'; // Also translate our custom status
        default: return status; // Return original if unknown
    }
}


/**
 * Fetches iCal data from the specified URL.
 * (從指定網址讀取 iCal 資料。)
 * @param {string} url The iCal URL.
 * @return {string|null} The iCal data as text, or null on error.
 */
function fetchIcalData(url) {
  try {
    const cacheBusterUrl = url + (url.includes('?') ? '&' : '?') + 'cachebust=' + Date.now();
    const response = UrlFetchApp.fetch(cacheBusterUrl, {
        validateHttpsCertificates: true,
        muteHttpExceptions: true,
        headers: { 'Cache-Control': 'no-cache', 'Pragma': 'no-cache' }
     });
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      return response.getContentText();
    } else {
      Logger.log(`讀取 iCal 網址 ${url} 時發生錯誤。回應碼： ${responseCode}。 回應內容： ${response.getContentText().substring(0, 500)}`);
      return null;
    }
  } catch (e) {
    Logger.log(`讀取 iCal 網址 ${url} 時發生例外狀況： ${e}`);
    return null;
  }
}

/**
 * Parses iCal text data into an array of event objects.
 * (將 iCal 文字資料解析為事件物件陣列。)
 * @param {string} icalText The raw iCal data.
 * @return {Array<Object>} An array of event objects.
 */
function parseIcalData(icalText) {
  const events = [];
  let currentEvent = null;
  const lines = icalText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const unfoldedLines = [];
  lines.forEach(line => {
      if (line.startsWith(' ') || line.startsWith('\t')) {
          if (unfoldedLines.length > 0) unfoldedLines[unfoldedLines.length - 1] += line.substring(1);
      } else if (line.trim() !== "") {
          unfoldedLines.push(line);
      }
  });

  unfoldedLines.forEach(line => {
    line = line.trim();
    if (line === 'BEGIN:VEVENT') {
      currentEvent = {};
    } else if (line === 'END:VEVENT') {
      if (currentEvent && currentEvent.uid && currentEvent.dtstart && currentEvent.dtend) {
         events.push(currentEvent);
      } else {
         Logger.log("跳過不完整的事件 (缺少 UID, DTSTART, 或 DTEND): " + JSON.stringify(currentEvent));
      }
      currentEvent = null;
    } else if (currentEvent) {
      const colonIndex = line.indexOf(':');
      if (colonIndex > 0) {
        let key = line.substring(0, colonIndex);
        const value = line.substring(colonIndex + 1);
        const params = {};
        const semicolonIndex = key.indexOf(';');
        if (semicolonIndex > 0) {
           const paramsStr = key.substring(semicolonIndex + 1);
           paramsStr.split(';').forEach(p => { const eqIndex = p.indexOf('='); if (eqIndex > 0) params[p.substring(0, eqIndex).toUpperCase()] = p.substring(eqIndex + 1); });
           key = key.substring(0, semicolonIndex);
        }
        key = key.toUpperCase();
        switch (key) {
          case 'UID': currentEvent.uid = value; break;
          case 'SUMMARY': currentEvent.summary = value; break;
          case 'STATUS': currentEvent.status = value; break; // Keep original status here, translate later
          case 'DTSTART': currentEvent.dtstart = parseIcalDate(value, params); break;
          case 'DTEND': currentEvent.dtend = parseIcalDate(value, params); break;
          case 'DESCRIPTION': currentEvent.description = value.replace(/\\n/g, '\n'); break;
        }
      }
    }
  });
  return events;
}

/**
 * Parses an iCal date string into a JavaScript Date object.
 * (將 iCal 日期字串解析為 JavaScript Date 物件。)
 * @param {string} dateString The iCal date string.
 * @param {Object} params Parameters associated with the date field.
 * @return {Date|null} A Date object or null if parsing fails.
 */
function parseIcalDate(dateString, params = {}) {
  try {
    const isValueDate = params['VALUE'] === 'DATE';
    const match = dateString.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})(Z)?)?$/);
    if (match) {
      const year = parseInt(match[1], 10), month = parseInt(match[2], 10) - 1, day = parseInt(match[3], 10);
      if (!isValueDate && match[4]) { // Date-time
        const hour = parseInt(match[4], 10), minute = parseInt(match[5], 10), second = parseInt(match[6], 10);
        return new Date(Date.UTC(year, month, day, hour, minute, second)); // Always use UTC for reliability
      } else { // Date only
        return new Date(Date.UTC(year, month, day, 0, 0, 0)); // Use UTC midnight
      }
    }
    Logger.log(`無法解析日期字串： ${dateString} 參數： ${JSON.stringify(params)}`);
    return null;
  } catch (e) { Logger.log(`解析日期字串 "${dateString}" 時發生錯誤： ${e}`); return null; }
}

/**
 * Reads existing data from the sheet and returns a map keyed by UID.
 * (讀取工作表現有資料並回傳以 UID 為鍵值的地圖。)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @return {Object} A map where keys are UIDs and values are { data: [row values], rowIndex: number }.
 */
function getSheetData(sheet) {
  ensureHeaderRow(sheet, BOOKINGS_HEADER_ROW, BOOKING_UID_COL);
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() <= 1) return {};
  const values = dataRange.getValues();
  const dataMap = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const uid = row[BOOKING_UID_COL - 1];
    if (uid && (typeof uid === 'string' || typeof uid === 'number') && String(uid).trim() !== '') {
      dataMap[String(uid)] = { data: row, rowIndex: i + 1 };
    } else {
      Logger.log(`跳過 ${sheet.getName()} 中的第 ${i + 1} 列，因缺少或無效的 UID： ${row[BOOKING_UID_COL - 1]}`);
    }
  }
  return dataMap;
}

/**
 * Calculates the number of nights between check-in and check-out dates.
 * (計算入住和退房日期之間的天數。)
 * @param {Date} checkInDate The check-in date object.
 * @param {Date} checkOutDate The check-out date object.
 * @return {number} The number of nights.
 */
function calculateNights(checkInDate, checkOutDate) {
  if (!checkInDate || !checkOutDate || isNaN(checkInDate.getTime()) || isNaN(checkOutDate.getTime())) {
    Logger.log(`用於計算晚數的日期無效： IN=${checkInDate}, OUT=${checkOutDate}`);
    return 0;
  }
  const diffTime = checkOutDate.getTime() - checkInDate.getTime();
  const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
  return diffDays > 0 ? diffDays : 0;
}

/**
 * Checks/ensures the header row exists and matches expected headers.
 * (檢查/確保標頭列存在且符合預期標頭。)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {Array<String>} expectedHeaders Expected header strings.
 * @param {number} firstColIndex 1-based index of the first column.
 */
function ensureHeaderRow(sheet, expectedHeaders, firstColIndex = 1) {
   const headerNumCols = expectedHeaders.length;
   if (sheet.getLastRow() === 0) {
      sheet.insertRowBefore(1);
      const headerRange = sheet.getRange(1, firstColIndex, 1, headerNumCols);
      headerRange.setValues([expectedHeaders]);
      headerRange.setFontWeight("bold");
      try { sheet.setFrozenRows(1); } catch(e) { /* Ignore */ }
      Logger.log(`工作表 "${sheet.getName()}" 是空的。已新增標頭列。`);
      return;
   }
   const currentHeaderRange = sheet.getRange(1, firstColIndex, 1, headerNumCols);
   const currentHeader = currentHeaderRange.getValues()[0];
   let needsUpdate = !currentHeader.every((value, index) => value === expectedHeaders[index]);

   if (needsUpdate) {
      currentHeaderRange.setValues([expectedHeaders]);
      currentHeaderRange.setFontWeight("bold");
      Logger.log(`工作表 "${sheet.getName()}" 中的標頭列不正確或不完整。已替換標頭列。`);
   }
   if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
}

/**
 * Sorts the sheet data based on a specific column, skipping the header row.
 * (根據指定欄位對工作表資料進行排序，跳過標頭列。)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object.
 * @param {number} sortColumnIndex The 1-based index of the column to sort by.
 */
function sortSheet(sheet, sortColumnIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    try {
        range.sort({ column: sortColumnIndex, ascending: true });
        Logger.log(`已依欄位 ${sortColumnIndex} 排序工作表 "${sheet.getName()}"。`);
    } catch (e) { Logger.log(`無法排序工作表 "${sheet.getName()}": ${e}`); }
  } else {
      Logger.log(`跳過排序 "${sheet.getName()}"：標頭下方沒有資料列。`);
  }
}