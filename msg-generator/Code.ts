function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("訊息產生器")
    .addItem("生成訊息", "applyMessageTemplate")
    .addToUi();
}

function applyMessageTemplate() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("訊息產生器");
  if (!sheet) {
    throw new Error("Could not find sheet '訊息產生器'");
  }

  // 取得變數映射區：從 A2:B 讀入變數與對應值
  const inputRange = sheet.getRange("A2:B20").getValues();
  const variables = {};
  inputRange.forEach(([key, value]) => {
    if (key && value) {
      variables[key] = value;
    }
  });

  // 取得模板內容
  const templateSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("情境模板");
  if (!templateSheet) {
    throw new Error("Could not find sheet '情境模板'");
  }
  const templateRange = templateSheet.getRange("A2:B20").getValues();
  const templates = {};
  templateRange.forEach(([key, value]) => {
    if (key && value) {
      templates[key] = value;
    }
  });

  const msgRange = sheet.getRange("D2:E20").getValues();
  for (let i = 0; i < msgRange.length; i++) {
    const key = msgRange[i][0];
    const value = msgRange[i][1];
    if (key && templates[key] && !value) {
      let template = templates[key];
      // 套用所有變數
      for (let key in variables) {
        const placeholder = `{${key}}`;
        template = template.split(placeholder).join(variables[key]);
      }
      // 寫入生成訊息
      sheet.getRange(`E${2 + i}`).setValue(template);
    }
  }
}
