function updateExchangeRateInSheet({sheetName, data}) {
  const functionName = "updateExchangeRateInSheet";

  if (!sheetName || !data) {
    throw new Error(`${functionName}: Sheet name and data are both required.`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`${functionName}: Sheet "${sheetName}" not found.`);
  }

  sheet.insertRowBefore(2);
  sheet.getRange(2, 1).setValue(data.date);
  sheet.getRange(2, 2).setValue(data.value);
  Logger.log(`Successfully updated exchange rate on ${data.date} in sheet "${sheetName}".`);
}
