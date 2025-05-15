function updateExchangeRateInSheet({sheetName, data}) {
  const functionName = "updateExchangeRateInSheet";

  if (!sheetName) {
    throw new Error(`${functionName}: Sheet name is required.`);
  }

  if (!data || !data.date || !data.value || !data.average) {
    throw new Error(`${functionName}: Data with date, value, and average is required.`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`${functionName}: Sheet "${sheetName}" not found.`);
  }

  sheet.insertRowBefore(2);
  sheet.getRange(2, 1).setValue(data.date);
  sheet.getRange(2, 2).setValue(data.value);
  sheet.getRange(2, 3).setValue(data.average);
  Logger.log(`${functionName}:Successfully updated exchange rate on ${data.date} in sheet "${sheetName}".`);
}
