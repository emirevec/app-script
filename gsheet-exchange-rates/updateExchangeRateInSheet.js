function updateExchangeRateInSheet({sheetName, data}) {
  const functionName = "updateExchangeRateInSheet";

  if (!sheetName) {
    throw new Error(`${functionName}: Sheet name is required.`);
  }

  if (!data || !data.date || !data.USD_ARS_Divisa || !data.USD_ARS_Divisa_Average || !data.USD_ARS_Billete) {
    throw new Error(`${functionName}: Data with date, USD_ARS_Divisa, USD_ARS_Divisa_Average, and USD_ARS_Billete is required.`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`${functionName}: Sheet "${sheetName}" not found.`);
  }

  sheet.insertRowBefore(2);
  sheet.getRange(2, 1).setValue(data.date);
  sheet.getRange(2, 2).setValue(data.USD_ARS_Divisa);
  sheet.getRange(2, 3).setValue(data.USD_ARS_Divisa_Average);
  sheet.getRange(2, 4).setValue(data.USD_ARS_Billete);
  Logger.log(`${functionName}:Successfully updated exchange rate on ${data.date} in sheet "${sheetName}".`);
}
