function getLastBCRAValidValue({sheetName, date}) {
  const functionName = "getLastBCRAValidValue";

  if (!sheetName || !date) {
    throw new Error(`${functionName}: Sheetname and date are both required.`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const lastValue = sheet.getRange(3, 2).getValue();

  return {date: date , value:lastValue};
}