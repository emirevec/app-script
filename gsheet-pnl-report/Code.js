function generateReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  const reportSheet = ss.getSheetByName(REPORT_SHEET_NAME);

  if (!sourceSheet || !reportSheet) {
    Logger.log(`Error: One or both sheets ("${SOURCE_SHEET_NAME}", "${REPORT_SHEET_NAME}") not found.`);
    return;
  }

  const sourceDataRange = sourceSheet.getDataRange();
  const sourceData = sourceDataRange.getValues();
  const headerRow = sourceData[0]; 
  const dataRows = sourceData.slice(1); 

  const yearColumnIndex = headerRow.indexOf(YEAR);
  const periodColumnIndex = headerRow.indexOf(PERIOD);
  const dateColumnIndex = headerRow.indexOf(DATE);
  const vendorColumnIndex = headerRow.indexOf(VENDOR);
  const expenseTypeColumnIndex = headerRow.indexOf(EXPENSE_TYPE);
  const amountColumnIndex = headerRow.indexOf(AMOUNT);

  if (yearColumnIndex === -1 || periodColumnIndex === -1 || expenseTypeColumnIndex === -1 || amountColumnIndex === -1) {
    Logger.log("Error: One or more required column headers (year, period, expenseType, amount) not found in the source sheet.");
    return;
  }

  const uniqueExpenseTypes = [...new Set(dataRows.map(row => row[expenseTypeColumnIndex]))];

  reportSheet.clearContents();
  let reportRow = 1; 

  uniqueExpenseTypes.forEach(expenseType => {
    reportSheet.getRange(reportRow, 1).setValue(`Expense Type: ${expenseType}`);
    reportRow++;

    reportSheet.getRange(reportRow, 1).setValue("Period");
    reportSheet.getRange(reportRow, 2).setValue("Total Amount");
    reportRow++;

    const groupedData = {};

    dataRows.forEach(row => {
      if (row[expenseTypeColumnIndex] === expenseType) {
        const period = row[periodColumnIndex];
        const amount = parseFloat(row[amountColumnIndex]) || 0;

        if (groupedData[period]) {
          groupedData[period] += amount;
        } else {
          groupedData[period] = amount;
        }
      }
    });

    for (const period in groupedData) {
      reportSheet.getRange(reportRow, 1).setValue(period);
      reportSheet.getRange(reportRow, 2).setValue(groupedData[period]);
      reportRow++;
    }

    reportRow++;
  });

  Logger.log("Expense report generated successfully!");
}
