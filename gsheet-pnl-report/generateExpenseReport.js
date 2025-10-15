const userInputRow = 1;
const yearInputColumn = 2;

function generateExpenseReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  const reportSheet = ss.getSheetByName(REPORT_SHEET_NAME);

  if (!sourceSheet || !reportSheet) {
    Logger.log(`Error: One or both sheets ("${SOURCE_SHEET_NAME}", "${REPORT_SHEET_NAME}") not found.`);
    SpreadsheetApp.getUi().alert('Error', `One or both sheets ("${SOURCE_SHEET_NAME}", "${REPORT_SHEET_NAME}") not found. Please check sheet names.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const userFilterYear = reportSheet.getRange(userInputRow, yearInputColumn).getValue();
  let filterYear = userFilterYear ? userFilterYear : getCurrentYear();

  const sourceDataRange = sourceSheet.getDataRange();
  const sourceData = sourceDataRange.getValues();
  const headerRow = sourceData[0];
  const dataRows = sourceData.slice(1);

  const yearColumnIndex = headerRow.indexOf(YEAR);
  const periodColumnIndex = headerRow.indexOf(PERIOD);
  const dateColumnIndex = headerRow.indexOf(DATE); // Not used in this specific report, but good to have
  const vendorColumnIndex = headerRow.indexOf(VENDOR); // Not used in this specific report, but good to have
  const expenseTypeColumnIndex = headerRow.indexOf(EXPENSE_TYPE);
  const amountColumnIndex = headerRow.indexOf(AMOUNT);

  if (yearColumnIndex === -1 || periodColumnIndex === -1 || expenseTypeColumnIndex === -1 || amountColumnIndex === -1 || expenseTypeColumnIndex === -1) {
    Logger.log("Error: One or more required column headers (Year, Period, Expense Type, Amount) not found in the source sheet.");
    SpreadsheetApp.getUi().alert('Error', 'One or more required column headers (Year, Period, Expense Type, Amount) not found in the source sheet. Please check your sheet headers.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  reportSheet.clearContents();
  let reportRow = 1;

  const expenseType = EXPENSE_TYPE_FILTER

  // Check if the filtered expense type exists in the data
  const uniqueExpenseTypesInSource = [...new Set(dataRows.map(row => row[expenseTypeColumnIndex]))];
  if (!uniqueExpenseTypesInSource.includes(expenseType)) {
    Logger.log(`Error: Expense type "${expenseType}" not found in the source data.`);
    SpreadsheetApp.getUi().alert('Error', `Expense type "${expenseType}" not found in the source data. Please check the expense type you provided.`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  reportSheet.getRange(reportRow, 1).setValue(`Expense Type: ${expenseType} for Year: ${filterYear}`);
  reportRow++;

  const groupedData = {};

  dataRows.forEach(row => {
    // Filter by year and the provided expense type
    if (row[yearColumnIndex] == filterYear && row[expenseTypeColumnIndex] === expenseType) {
      const period = row[periodColumnIndex];
      const amount = parseFloat(row[amountColumnIndex]) || 0;

      if (groupedData[period]) {
        groupedData[period] += amount;
      } else {
        groupedData[period] = amount;
      }
    }
  });

  // 2. Change the way it displays the data, from vertical to horizontal
  const periods = Object.keys(groupedData).sort(); // Get sorted periods for consistent order
  const amounts = periods.map(period => groupedData[period]);

  // Set headers horizontally
  reportSheet.getRange(reportRow, 1).setValue("Período"); 
  reportSheet.getRange(reportRow, 2, 1, periods.length).setValues([periods]);
  reportRow++;

  // Set amounts horizontally
  reportSheet.getRange(reportRow, 1).setValue(expenseType); // Label for the row
  reportSheet.getRange(reportRow, 2, 1, amounts.length).setValues([amounts]);
  reportRow++;

  Logger.log(`Expense report for ${expenseType} generated successfully!`);
  
  // 3. Add the necessary code to send the result by email
  return {expenseType, filterYear, periods, amounts}
}
