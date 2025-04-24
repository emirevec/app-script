function generateReport() {

  // 2. Get references to the source and report sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  const reportSheet = ss.getSheetByName(REPORT_SHEET_NAME);

  // 3. Check if the sheets exist
  if (!sourceSheet || !reportSheet) {
    Logger.log(`Error: One or both sheets ("${sourceSheetName}", "${reportSheetName}") not found.`);
    return;
  }

  // 4. Get all the data from the source sheet
  // Assuming your data starts from the first row and goes to the last row with data
  const sourceDataRange = sourceSheet.getDataRange();
  const sourceData = sourceDataRange.getValues();
  const headerRow = sourceData[0]; // Assuming the first row contains headers
  const dataRows = sourceData.slice(1); // Get data excluding the header row

  // 5. Identify the column indices (0-based) for the relevant columns
  const yearColumnIndex = headerRow.indexOf("year");
  const periodColumnIndex = headerRow.indexOf("period");
  const dateColumnIndex = headerRow.indexOf("date");
  const vendorColumnIndex = headerRow.indexOf("vendor");
  const expenseTypeColumnIndex = headerRow.indexOf("expenseType");
  const amountColumnIndex = headerRow.indexOf("amount");

  // Basic error checking for column headers
  if (yearColumnIndex === -1 || periodColumnIndex === -1 || expenseTypeColumnIndex === -1 || amountColumnIndex === -1) {
    Logger.log("Error: One or more required column headers (year, period, expenseType, amount) not found in the source sheet.");
    return;
  }

  // 6. Get the unique expense types from your data (for filtering)
  const uniqueExpenseTypes = [...new Set(dataRows.map(row => row[expenseTypeColumnIndex]))];

  // 7. Prepare the report sheet for new data (clear existing content)
  reportSheet.clearContents();
  let reportRow = 1; // Start writing the report from the first row

  // 8. Loop through each unique expense type and generate a section in the report
  uniqueExpenseTypes.forEach(expenseType => {
    // Add a header for the current expense type
    reportSheet.getRange(reportRow, 1).setValue(`Expense Type: ${expenseType}`);
    reportRow++;

    // Add column headers for the grouped report
    reportSheet.getRange(reportRow, 1).setValue("Period");
    reportSheet.getRange(reportRow, 2).setValue("Total Amount");
    reportRow++;

    // Create an object to store the grouped data by period
    const groupedData = {};

    // Filter data by the current expense type and group by period
    dataRows.forEach(row => {
      if (row[expenseTypeColumnIndex] === expenseType) {
        const period = row[periodColumnIndex];
        const amount = parseFloat(row[amountColumnIndex]) || 0; // Ensure amount is a number

        if (groupedData[period]) {
          groupedData[period] += amount;
        } else {
          groupedData[period] = amount;
        }
      }
    });

    // Write the grouped data to the report sheet
    for (const period in groupedData) {
      reportSheet.getRange(reportRow, 1).setValue(period);
      reportSheet.getRange(reportRow, 2).setValue(groupedData[period]);
      reportRow++;
    }

    // Add an empty row for separation between expense types
    reportRow++;
  });

  Logger.log("Expense report generated successfully!");
}
