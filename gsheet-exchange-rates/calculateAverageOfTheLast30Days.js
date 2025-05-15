function calculateAverageOfTheLast30Days({sheetName, thirtythValue}) {
  const functionName = 'calculateAverageOfTheLast30Days';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`${functionName}: Sheet "${sheetName}" not found.`);
      return null;
    }

    const rateValues = sheet.getRange(2, 2, 29, 1).getValues();
    let totalRate = 0;
    let validRateCount = 0;

    for (const row of rateValues) {
      const rate = parseFloat(row[0]);
      if (!isNaN(rate)) {
        totalRate += rate;
        validRateCount++;
      }
    }

    const thirtythValueFloat = parseFloat(thirtythValue);
    if (!isNaN(thirtythValueFloat)) {
      totalRate += thirtythValueFloat;
      validRateCount++;
    } else {
      Logger.log(`${functionName}: Provided thirtythValue "${thirtythValue}" is not a valid number and will be excluded from the average.`);
    }

    if (validRateCount > 0) {
      const averageRate = totalRate / validRateCount;
      Logger.log(`${functionName}: Average exchange rate of ${validRateCount} values (up to 29 from sheet + 1 arg): ${averageRate}`);
      return averageRate;
    } else {
      Logger.log(`${functionName}: No valid exchange rate data found to calculate the average.`);
      return null;
    }

  } catch (error) {
    Logger.log(`${functionName}: An error occurred: ${error}`);
    return null;
  }
}
