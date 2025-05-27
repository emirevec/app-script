function calculate30DaysAverage({sheetName, thirtythValue}) {
  const functionName = 'calculate30DaysAverage';

  if (!sheetName) {
    throw new Error(`${functionName}: SheetName is required.`); 
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }

    const rateValues = sheet.getRange(2, 4, 29, 1).getValues();
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
      Logger.log(`${functionName}: Average exchange rate of ${validRateCount} values: ${averageRate}`);
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

function getLastValidRateFromSheet({sheetName}) {
  const functionName = `getLastValidRateFromSheet (${sheetName})`;

  if (!sheetName) {
    throw new Error(`${functionName}: Sheetname is required.`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  lastValue = sheet.getRange(2, 2).getValue();

  return lastValue; 
}

function updateExchangeRateInSheet({sheetName, data}) {
  const functionName = "updateExchangeRateInSheet";

  if (!sheetName) {
    throw new Error(`${functionName}: Sheet name is required.`);
  }

  if (!data || !data.date || !data.dailyRate || !data.averageRate) {
    throw new Error(`${functionName}: Data with all values are required, except for originalRate wich can be null.`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`${functionName}: Sheet "${sheetName}" not found.`);
  }

  sheet.insertRowBefore(2);
  sheet.getRange(2, 1).setValue(data.date);
  sheet.getRange(2, 2).setValue(data.dailyRate);
  sheet.getRange(2, 3).setValue(data.averageRate);
  sheet.getRange(2, 4).setValue(data.originalRate);
  Logger.log(`${functionName}:Successfully updated exchange rate on ${data.date} in sheet "${sheetName}".`);
}
