function handleCurrencyExchangeRate({sheetName, date, fetcherFunction}) {
  const functionName = `handleCurrencyExchangeRate (${sheetName})`;

  if (!sheetName || !date || !fetcherFunction) {
    throw new Error(`${functionName}: sheetName, date, and fetcherFunction are required.`);
  }

  try {
    const newRate = fetcherFunction();
    let dailyRate = newRate;
    
    if (!dailyRate) {
      dailyRate = getLastValidRateFromSheet({sheetName: sheetName});
    }
    
    const averageRate = calculate30DaysAverage({sheetName: sheetName, thirtythValue: newRate});
    
    const data = {
      date: date,
      dailyRate: dailyRate,
      averageRate: averageRate,
      originalRate: newRate,
    }
    
    updateExchangeRateInSheet({ sheetName: sheetName, data: data});

  } catch (error) {
    Logger.log(`${functionName}: An error occurred while processing: ${error}`);
  }
}
