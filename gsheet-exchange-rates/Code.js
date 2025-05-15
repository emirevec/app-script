function main(){
  const functionName = 'main';
  const config = getConfig();

  try {
    const exchangeRate = fetchBCRAExchangeRateByDate({date: config.today});
    const averageRate = calculateAverageRate({sheetName: config.reportSheetName, thirtythValue: exchangeRate.value});
    Logger.log(`${functionName}: Average exchange rate: ${averageRate}`);
    updateExchangeRateInSheet({sheetName: config.reportSheetName, data: exchangeRate});
  } catch (error) {
    Logger.log(`${functionName}: An error occurred: ${error}`);
  }
}

