function main(){
  const functionName = 'main';
  const config = getConfig();

  try {
    const exchangeRate = fetchBCRAExchangeRateByDate({date: config.today});
    updateExchangeRateInSheet({sheetName: config.reportSheetName, data: exchangeRate});
  } catch (error) {
    Logger.log(`${functionName}: An error occurred: ${error}`);
  }
}
