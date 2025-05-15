function main(){
  const functionName = 'main';
  const config = getConfig();

  try {
    const exchangeRate = fetchBCRAExchangeRateByDate({date: config.today});
    const averageRate = calculateAverageOfTheLast30Days({sheetName: config.reportSheetName, thirtythValue: exchangeRate.value});
    const updatedExchageRate = {
      date: exchangeRate.date,
      value: exchangeRate.value,
      average: averageRate
    }

    updateExchangeRateInSheet({sheetName: config.reportSheetName, data: updatedExchageRate});
  } catch (error) {
    Logger.log(`${functionName}: An error occurred: ${error}`);
  }
}

