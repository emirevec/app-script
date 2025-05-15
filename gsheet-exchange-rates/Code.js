function main(){
  const functionName = 'main';
  const config = getConfig();

  try {
    const divisaVendedorBCRA = fetchBCRAExchangeRateByDate({date: config.today});
    const averageRate = calculateAverageOfTheLast30Days({sheetName: config.reportSheetName, thirtythValue: divisaVendedorBCRA.value});
    const billeteVendedorBNA = scrapeBNAExchangeRate()
    const updatedExchageRate = {
      date: config.today,
      USD_ARS_Divisa: divisaVendedorBCRA.value,
      USD_ARS_Divisa_Average: averageRate,
      USD_ARS_Billete: billeteVendedorBNA,
    }

    updateExchangeRateInSheet({sheetName: config.reportSheetName, data: updatedExchageRate});
  } catch (error) {
    Logger.log(`${functionName}: An error occurred: ${error}`);
  }
}

