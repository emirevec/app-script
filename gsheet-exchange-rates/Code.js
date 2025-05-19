function main(){
  const functionName = 'main';
  const config = getConfig();

  try {
    let divisaVendedorBCRA = fetchBCRAExchangeRateByDate({date: config.today});

    if (!divisaVendedorBCRA) {
      Logger.log(`${functionName}: The last available value will be repeated.`);
      divisaVendedorBCRA = getLastBCRAValidValue({sheetName: config.reportSheetName, date: config.today});
    }

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

