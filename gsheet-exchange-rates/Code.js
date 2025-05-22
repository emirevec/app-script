function main(){
  const functionName = 'main';
  const config = getConfig();

  handleCurrencyExchangeRate({sheetName: config.usdArsDivisa, date: config.today , fetcherFunction: ()=>fetchBCRAExchangeRateByDate({ date: config.today })});
  Logger.log(`${functionName}: Finished processing for ${config.usdArsDivisa} on ${config.today}`);
}

