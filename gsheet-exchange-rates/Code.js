function main(){
  const functionName = 'main';
  const config = getConfig();

  handleCurrencyExchangeRate({sheetName: config.usdArsDivisa, date: config.today , fetcherFunction: ()=>fetchBCRAExchangeRateByDate({ date: config.today })});
  handleCurrencyExchangeRate({sheetName: config.usdArsBillete, date: config.today , fetcherFunction: ()=>scrapeBNAExchangeRate()});
  handleCurrencyExchangeRate({sheetName: config.usdClpObservado, date: config.today , fetcherFunction: ()=>scrapeSIIExchangeRateByDate({ date: config.today })});
  Logger.log(`[${functionName}] Finished processing for ${config.usdArsDivisa}, ${config.usdArsBillete} and ${config.usdClpObservado} on ${config.today}`); 
  
}
