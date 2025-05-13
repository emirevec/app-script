function main(){
  const config = getConfig();  
  try {
    const exchangeRate = fetchBCRAExchangeRateByDate({date: config.today});
  } catch (error) {
    Logger.log(`An error occurred: ${error}`);
  }
}
