function fetchBCRAExchangeRateByDate({date}){
  const functionName = "fetchBCRAExchangeRateByDate";

  if (!date) {
    throw new Error(`[${functionName}] Date is required.`);
  }

  /* const beforeToday = new Date();
  beforeToday.setDate(date.getDate() - 1);  */
  const formattedDate = Utilities.formatDate(date, "America/Argentina/Buenos_Aires", "yyyy-MM-dd");
  const API_BASE_URL = "https://api.bcra.gob.ar/estadisticascambiarias/v1.0/Cotizaciones/usd?"
  const queryParams = `fechaDesde=${formattedDate}&fechaHasta=${formattedDate}`
  const apiEndpoint = API_BASE_URL + queryParams;

  try {
    Logger.log(`[${functionName}] Fetching exchange rate from: ${apiEndpoint}`);
    const response = UrlFetchApp.fetch(apiEndpoint);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      throw new Error(`Failed to fetch data from BCRA API. Status code: ${responseCode}`);
    }
    
    const jsonResponse = response.getContentText();
    const data = JSON.parse(jsonResponse);

    if (!data || !data.results) {
      throw new Error("No results found in the API response.");
    }

    if (data.results.length === 0) {
      Logger.log(`[${functionName}] No data found for the specified date.`);
      return null;
    }

    const latestResult = data.results[0]; 

    if (!latestResult || !latestResult.fecha || !latestResult.detalle || latestResult.detalle.length === 0) {
      throw new Error("No valid data found for the specified date.");
    }

    const exchangeRateData = latestResult.detalle[0];

    if(!exchangeRateData || !exchangeRateData.tipoCotizacion) {
      throw new Error("No valid exchange rate data found.");
    }

    const value = exchangeRateData.tipoCotizacion;

    return value;
  }

  catch (error) {
    Logger.log(`[${functionName}] An error occurred: ${error}`);
    return null
  }
}
