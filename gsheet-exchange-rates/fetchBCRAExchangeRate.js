function fetchBCRAExchangeRateByDate({date}){
  const functionName = "fetchBCRAExchangeRateByDate";
  if (!date) {
    throw new Error("Date is required.");
  }
  const formattedDate = Utilities.formatDate(date, "America/Argentina/Buenos_Aires", "yyyy-MM-dd");
  const API_BASE_URL = "https://api.bcra.gob.ar/estadisticascambiarias/v1.0/Cotizaciones/usd?"
  const queryParams = `fechaDesde=${formattedDate}&fechaHasta=${formattedDate}`
  const apiEndpoint = API_BASE_URL + queryParams;

  try {
    Logger.log(`${functionName}: Fetching exchange rate from: ${apiEndpoint}`);
    const response = UrlFetchApp.fetch(apiEndpoint);
    const jsonResponse = response.getContentText();
    const data = JSON.parse(jsonResponse);

    if (!data || !data.results || data.results.length === 0) {
      throw new Error("No results found in the API response.");
    }

    const latestResult = data.results[0]; 

    if (!latestResult || !latestResult.fecha || !latestResult.detalle || latestResult.detalle.length === 0) {
      throw new Error("No valid data found for the specified date.");
    }

    const exchangeRateData = latestResult.detalle[0];

    if(!exchangeRateData || !exchangeRateData.tipoCotizacion) {
      throw new Error("No valid exchange rate data found.");
    }

    const date = new Date(latestResult.fecha); 
    const value = exchangeRateData.tipoCotizacion;

    return { date, value };
  }

  catch (error) {
    Logger.log(`${functionName}: An error occurred: ${error}`);
  }
}
