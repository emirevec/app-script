function fetchBCRAExchangeRateByDate({date: date}) {
  const formattedDate = Utilities.formatDate(date, "America/Argentina/Buenos_Aires", "yyyy-MM-dd");
  const API_BASE_URL = "https://api.bcra.gob.ar/estadisticascambiarias/v1.0/Cotizaciones/usd?"
  const queryParams = `fechaDesde=${formattedDate}&fechaHasta=${formattedDate}`
  const apiEndpoint = API_BASE_URL + queryParams;
  //const apiEndpoint = "https://api.bcra.gob.ar/estadisticascambiarias/v1.0/Cotizaciones/usd?fechaDesde=2025-05-09&fechaHasta=2025-05-09"

  try {
    Logger.log(`Fetching exchange rate from: ${apiEndpoint}`);
    const response = UrlFetchApp.fetch(apiEndpoint);
    const jsonResponse = response.getContentText();
    const data = JSON.parse(jsonResponse);

    // 3. Extract the USD to ARS official exchange rate for the current day
    // Assuming the API returns the latest rate as the first element
    const latestRateData = data[0]; 

    if (data && data.results && data.results.length > 0) {
      const latestResult = data.results[0]; // The first element in results array

      if (latestResult && latestResult.fecha && latestResult.detalle && latestResult.detalle.length > 0) {
        const exchangeRateData = latestResult.detalle[0]; // The first (and likely only) item in the detalle array

        if (exchangeRateData && exchangeRateData.tipoCotizacion) {
          const exchangeRate = exchangeRateData.tipoCotizacion;
          const date = new Date(latestResult.fecha); // Convert the date string to a Date object

          // Get the active Google Sheet and the target sheet
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          const sheet = ss.getSheetByName(REPORT_SHEET_NAME); // Make sure REPORT_SHEET_NAME is defined

          if (sheet) {
            // Write the date and exchange rate to cell E2
            sheet.getRange("E2").setValue(date);
            sheet.getRange("F2").setValue(exchangeRate); // Assuming you want the rate in the next column (F)

            Logger.log(`Successfully updated exchange rate on ${date} to ${exchangeRate}`);
          } else {
            Logger.log(`Sheet "${REPORT_SHEET_NAME}" not found.`);
          }
        } else {
          Logger.log("Could not find 'tipoCotizacion' in the detalle.");
        }
      } else {
        Logger.log("Could not find 'fecha' or 'detalle' in the results.");
      }
    } else {
      Logger.log("Could not find any results in the API response.");
    }

  } catch (error) {
    Logger.log(`An error occurred: ${error}`);
  }
}
