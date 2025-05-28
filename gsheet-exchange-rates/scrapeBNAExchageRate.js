function scrapeBNAExchangeRate() {
  const functionName = "scrapeBNAExchangeRate";
  const WEB_URL = "https://www.bna.com.ar/";
  const ventaRegex = /<td class="tit">Dolar U\.S\.A<\/td>\s*<td>[\d,.]+<\/td>\s*<td>([\d,.]+)<\/td>/s;
  
  try {
    Logger.log(`[${functionName}] Fetching content from: ${WEB_URL}`);
    const response = UrlFetchApp.fetch(WEB_URL, { muteHttpExceptions: true });
    const htmlContent = response.getContentText("UTF-8");

    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch the website. Status code: ${response.getResponseCode()}`);
    }

    const ventaMatch = htmlContent.match(ventaRegex);
    let ventaRate = null;

    if (ventaMatch && ventaMatch[1]) {
      ventaRate = parseFloat(ventaMatch[1].replace(",", "."));
      Logger.log(`[${functionName}] BNA USD Billete - Venta: ${ventaRate}`);
      return ventaRate;
    } else {
      Logger.log(`[${functionName}]Could not find 'Dolar U.S.A' venta rate using regular expression. Please ensure the website structure hasn't changed.`);
      return null;
    }

  } catch (error) {
    Logger.log(`[${functionName}] Failed to extract 'Cotización Venta' rate using regular expression. ${error}`);
    return null;
  }
}
