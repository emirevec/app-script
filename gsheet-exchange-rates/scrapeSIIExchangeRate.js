function scrapeSIIExchangeRate() {
  const functionName = "scrapeSIIExchangeRate";
  const WEB_URL = "https://www.sii.cl/valores_y_fechas/dolar/dolar2025.htm";
  const today = new Date();
  const dayOfMonth = today.getDate(); // Get the current day of the month

  try {
    Logger.log(`${functionName}: Fetching content from: ${WEB_URL}`);
    const response = UrlFetchApp.fetch(WEB_URL, { muteHttpExceptions: true });
    const htmlContent = response.getContentText("UTF-8");

    if (response.getResponseCode() !== 200) {
      throw new Error(`${functionName}: Failed to fetch the website. Status code: ${response.getResponseCode()}`);
    }

    // Construct a regular expression to find the <td> containing today's date
    // and the <td> immediately following it with the value.
    const dayRegex = new RegExp(`<th width="40"><strong>${dayOfMonth}</strong></th>\\s*<td width="200">([^<]*)</td>`, "s");
    const dayMatch = htmlContent.match(dayRegex);

    let clpValue = null;

    if (dayMatch && dayMatch[1]) {
      clpValue = parseFloat(dayMatch[1].replace(",", "."));
      Logger.log(`[${functionName}] SII CLP Exchange Rate for ${Utilities.formatDate(today, Session.getTimeZone(), "yyyy-MM-dd")}: ${clpValue}`);
      return clpValue;
    } else {
      Logger.log(`[${functionName}] Could not find CLP exchange rate for day ${dayOfMonth}. Please check the website.`);
      return null;
    }

  } catch (error) {
    Logger.log(`${functionName}: An error occurred: ${error}`);
    return null;
  }
}