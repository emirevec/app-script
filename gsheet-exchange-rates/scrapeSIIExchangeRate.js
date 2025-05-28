function scrapeSIIExchangeRate({date}) {
  const functionName = "scrapeSIIExchangeRate";
  
  if (!date) {
    throw new Error(`[${functionName}] Date is required.`);
  }
  
  const year = date.getFullYear();
  const WEB_URL = `https://www.sii.cl/valores_y_fechas/dolar/dolar${year}.htm`;
  const dayOfMonth = date.getDate();
  const monthNames = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
  ];
  const monthName = monthNames[date.getMonth()];

  try {
    Logger.log(`[${functionName}] Fetching content from: ${WEB_URL}`);
    const response = UrlFetchApp.fetch(WEB_URL, { muteHttpExceptions: true });
    const htmlContent = response.getContentText("UTF-8");

    if (response.getResponseCode() !== 200) {
      throw new Error(`Failed to fetch the website. Status code: ${response.getResponseCode()}`);
    }

    const monthBlockRegex = new RegExp(
      `<div[^>]*id=['"]mes_${monthName}['"][^>]*>` + // Start of the current month's div (flexible quotes)
      `([\\s\\S]*?)` + // Non-greedy capture of everything inside
      `(?:<div[^>]*class=['"]meses['"]` + // STOP if we find the next div with class 'meses' (this covers all month divs and mes_all)
      `|$)`, // OR STOP at the very end of the string if it's the last month/data block
      "is" // i: case-insensitive, s: dotall (for multi-line match)
    );
    const monthBlockMatch = htmlContent.match(monthBlockRegex);

    if (!monthBlockMatch || !monthBlockMatch[1]) {
      throw new Error(`Could not find the HTML block for month '${monthName}' (ID: mes_${monthName}). Regex failed to match.`);
    }

    const monthHtml = monthBlockMatch[1];
    
    const dayValueRegex = new RegExp(
      `<th[^>]*>\\s*<strong>\\s*${dayOfMonth}\\s*</strong>\\s*</th>` + // Match the specific day's th
      '\\s*<td[^>]*>([\\d,.]*)<\\/td>', // Match the immediate td following it, and capture the value
      "i" // Case-insensitive
    );

    const dayValueMatch = monthHtml.match(dayValueRegex);
    let clpValue = null;

    if (dayValueMatch && dayValueMatch[1]) {
      const capturedValue = dayValueMatch[1].trim();
      if (capturedValue !== "" && capturedValue !== "&nbsp;") { 
        clpValue = parseFloat(capturedValue.replace(",", "."));
      }
    }

    if (clpValue !== null) {
      Logger.log(`[[${functionName}] SII CLP Exchange Rate: ${clpValue}`);
      return clpValue;
    } else {
      Logger.log(`[[${functionName}] Could not find or extract CLP exchange rate for day ${dayOfMonth} within month '${monthName}' block. Please ensure the website structure hasn't changed.`);
      return null;
    }

  } catch (error) {
    Logger.log(`[${functionName}] An error occurred: ${error}.`);
    return null;
  }
}