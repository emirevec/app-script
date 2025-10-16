// dataImport.js
// Functions to import and filter data from XLSX file


/**
 * Main import function - orchestrates the import process
 */
function importXlsxData() {
  try {
    const ui = SpreadsheetApp.getUi();
    const config = getLocalConfig();
    
    const allData = fetchXlsxData(config.folderId, config.fileName);
    
    if (!validateDataExists(allData, ui)) {
      return;
    }
    
    const headers = allData[0];
    const filteredData = filterDataByComercial(allData, config.comercial);
    
    if (!validateFilteredData(filteredData, config.comercial, ui)) {
      return;
    }
    
    writeDataToSheet(headers, filteredData, config.sheetNames.import);
    updatePendingSheet();
    
    showSuccessMessage(filteredData.length, config.comercial, ui);
    
  } catch (error) {
    handleImportError(error);
  }
}

/**
 * Fetches and converts XLSX data to array using folder ID and file name
 * @param {string} folderId - The folder ID containing the XLSX file
 * @param {string} fileName - The name of the XLSX file
 * @returns {Array} All data from XLSX file
 */
function fetchXlsxData() {
  const folderId = XLXS_FOLDER_ID;
  const fileName = XLXS_FILE_NAME;
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName(fileName);
  
  if (!files.hasNext()) {
    throw new Error(`Archivo "${fileName}" no encontrado en la carpeta especificada.`);
  }
  
  const xlsxFile = files.next();
  const xlsxBlob = xlsxFile.getBlob();
  
  const tempSs = convertXlsxToSpreadsheet(xlsxBlob);
  const tempSheet = tempSs.getSheets()[0];
  const allData = tempSheet.getDataRange().getValues();
  
  DriveApp.getFileById(tempSs.getId()).setTrashed(true);
  
  return allData;
}

/**
 * Validates that data exists in the file
 * @param {Array} allData - All data from file
 * @param {Ui} ui - SpreadsheetApp UI object
 * @returns {boolean} True if data exists
 */
function validateDataExists(allData, ui) {
  if (allData.length <= 1) {
    ui.alert(
      'Sin datos',
      'El archivo XLSX no contiene datos.',
      ui.ButtonSet.OK
    );
    return false;
  }
  return true;
}

/**
 * Validates that filtered data exists for the user
 * @param {Array} filteredData - Filtered data array
 * @param {string} comercial - User's comercial value
 * @param {Ui} ui - SpreadsheetApp UI object
 * @returns {boolean} True if filtered data exists
 */
function validateFilteredData(filteredData, comercial, ui) {
  if (filteredData.length === 0) {
    ui.alert(
      'Sin datos para tu usuario',
      'No se encontraron registros para el comercial: ' + comercial,
      ui.ButtonSet.OK
    );
    return false;
  }
  return true;
}

/**
 * Writes headers and data to the target sheet
 * @param {Array} headers - Column headers
 * @param {Array} data - Data rows
 * @param {string} sheetName - Target sheet name
 */
function writeDataToSheet(headers, data, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  sheet.clear();
  
  writeHeaders(sheet, headers);
  writeDataRows(sheet, data);
  autoResizeColumns(sheet, headers.length);
}

/**
 * Writes and formats headers
 * @param {Sheet} sheet - Target sheet
 * @param {Array} headers - Column headers
 */
function writeHeaders(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

/**
 * Writes data rows to sheet
 * @param {Sheet} sheet - Target sheet
 * @param {Array} data - Data rows
 */
function writeDataRows(sheet, data) {
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }
}

/**
 * Auto-resizes all columns in the sheet
 * @param {Sheet} sheet - Target sheet
 * @param {number} columnCount - Number of columns
 */
function autoResizeColumns(sheet, columnCount) {
  for (let i = 1; i <= columnCount; i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Shows success message to user
 * @param {number} recordCount - Number of imported records
 * @param {string} comercial - User's comercial value
 * @param {Ui} ui - SpreadsheetApp UI object
 */
function showSuccessMessage(recordCount, comercial, ui) {
  ui.alert(
    'Importación exitosa',
    `Se importaron ${recordCount} registros para ${comercial}.`,
    ui.ButtonSet.OK
  );
}

/**
 * Handles import errors
 * @param {Error} error - The error object
 */
function handleImportError(error) {
  Logger.log('Error in importXlsxData: ' + error.message);
}

/**
 * Main function to import data from XLSX file filtered by current user
 * This function will be triggered from the menu
 */
function importXlsxData() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Get configuration
    const config = getLocalConfig();
    const comercial = config.comercial;
    
    // Get the XLSX file
    const xlsxFile = DriveApp.getFileById(config.xlsxFileId);
    const xlsxBlob = xlsxFile.getBlob();
    
    // Convert XLSX to temporary spreadsheet to read data
    const tempSs = convertXlsxToSpreadsheet(xlsxBlob);
    const tempSheet = tempSs.getSheets()[0];
    const allData = tempSheet.getDataRange().getValues();
    
    // Delete temporary spreadsheet
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
    
    if (allData.length <= 1) {
      ui.alert(
        'Sin datos',
        'El archivo XLSX no contiene datos.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Filter data by comercial
    const headers = allData[0];
    const filteredData = filterDataByComercial(allData, comercial);
    
    if (filteredData.length === 0) {
      ui.alert(
        'Sin datos para tu usuario',
        'No se encontraron registros para el comercial: ' + comercial,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Write to "de facturar" sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let importSheet = ss.getSheetByName(config.sheetNames.import);
    
    // Create sheet if it doesn't exist
    if (!importSheet) {
      importSheet = ss.insertSheet(config.sheetNames.import);
    }
    
    // Clear existing data
    importSheet.clear();
    
    // Write headers
    importSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    importSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // Write filtered data
    if (filteredData.length > 0) {
      importSheet.getRange(2, 1, filteredData.length, filteredData[0].length)
        .setValues(filteredData);
    }
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      importSheet.autoResizeColumn(i);
    }
    
    // Update pending sheet
    updatePendingSheet();
    
    ui.alert(
      'Importación exitosa',
      `Se importaron ${filteredData.length} registros para ${comercial}.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error in importXlsxData: ' + error.message);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Error al importar datos: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Converts an XLSX blob to a temporary Google Spreadsheet
 * @param {Blob} xlsxBlob - The XLSX file blob
 * @returns {Spreadsheet} Temporary spreadsheet with the data
 */
function convertXlsxToSpreadsheet(xlsxBlob) {
  const resource = {
    title: 'temp_xlsx_' + new Date().getTime(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  
  const file = Drive.Files.insert(resource, xlsxBlob, {
    convert: true
  });
  
  return SpreadsheetApp.openById(file.id);
}

/**
 * Filters data rows by comercial name
 * @param {Array} allData - All data including headers
 * @param {string} comercial - The comercial name to filter by
 * @returns {Array} Filtered data rows (without headers)
 */
function filterDataByComercial(allData, comercial) {
  const headers = allData[0];
  const comercialIndex = XLSX_COLUMNS.COMERCIAL;
  const filtered = [];
  
  // Start from row 1 (skip headers)
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (row[comercialIndex] === comercial) {
      filtered.push(row);
    }
  }
  
  return filtered;
}

/**
 * Updates the "de enviar a emisión" sheet with items not yet logged
 * This sheet shows filtered items from "de facturar" that aren't in the Log
 */
function updatePendingSheet() {
  try {
    const config = getRemoteConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get sheets
    const importSheet = ss.getSheetByName(config.sheetNames.import);
    const logSheet = ss.getSheetByName(config.sheetNames.log);
    let pendingSheet = ss.getSheetByName(config.sheetNames.pending);
    
    if (!importSheet) {
      throw new Error(`Sheet "${config.sheetNames.import}" not found`);
    }
    
    // Create pending sheet if it doesn't exist
    if (!pendingSheet) {
      pendingSheet = ss.insertSheet(config.sheetNames.pending);
    }
    
    // Get all data from import sheet
    const importData = importSheet.getDataRange().getValues();
    if (importData.length <= 1) {
      pendingSheet.clear();
      return;
    }
    
    const headers = importData[0];
    const cargaNIndex = XLSX_COLUMNS.CARGA_N;
    
    // Get logged Carga N values
    const loggedCargaNs = new Set();
    if (logSheet && logSheet.getLastRow() > 1) {
      const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues();
      logData.forEach(row => {
        if (row[0]) {
          loggedCargaNs.add(row[0].toString());
        }
      });
    }
    
    // Filter pending items
    const pendingData = [];
    for (let i = 1; i < importData.length; i++) {
      const row = importData[i];
      const cargaN = row[cargaNIndex];
      if (cargaN && !loggedCargaNs.has(cargaN.toString())) {
        pendingData.push(row);
      }
    }
    
    // Clear and write to pending sheet
    pendingSheet.clear();
    pendingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    pendingSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    if (pendingData.length > 0) {
      pendingSheet.getRange(2, 1, pendingData.length, pendingData[0].length)
        .setValues(pendingData);
    }
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      pendingSheet.autoResizeColumn(i);
    }
    
  } catch (error) {
    Logger.log('Error in updatePendingSheet: ' + error.message);
    throw error;
  }
}