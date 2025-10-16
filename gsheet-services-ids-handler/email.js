// emailService.js
// Functions to send emails and log service IDs

/**
 * Main function to send selected rows via email and log them
 * User must select rows in the "de enviar a emisión" sheet
 */
function sendSelectedToEmision() {
  try {
    const ui = SpreadsheetApp.getUi();
    const config = getRemoteConfig();
    const localConfig = getLocalConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    
    // Verify we're on the pending sheet
    if (activeSheet.getName() !== config.sheetNames.pending) {
      ui.alert(
        'Hoja incorrecta',
        `Por favor selecciona las filas en la hoja "${config.sheetNames.pending}".`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Get selected range
    const selectedRange = activeSheet.getActiveRange();
    if (!selectedRange) {
      ui.alert(
        'Sin selección',
        'Por favor selecciona las filas que deseas enviar.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Get selected rows
    const startRow = selectedRange.getRow();
    const numRows = selectedRange.getNumRows();
    
    // Don't allow header selection
    if (startRow === 1) {
      ui.alert(
        'Selección inválida',
        'No puedes seleccionar la fila de encabezados.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Get data from selected rows
    const selectedData = activeSheet.getRange(startRow, 1, numRows, activeSheet.getLastColumn())
      .getValues();
    
    if (selectedData.length === 0) {
      ui.alert(
        'Sin datos',
        'No hay datos en las filas seleccionadas.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Confirm before sending
    const response = ui.alert(
      'Confirmar envío',
      `¿Enviar ${selectedData.length} servicio(s) a emisión?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Extract Carga N and Remitente from selected rows
    const cargaNIndex = XLSX_COLUMNS.CARGA_N;
    const remitenteIndex = XLSX_COLUMNS.REMITENTE;
    
    const servicesToSend = [];
    for (let i = 0; i < selectedData.length; i++) {
      const row = selectedData[i];
      const cargaN = row[cargaNIndex];
      const remitente = row[remitenteIndex];
      
      if (cargaN) {
        servicesToSend.push({
          cargaN: cargaN,
          remitente: remitente || ''
        });
      }
    }
    
    if (servicesToSend.length === 0) {
      ui.alert(
        'Sin datos válidos',
        'No se encontraron servicios válidos en la selección.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Send email
    sendEmailWithServices(servicesToSend, config);
    
    // Log services
    logServices(servicesToSend, config);
    
    // Update pending sheet
    updatePendingSheet();
    
    ui.alert(
      'Envío exitoso',
      `Se enviaron ${servicesToSend.length} servicio(s) a emisión.\n\n` +
      `Email enviado a: ${config.emailRecipient}`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error in sendSelectedToEmision: ' + error.message);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Error al enviar servicios: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Sends an email with the list of services
 * @param {Array} services - Array of service objects with cargaN and remitente
 * @param {Object} config - Configuration object
 */
function sendEmailWithServices(services, config) {
  const localConfig = getLocalConfig();
  const currentUser = getCurrentUserEmail();
  const comercial = getCurrentUserComercial();
  
  // Build email body
  let emailBody = `Hola,\n\n`;
  emailBody += `El usuario ${comercial} (${currentUser}) solicita la emisión de los siguientes servicios:\n\n`;
  
  // Add table header
  emailBody += `Carga N\t\tRemitente\n`;
  emailBody += `${'='.repeat(60)}\n`;
  
  // Add each service
  services.forEach(service => {
    emailBody += `${service.cargaN}\t\t${service.remitente}\n`;
  });
  
  emailBody += `\n${'='.repeat(60)}\n`;
  emailBody += `Total de servicios: ${services.length}\n\n`;
  emailBody += `Fecha de solicitud: ${new Date().toLocaleString('es-AR')}\n\n`;
  emailBody += `Saludos,\n`;
  emailBody += `Sistema de Gestión de Servicios`;
  
  // Build HTML version for better formatting
  let htmlBody = `<p>Hola,</p>`;
  htmlBody += `<p>El usuario <strong>${comercial}</strong> (${currentUser}) solicita la emisión de los siguientes servicios:</p>`;
  htmlBody += `<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">`;
  htmlBody += `<thead><tr style="background-color: #f0f0f0;">`;
  htmlBody += `<th>Carga N</th><th>Remitente</th>`;
  htmlBody += `</tr></thead><tbody>`;
  
  services.forEach(service => {
    htmlBody += `<tr>`;
    htmlBody += `<td>${service.cargaN}</td>`;
    htmlBody += `<td>${service.remitente}</td>`;
    htmlBody += `</tr>`;
  });
  
  htmlBody += `</tbody></table>`;
  htmlBody += `<p><strong>Total de servicios:</strong> ${services.length}</p>`;
  htmlBody += `<p><small>Fecha de solicitud: ${new Date().toLocaleString('es-AR')}</small></p>`;
  htmlBody += `<p>Saludos,<br><em>Sistema de Gestión de Servicios</em></p>`;
  
  // Send email
  MailApp.sendEmail({
    to: config.emailRecipient,
    subject: localConfig.email.subject,
    body: emailBody,
    htmlBody: htmlBody
  });
  
  Logger.log(`Email sent to ${config.emailRecipient} with ${services.length} services`);
}

/**
 * Logs services to the Log sheet
 * @param {Array} services - Array of service objects with cargaN
 * @param {Object} config - Configuration object
 */
function logServices(services, config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(config.sheetNames.log);
  
  // Create log sheet if it doesn't exist
  if (!logSheet) {
    logSheet = ss.insertSheet(config.sheetNames.log);
    // Add headers
    logSheet.getRange(1, 1, 1, 2).setValues([LOG_HEADERS]);
    logSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  
  // Get current timestamp
  const timestamp = new Date();
  
  // Prepare log entries
  const logEntries = services.map(service => [
    service.cargaN,
    timestamp
  ]);
  
  // Append to log sheet
  const lastRow = logSheet.getLastRow();
  if (logEntries.length > 0) {
    logSheet.getRange(lastRow + 1, 1, logEntries.length, 2)
      .setValues(logEntries);
  }
  
  Logger.log(`Logged ${logEntries.length} services to Log sheet`);
}

/**
 * Helper function to manually refresh the pending sheet
 * Can be called from the menu
 */
function refreshPendingSheet() {
  try {
    updatePendingSheet();
    SpreadsheetApp.getUi().alert(
      'Actualización exitosa',
      'La hoja de pendientes ha sido actualizada.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    Logger.log('Error in refreshPendingSheet: ' + error.message);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Error al actualizar pendientes: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}