function testConfiguration() {
  try {
    Logger.log('Testing configuration...');
    
    const userEmail = Session.getActiveUser().getEmail();
    
    Logger.log('✅ Configuration loaded successfully');
    //Logger.log('XLSX File ID: ' + config.xlsxFileId);
    Logger.log('User Comercial: ' + config.comercial);
    Logger.log('Current User Email: ' + userEmail);
    
    SpreadsheetApp.getUi().alert(
      'Test de Configuración',
      '✅ Configuración cargada correctamente\n\n' +
      'Usuario: ' + userEmail + '\n' +
      'Comercial: ' + config.comercial + '\n\n' +
      'Revisa los Logs para más detalles.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('❌ Configuration test failed: ' + error.message);
    SpreadsheetApp.getUi().alert(
      'Error en Configuración',
      '❌ Error: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}