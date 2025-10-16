// ui.js
// User interface and menu functions

/**
 * Creates custom menu when the spreadsheet is opened
 * This is a Simple Trigger that runs automatically
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const localConfig = getLocalConfig();
  
  const menu = ui.createMenu(localConfig.menu.title);
  
  localConfig.menu.items.forEach(item => {
    menu.addItem(item.name, item.functionName);
  });
  
  menu.addToUi();
}

/**
Shows a welcome message with instructions
Can be called manually or from menu

function showWelcomeMessage() {
  const ui = SpreadsheetApp.getUi();
  const config = getRemoteConfig();
  const comercial = config.comercial;
  const userEmail = Session.getActiveUser().getEmail();
  
  let message = `Bienvenido al Sistema de Gestión de Servicios\n\n`;
  message += `Usuario: ${comercial} (${userEmail})\n\n`;
  
  message += `Instrucciones de uso:\n\n`;
  message += `1. "Actualizar datos desde XLSX": Importa y filtra los datos desde el archivo XLSX.\n\n`;
  message += `2. "Enviar seleccionados a emisión": Selecciona filas en la hoja "de enviar a emisión" y envía por email.\n\n`;
  message += `3. "Actualizar pendientes": Refresca la lista de servicios pendientes.\n\n`;
  message += `Puedes acceder a estas funciones desde el menú "Gestión de Servicios".`;
  
  ui.alert('Bienvenido', message, ui.ButtonSet.OK);
} */

/**
  Shows information about the current configuration
  Useful for debugging

function showConfigInfo() {
  try {
    const config = getRemoteConfig();
    const comercial = config.comercial;
    const userEmail = Session.getActiveUser().getEmail();
    
    let message = `Información de Configuración\n\n`;
    message += `Usuario: ${userEmail}\n`;
    message += `Comercial: ${comercial}\n\n`;
    message += `Hojas configuradas:\n`;
    message += `• Importación: ${config.sheetNames.import}\n`;
    message += `• Pendientes: ${config.sheetNames.pending}\n`;
    message += `• Log: ${config.sheetNames.log}\n\n`;
    message += `Email destinatario: ${config.emailRecipient}`;
    
    SpreadsheetApp.getUi().alert('Configuración', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'No se pudo cargar la configuración: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
 */