// AÑADE UN MENÚ PERSONALIZADO AL ABRIR EL DOCUMENTO
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Indexa Capital')
    .addItem('1. Configurar Credenciales', 'configurarCredenciales')
    .addSeparator()
    .addItem('2. Actualizar Dashboard', 'actualizarDatosIndexa')
    .addToUi();
}

/**
 * Pide al usuario el token de la API y el ID de la cuenta y los guarda de forma segura.
 */
function configurarCredenciales() {
  const ui = SpreadsheetApp.getUi();
  
  const tokenResponse = ui.prompt('Configuración', 'Introduce tu Token de la API de Indexa Capital:', ui.ButtonSet.OK_CANCEL);
  if (tokenResponse.getSelectedButton() == ui.Button.OK && tokenResponse.getResponseText().length > 0) {
    PropertiesService.getScriptProperties().setProperty('INDEXA_TOKEN', tokenResponse.getResponseText());
  } else {
    ui.alert('No se guardó el token.');
    return;
  }
  
  const accountResponse = ui.prompt('Configuración', 'Introduce tu ID de cuenta (Ej: PD123ABC):', ui.ButtonSet.OK_CANCEL);
    if (accountResponse.getSelectedButton() == ui.Button.OK && accountResponse.getResponseText().length > 0) {
    PropertiesService.getScriptProperties().setProperty('INDEXA_ACCOUNT', accountResponse.getResponseText());
    ui.alert('¡Credenciales guardadas con éxito! Ya puedes actualizar los datos.');
  } else {
    ui.alert('No se guardó el ID de cuenta.');
  }
}

/**
 * Función principal para obtener datos de Indexa, procesarlos y crear un dashboard.
 */
/**
 * Función principal para obtener datos de Indexa, procesarlos y crear un dashboard.
 * VERSIÓN CORREGIDA para evitar el error de formato de número.
 */
function actualizarDatosIndexa() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  const token = scriptProperties.getProperty('INDEXA_TOKEN');
  const account = scriptProperties.getProperty('INDEXA_ACCOUNT');

  if (!token || !account) {
    ui.alert('Por favor, configura tus credenciales primero desde el menú "Indexa Capital > 1. Configurar Credenciales".');
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Dashboard Indexa';
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  try {
    ui.showSidebar(HtmlService.createHtmlOutput('<p>Actualizando datos desde Indexa Capital, por favor espera...</p>').setTitle('Cargando...'));

    // 1. OBTENER DATOS DE LA CARTERA
    const portfolioResponse = apiCall("https://api.indexacapital.com/accounts/" + account + "/portfolio", token);
    const data = JSON.parse(portfolioResponse.getContentText());

    // 2. EXTRAER Y CALCULAR DATOS DE RESUMEN
    const summary = data.portfolio;
    const totalValue = summary.total_amount;
    const investedValue = summary.instruments_amount;
    const cashValue = summary.cash_amount;
    const costBasis = summary.instruments_cost;
    const totalProfitLoss = investedValue - costBasis;
    const totalProfitLossPct = costBasis > 0 ? (totalProfitLoss / costBasis) : 0;
    const lastUpdateDate = new Date(summary.date);

    const summaryData = [
      ['Resumen de la Cartera', `Última actualización: ${lastUpdateDate.toLocaleDateString('es-ES')}`],
      ['Valor Total de la Cartera', totalValue],
      ['Total Invertido (Fondos)', investedValue],
      ['Efectivo (Cash)', cashValue],
      ['Ganancia / Pérdida Total (€)', totalProfitLoss],
      ['Ganancia / Pérdida Total (%)', totalProfitLossPct]
    ];
    
    // 3. PROCESAR DATOS DE DESGLOSE POR ACTIVO
    const positions = data.instrument_accounts[0].positions;
    const breakdownHeader = [
        "Fondo de Inversión", "Clase de Activo", "ISIN", 
        "Valor Actual (€)", "Coste (€)", "G/P (€)", "G/P (%)", "Peso (%)"
    ];
    const breakdownData = positions.map(pos => {
      const fundProfitLoss = pos.amount - pos.cost_amount;
      const fundProfitLossPct = pos.cost_amount > 0 ? (fundProfitLoss / pos.cost_amount) : 0;
      const weightPct = totalValue > 0 ? (pos.amount / totalValue) : 0;
      return [
        pos.instrument.name,
        pos.instrument.asset_class_description.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()),
        pos.instrument.isin_code,
        pos.amount,
        pos.cost_amount,
        fundProfitLoss,
        fundProfitLossPct,
        weightPct
      ];
    });

    // 4. ESCRIBIR DATOS EN LA HOJA
    sheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);
    const breakdownHeaderRow = summaryData.length + 2;
    sheet.getRange(breakdownHeaderRow, 1, 1, breakdownHeader.length).setValues([breakdownHeader]);
    if (breakdownData.length > 0) {
      sheet.getRange(breakdownHeaderRow + 1, 1, breakdownData.length, breakdownData[0].length).setValues(breakdownData);
    }
    
    // 5. APLICAR FORMATO
    // Formato de resumen
    sheet.getRange(1, 1, summaryData.length, 1).setFontWeight('bold');
    sheet.getRange(2, 2, 4, 1).setNumberFormat('€#,##0.00');
    sheet.getRange(6, 2).setNumberFormat('0.00%');
    
    // Formato de desglose (LA PARTE CORREGIDA)
    sheet.getRange(breakdownHeaderRow, 1, 1, breakdownHeader.length).setFontWeight('bold');
    if (breakdownData.length > 0) {
      const firstDataRow = breakdownHeaderRow + 1;
      const numDataRows = breakdownData.length;
      // Aplicar formato de moneda a las columnas 4, 5 y 6 (Valor, Coste, G/P €)
      sheet.getRange(firstDataRow, 4, numDataRows, 3).setNumberFormat('€#,##0.00');
      // Aplicar formato de porcentaje a las columnas 7 y 8 (G/P %, Peso %)
      sheet.getRange(firstDataRow, 7, numDataRows, 2).setNumberFormat('0.00%');
    }

    // Conditional Formatting para Ganancias/Pérdidas
    if (breakdownData.length > 0) {
      const plEuroRange = sheet.getRange(breakdownHeaderRow + 1, 6, breakdownData.length, 1);
      const plPctRange = sheet.getRange(breakdownHeaderRow + 1, 7, breakdownData.length, 1);
      
      const redRule = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground("#f4cccc").setBold(true).build();
      const greenRule = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground("#d9ead3").build();
      
      plEuroRange.addConditionalFormatRule(redRule);
      plEuroRange.addConditionalFormatRule(greenRule);
      plPctRange.addConditionalFormatRule(redRule);
      plPctRange.addConditionalFormatRule(greenRule);
    }
    
    // Auto-ajustar columnas y congelar filas
    sheet.autoResizeColumns(1, breakdownHeader.length);
    sheet.setFrozenRows(breakdownHeaderRow);
    
    ui.showSidebar(HtmlService.createHtmlOutput('<p>¡Dashboard actualizado con éxito!</p>').setTitle('Completado'));
    Utilities.sleep(4000);
    ui.showSidebar(HtmlService.createHtmlOutput(''));

  } catch (e) {
    Logger.log(e.toString());
    ui.alert('Ha ocurrido un error al crear el dashboard: ' + e.toString());
  }
}
/**
 * Función genérica para realizar llamadas a la API de Indexa.
 */
function apiCall(url, token) {
  const options = {
    "method": "GET",
    "headers": {
      "X-Auth-Token": token,
      "Content-Type": "application/json",
      "Accept": "application/json"
    },
    "muteHttpExceptions": true
  };
  
  const result = UrlFetchApp.fetch(url, options);
  
  if (result.getResponseCode() !== 200) {
    throw new Error(`Error en la API. Código: ${result.getResponseCode()}. Respuesta: ${result.getContentText()}`);
  }
  
  return result;
}
