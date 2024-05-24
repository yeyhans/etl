// El script automatiza la recopilación y consolidación de datos de múltiples hojas de cálculo en una única hoja de dashboard, permitiendo una fácil visualización y análisis de la información más relevante de las boletas o presupuestos. Este enfoque optimiza el tiempo y minimiza errores al centralizar la información clave en un solo lugar.
function actualizarDashboard() {
  var spreadsheet = SpreadsheetApp.openById("**ID**"); // Reemplaza con el ID de tu Google Sheets
  var dashboard = SpreadsheetApp.openById("**ID**"); // Reemplaza con el ID de tu dashboard
  var sheets = spreadsheet.getSheets();
  var dashboardSheet = dashboard.getSheetByName("Dashboard"); // Reemplaza con el nombre de tu hoja de dashboard

  var allData = []; // Almacenar todos los datos en una sola matriz

  // Bucle a través de todas las hojas del Google Sheets
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName(); // Obtener el nombre de la hoja actual
    if (sheetName !== "Resumen" && sheetName !== "Inventario" && sheetName !== "Planilla") { // Verificar si el nombre de la hoja no es "Resumen" ni "Inventario" ni "Planilla"
      var j5Data = sheet.getRange("J5").getValue(); // Obtener valor de J5 (Fecha de producción)
      var j25Data = sheet.getRange("J25").getValue(); // Obtener valor de J25 (Sub-Total Boleta Factura)
      var h22Data = sheet.getRange("H22").getValue(); // Obtener valor de H22 (Total Boleta Factura)
      var g4Data = sheet.getRange("G4:H4").getValues(); // Obtener valores de G4 como matriz (contacto nombre)
      var g5Data = sheet.getRange("G5:H5").getValues(); // Obtener valores de G5 como matriz (email)
      var c4Data = sheet.getRange("C4").getValue(); // Obtener valor c4 (Nº boleta)
      var c5Data = sheet.getRange("C5").getValue(); // Obtener valor c5 (Fecha creación)

      var c4Link = '=HYPERLINK(\'N' + c4Data + '\'!C4)'; // Enlaza a la hoja de calculo 


      // Agregar todos los datos a la matriz
      allData.push([c5Data, c4Link, g4Data, g5Data, j5Data, j25Data, h22Data]);
    }
  }

  // Limpiar la hoja de dashboard antes de actualizar
  var lastRow = dashboardSheet.getLastRow();
  if (lastRow > 1) {
    dashboardSheet.deleteRows(2, lastRow - 1);
  }

  // Agregar todos los datos al dashboard
  if (allData.length > 0) {
    dashboardSheet.getRange(2, 1, allData.length, allData[0].length).setValues(allData);
  }
}
