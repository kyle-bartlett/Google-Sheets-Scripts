function getImportRangeSheetIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetIds = [];
  var sheetsWithImportRange = [];

  // Process only the first 20 rows and 30 columns
  var numRows = 20;
  var numCols = 30;

  sheets.forEach(sheet => {
    var lastRow = Math.min(sheet.getLastRow(), numRows);
    var lastCol = Math.min(sheet.getLastColumn(), numCols);
    var range = sheet.getRange(1, 1, lastRow, lastCol);
    var formulas = range.getFormulas();

    Logger.log('Scanning sheet: ' + sheet.getName());

    formulas.forEach((row, rowIndex) => {
      row.forEach((cellFormula, colIndex) => {
        if (cellFormula.toLowerCase().includes("importrange")) {
          var cellA1Notation = range.getCell(rowIndex + 1, colIndex + 1).getA1Notation();
          Logger.log('Sheet: ' + sheet.getName() + ', Cell: ' + cellA1Notation + ', Formula: ' + cellFormula);

          // Modified regex pattern to accommodate different IMPORTRANGE formula formats
          var match = cellFormula.match(/(?:[\"\']https:\/\/docs\.google\.com\/spreadsheets\/d\/)([a-zA-Z0-9-_]+)/) || cellFormula.match(/(?:IMPORTRANGE\(\")([a-zA-Z0-9-_]+)/);
          
          if (match && match[1]) {
            sheetIds.push(match[1]);
            sheetsWithImportRange.push(sheet.getName());
            Logger.log('IMPORTRANGE found in sheet: ' + sheet.getName() + ', Cell: ' + cellA1Notation + ', Sheet ID: ' + match[1]);
          } else {
            Logger.log('IMPORTRANGE formula detected but no valid Sheet ID found in: ' + cellFormula);
          }
        }
      });
    });
  });

  Logger.log('Sheets with IMPORTRANGE identified: ' + JSON.stringify(sheetsWithImportRange));
  Logger.log('Unique Sheet IDs identified: ' + JSON.stringify(sheetIds));

  // Remove duplicates from the lists
  var uniqueSheetIds = [...new Set(sheetIds)];
  var uniqueSheetsWithImportRange = [...new Set(sheetsWithImportRange)];

  // Update "ImportRange Sheet IDs" sheet
  var sheetIdSheet = ss.getSheetByName('ImportRange Sheet IDs');
  if (!sheetIdSheet) {
    sheetIdSheet = ss.insertSheet('ImportRange Sheet IDs');
  }
  sheetIdSheet.clearContents();

  // Write sheet IDs and names
  if (uniqueSheetIds.length > 0) {
    var sheetIdArray = uniqueSheetIds.map(id => ['"' + id + '",']);
    sheetIdSheet.getRange(1, 1, sheetIdArray.length, 1).setValues(sheetIdArray);
  }

 if (uniqueSheetsWithImportRange.length > 0) {
    // Add quotes around each sheet name
    var sheetNamesString = uniqueSheetsWithImportRange.map(name => '"' + name + '"').join(", ");
    sheetIdSheet.getRange(1, 2).setValue(sheetNamesString);
  }
}
