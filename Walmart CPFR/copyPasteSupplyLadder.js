function copyPasteSupplyLadder() {
  var sourceSpreadsheetId = '1K2yzFBkPgeb_2L5IfZ9wKHPY7fW2gkycRy_4RJAiwKg';
  var sourceSheetName = 'WM-Charging';

  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

  if (sourceSheet === null) {
    Logger.log('Sheet not found: ' + sourceSheetName);
    return; // Exit the function if the sheet is not found
  }

  var sourceRange = sourceSheet.getDataRange();
  var valuesToCopy = sourceRange.getValues();
  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = targetSpreadsheet.getActiveSheet();
  var targetRangeA1Notation = 'B2';

  targetSheet.getRange(targetRangeA1Notation)
              .offset(0, 0, sourceRange.getNumRows(), sourceRange.getNumColumns())
              .setValues(valuesToCopy);
}
