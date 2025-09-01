function copyPasteTotalPipe() {
  var sourceSpreadsheetId = '1-VsBCoShVk106dSNali-4TuCIAanKnVpOW7n7r6U4b8';
  var sourceSheetName = 'Pipeline Overview';
  Logger.log('Starting the copyPasteTotalPipe function.');

  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Replace "YourTargetSheetName" with the actual name of your target sheet
  var targetSheetName = "Total Pipeline"; 
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

  if (!targetSheet) {
    Logger.log('"' + targetSheetName + '" sheet not found. Exiting the function.');
    return;
  }

  Logger.log('The target sheet is: ' + targetSheet.getName());

  try {
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    Logger.log('Accessing source sheet: ' + sourceSheetName);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

    if (sourceSheet === null) {
      throw new Error('Sheet not found: ' + sourceSheetName);
    }

    var sourceRange = sourceSheet.getDataRange();
    var valuesToCopy = sourceRange.getValues();

    // Adjust the targetRangeA1Notation if the starting point for pasting changes
    var targetRangeA1Notation = 'A2'; 
    targetSheet.getRange(targetRangeA1Notation)
               .offset(0, 0, sourceRange.getNumRows(), sourceRange.getNumColumns())
               .setValues(valuesToCopy);
    Logger.log('Values pasted successfully.');

    var currentDate = new Date();
    var timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy HH:mm:ss");

    var endRow = 1 + sourceRange.getNumRows(); 
    var endColumn = columnToLetter(columnToLetterToNumber('A') + sourceRange.getNumColumns() - 1);

    var finalLogMessage = "Pasted rows from 2 to " + endRow + 
                          " and columns from A to " + endColumn + 
                          " on " + formattedDate;
    Logger.log(finalLogMessage);
    targetSheet.getRange('G1').setValue(finalLogMessage);

  } catch (e) {
    var errorMessage = "Error: " + e.message;
    Logger.log(errorMessage);
    targetSheet.getRange('G1').setValue(errorMessage);
  }
}

function columnToLetter(columnNum) {
  let temp, letter = '';
  while (columnNum > 0) {
    temp = (columnNum - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNum = (columnNum - temp - 1) / 26;
  }
  return letter;
}

function columnToLetterToNumber(columnLetter) {
  let column = 0, length = columnLetter.length;
  for (let i = 0; i < length; i++) {
    column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}
