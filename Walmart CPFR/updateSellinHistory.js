function updateSellinHistory() {
  var sourceSpreadsheetId = '18B7eX7p_fQXyDXi_lwdf13gFNyJOjGWMX9nQB7Ak-xE';
  var sourceSheetName = 'WM Pivot Table';
  Logger.log('Starting the updateSellinHistory function.');

  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Target a specific sheet by name for pasting the values
  var targetSheetName = "Sellin History"; // Update this to your actual target sheet name
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

  // Verify that the target sheet exists
  if (!targetSheet) {
    Logger.log('"' + targetSheetName + '" sheet not found. Exiting the function.');
    return; // Exit the function if the sheet does not exist
  }

  Logger.log('The target sheet is: ' + targetSheet.getName());

  try {
    Logger.log('Attempting to open source spreadsheet.');
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    Logger.log('Accessing source sheet: ' + sourceSheetName);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

    if (sourceSheet === null) {
      throw new Error('Sheet not found: ' + sourceSheetName);
    }

    Logger.log('Retrieving data range from source sheet.');
    var sourceRange = sourceSheet.getDataRange();
    var valuesToCopy = sourceRange.getValues();

    Logger.log('Preparing to paste values into "' + targetSheetName + '" sheet.');
    var targetRangeA1Notation = 'B2'; // Adjust as needed for this script's target start cell

    // Execute the paste operation
    targetSheet.getRange(targetRangeA1Notation)
                .offset(0, 0, sourceRange.getNumRows(), sourceRange.getNumColumns())
                .setValues(valuesToCopy);
    Logger.log('Values pasted successfully.');

    // Get the current date and time for the log message
    var currentDate = new Date();
    var timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy HH:mm:ss");

    // Construct and log the final message with details
    var endRow = 1 + sourceRange.getNumRows(); // Adjusted for clarity
    var endColumn = columnToLetter(1 + sourceRange.getNumColumns()); // Assuming B2 start, hence +1
    var finalLogMessage = "Pasted rows from 2 to " + endRow + 
                          " and columns from B to " + endColumn + 
                          " on " + formattedDate;
    Logger.log(finalLogMessage);
    targetSheet.getRange('B1').setValue(finalLogMessage); // Set the final log message in the target sheet

    Logger.log('Last row and column highlighted in yellow.');
  } catch (e) {
    var errorMessage = "Error: " + e.message;
    Logger.log(errorMessage);
    targetSheet.getRange('B1').setValue(errorMessage); // Display the error message in the target sheet
  }
}
