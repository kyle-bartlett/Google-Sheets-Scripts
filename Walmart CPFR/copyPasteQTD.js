function copyPasteQTD() {
  var sourceSpreadsheetId = '1bQxaNJwspIYmHGhRemKHOp0Y21fCu9Y-s3nLyFqxbiU'; // ID of the source spreadsheet
  var sourceSheetName = 'SKU level WoW delta -25Q3'; // Name of the source sheet
  Logger.log('Starting the copyPasteQTD function.');

  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet

  // Specify the target sheet by name
  var targetSheetName = "CW QTD"; // Replace with your actual target sheet name
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName); // Get the target sheet by name

  // Check if the target sheet exists
  if (!targetSheet) {
    Logger.log('"' + targetSheetName + '" sheet not found. Exiting the function.');
    return; // Exit the function if the sheet does not exist
  }

  Logger.log('The target sheet is: ' + targetSheet.getName());

  try {
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId); // Open the source spreadsheet by ID
    Logger.log('Accessing source sheet: ' + sourceSheetName);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName); // Get the source sheet by name

    // Check if the source sheet exists
    if (sourceSheet === null) throw new Error('Sheet not found: ' + sourceSheetName);

    var manualRange = sourceSheet.getRange('A1:AU3000'); // Define the range to copy from the source sheet
    var valuesToCopy = manualRange.getValues(); // Get the values from the defined range

    Logger.log('Clearing content in the target sheet range A3:AU3002.');
    // Clear the target sheet contents from A3 to AU3002
    targetSheet.getRange('A3:AU3002').clearContent();

    Logger.log('Preparing to paste values into "' + targetSheetName + '" sheet starting at A3.');
    // Paste the copied values into the target sheet starting at A3
    targetSheet.getRange('A3').offset(0, 0, 3000, manualRange.getNumColumns()).setValues(valuesToCopy);
    Logger.log('Values pasted successfully.');

    // Get the current date and time for the log message
    var currentDate = new Date();
    var timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy HH:mm:ss");

    // Calculate the end row and end column for logging purposes
    var endRow = 2 + manualRange.getNumRows() - 1; // Adjust based on your start row and the number of rows pasted
    var endColumn = columnToLetter(columnToLetterToNumber('A') + manualRange.getNumColumns() - 1);

    // Construct and log the final message with details
    var finalLogMessage = "Pasted rows from 3 to " + endRow +
                          " and columns from A to " + endColumn +
                          " on " + formattedDate;
    Logger.log(finalLogMessage);
    targetSheet.getRange('D1').setValue(finalLogMessage); // Set the final log message in the target sheet

  } catch (e) {
    // Log and display the error message with a timestamp
    var errorMessage = "Error: " + e.message;
    var currentDate = new Date();
    var timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy HH:mm:ss");
    Logger.log(formattedDate + " - " + errorMessage);
    targetSheet.getRange('D1').setValue(formattedDate + " - " + errorMessage); // Display the error message in the target sheet
  }
}

// Function to convert a column number to a letter
function columnToLetter(columnNum) {
  let temp, letter = '';
  while (columnNum > 0) {
    temp = (columnNum - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNum = (columnNum - temp - 1) / 26;
  }
  return letter;
}

// Function to convert a column letter to a number
function columnToLetterToNumber(columnLetter) {
  let column = 0, length = columnLetter.length;
  for (let i = 0; i < length; i++) {
    column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}