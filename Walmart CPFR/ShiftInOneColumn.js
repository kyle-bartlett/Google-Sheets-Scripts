// Script Summary:
// This script shifts cell contents to the left by one column in rows where specific text strings are found.
// It's designed to work on the active sheet of a Google Spreadsheet.
// The script processes multiple search strings and handles each row where these strings are found.

var sheet = SpreadsheetApp.getActiveSheet();

// Function to find the target column where the text "Shift" is found starting from row 2
function findTargetCol() {
  var startRow = 2;  // Start searching from row 2
  var range = sheet.getRange(startRow, 1, 1, sheet.getLastColumn());
  var values = range.getValues()[0]; // Get the first (and only) row of values

  var colToShift;
  values.forEach((value, index) => {
    if (value === "Shift") {
      colToShift = index + 1; // +1 because array indexes start at 0, but spreadsheet columns start at 1
      return;
    }
  });

  return colToShift || null; // Return the found column or null if not found
}



// Function to process shifting for multiple search texts
function _shiftAll(searchTexts) {
  // Ensure searchTexts is an array
  if (!Array.isArray(searchTexts)) {
    searchTexts = [searchTexts];
  }

  Logger.log("Search texts: " + searchTexts.join(", "));

  searchTexts.forEach(searchTxt => {
    var tf = sheet.createTextFinder(searchTxt);
    var cellsFound = tf.findAll();
    Logger.log("Found " + cellsFound.length + " cells with text: " + searchTxt);

    var rowsToProcess = cellsFound.filter(v => v.getValue() == searchTxt).map(v => v.getRow());

    if (rowsToProcess.length == 0) {
      Logger.log("No rows found with the text: " + searchTxt);
      return;
    }

    var col = findTargetCol();
    rowsToProcess.forEach((row, i) => {
      if (i > 0 && row == rowsToProcess[i - 1]) return; // Skip duplicate rows
      shiftLeftByOne(row, col);
    });
  });
}

// Function to initiate the shift process for multiple texts
function shiftAll() {
  _shiftAll(["Sellin FC (Uncon.)", "Sellin FC (Con.)", "Seasonality index"]);
}

// Function to shift cells left by one column
function shiftLeftByOne(row, col) {
  var range = sheet.getRange(row, col + 1, 1, sheet.getLastColumn() - col);
  var dst = sheet.getRange(row, col, 1, sheet.getLastColumn() - col);
  range.copyTo(dst);
}

// Run the script
function run() {
  shiftAll();
}
