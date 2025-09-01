function duplicateColumns() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Notes");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "Notes" not found');
    return;
  }
  
  // Define the range I:AH (columns 9-34, which is 26 columns)
  const sourceStartCol = 9;  // Column I
  const sourceEndCol = 34;   // Column AH
  const numColumns = 26;
  
  // Insert 26 new columns after column AH (position 35)
  sheet.insertColumns(sourceEndCol + 1, numColumns);
  
  // Get the last row with data to define our range
  const lastRow = sheet.getLastRow();
  
  // Define source range (I:AH) - all rows with data
  const sourceRange = sheet.getRange(1, sourceStartCol, lastRow, numColumns);
  
  // Define destination range (AI:BH) - columns 35-60 after insertion
  const destStartCol = sourceEndCol + 1; // Column AI (35)
  const destRange = sheet.getRange(1, destStartCol, lastRow, numColumns);
  
  // Copy values only (not formulas or formatting)
  const values = sourceRange.getValues();
  destRange.setValues(values);
  
  SpreadsheetApp.getUi().alert('Successfully duplicated columns I:AH to AI:BH as values only');
}