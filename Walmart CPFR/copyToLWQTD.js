function copyToLWQTD() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source sheet (CW QTD)
  const sourceSheet = spreadsheet.getSheetByName("CW QTD");
  if (!sourceSheet) {
    throw new Error("Source sheet 'CW QTD' not found");
  }
  
  // Get destination sheet (LW QTD)
  const destSheet = spreadsheet.getSheetByName("LW QTD");
  if (!destSheet) {
    throw new Error("Destination sheet 'LW QTD' not found");
  }
  
  // Get all data from A1:AU (columns A through AU)
  const sourceRange = sourceSheet.getRange("A1:AU");
  const sourceValues = sourceRange.getValues();
  
  // Clear destination sheet first (optional - remove if you don't want this)
  destSheet.clear();
  
  // Paste values only to destination sheet starting at A1
  const destRange = destSheet.getRange(1, 1, sourceValues.length, sourceValues[0].length);
  destRange.setValues(sourceValues);
  
  // Optional: Show confirmation
  SpreadsheetApp.getUi().alert('Data copied successfully from CW QTD to LW QTD')}