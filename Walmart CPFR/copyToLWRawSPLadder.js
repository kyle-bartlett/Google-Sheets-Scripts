function copyToLWRawSPLadder () {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source sheet (Raw_SP Ladder)
  const sourceSheet = spreadsheet.getSheetByName("Raw_SP Ladder");
  if (!sourceSheet) {
    throw new Error("Source sheet 'Raw_SP Ladder' not found");
  }
  
  // Get destination sheet (LW Raw_SP Ladder)
  const destSheet = spreadsheet.getSheetByName("LW Raw_SP Ladder");
  if (!destSheet) {
    throw new Error("Destination sheet 'LW Raw_SP Ladder' not found");
  }
  
  // Get all data from A1:AW (columns A through AW)
  const sourceRange = sourceSheet.getRange("A1:AW");
  const sourceValues = sourceRange.getValues();
  
  // Clear destination sheet first (optional - remove if you don't want this)
  destSheet.clear();
  
  // Paste values only to destination sheet starting at A1
  const destRange = destSheet.getRange(1, 1, sourceValues.length, sourceValues[0].length);
  destRange.setValues(sourceValues);
  
  // Optional: Show confirmation
  SpreadsheetApp.getUi().alert('Data copied successfully from Raw_SP Ladder to LW Raw_SP Ladder')}