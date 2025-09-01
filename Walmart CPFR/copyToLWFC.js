function copyToLWFC () {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source sheet (ReportUpload)
  const sourceSheet = spreadsheet.getSheetByName("CW FC");
  if (!sourceSheet) {
    throw new Error("Source sheet 'CW FC' not found");
  }
  
  // Get destination sheet (LW FC)
  const destSheet = spreadsheet.getSheetByName("LW FC");
  if (!destSheet) {
    throw new Error("Destination sheet 'LW FC' not found");
  }
  
  // Get all data from A1:BT (columns A through BT)
  const sourceRange = sourceSheet.getRange("A1:BT");
  const sourceValues = sourceRange.getValues();
  
  // Clear destination sheet first (optional - remove if you don't want this)
  destSheet.clear();
  
  // Paste values only to destination sheet starting at A1
  const destRange = destSheet.getRange(1, 1, sourceValues.length, sourceValues[0].length);
  destRange.setValues(sourceValues);
  
  // Optional: Show confirmation
  SpreadsheetApp.getUi().alert('Data copied successfully from CW FC to LW FC')}