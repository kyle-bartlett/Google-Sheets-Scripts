function copyToLWReportUpload() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source sheet (ReportUpload)
  const sourceSheet = spreadsheet.getSheetByName("ReportUpload");
  if (!sourceSheet) {
    throw new Error("Source sheet 'ReportUpload' not found");
  }
  
  // Get destination sheet (LW ReportUpload)
  const destSheet = spreadsheet.getSheetByName("LW ReportUpload");
  if (!destSheet) {
    throw new Error("Destination sheet 'LW ReportUpload' not found");
  }
  
  // Get all data from A1:JT (columns A through JT)
  const sourceRange = sourceSheet.getRange("A1:JT");
  const sourceValues = sourceRange.getValues();
  
  // Clear destination sheet first (optional - remove if you don't want this)
  destSheet.clear();
  
  // Paste values only to destination sheet starting at A1
  const destRange = destSheet.getRange(1, 1, sourceValues.length, sourceValues[0].length);
  destRange.setValues(sourceValues);
  
  // Optional: Show confirmation
  SpreadsheetApp.getUi().alert('Data copied successfully from ReportUpload to LW ReportUpload')}