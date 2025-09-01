function updateDailyInv() {
  // File IDs and configuration
  const DATA_FILE_ID = '1265dP-LCedwpxZzluG-KlCwd65lMPDfpBJD2mAc2Nq4';
  const HOST_FILE_ID = '1SYIFq0_AN9ziuPe4N59V1tDdHylhtbpSLUwVZMvyCsU';
  const DATA_TAB_NAME = 'Summary tab';
  const HOST_TAB_NAME = 'Daily Inv';
  const FILTER_VALUE = 'Walmart stores';
  
  try {
    // Open spreadsheets
    const dataSpreadsheet = SpreadsheetApp.openById(DATA_FILE_ID);
    const hostSpreadsheet = SpreadsheetApp.openById(HOST_FILE_ID);
    
    // Get sheets
    const dataSheet = dataSpreadsheet.getSheetByName(DATA_TAB_NAME);
    const hostSheet = hostSpreadsheet.getSheetByName(HOST_TAB_NAME);
    
    if (!dataSheet) {
      throw new Error(`Sheet "${DATA_TAB_NAME}" not found in data file`);
    }
    if (!hostSheet) {
      throw new Error(`Sheet "${HOST_TAB_NAME}" not found in host file`);
    }
    
    // Get all data from A1:BA (up to last row with data)
    const lastRow = dataSheet.getLastRow();
    const lastCol = 53; // BA is column 53
    
    if (lastRow < 3) {
      console.log('Not enough data rows (need at least 3 for headers in row 3)');
      return;
    }
    
    const allData = dataSheet.getRange(1, 1, lastRow, lastCol).getValues();
    
    // Get headers from row 3 (index 2)
    const headers = allData[2];
    
    // Find the "Type" column index
    const typeColumnIndex = headers.indexOf('Type');
    if (typeColumnIndex === -1) {
      throw new Error('Column "Type" not found in row 3');
    }
    
    // Filter data: keep rows 1-3 (headers) + rows where Type = "Walmart stores"
    const filteredData = [];
    
    // Add rows 1-3 (headers)
    filteredData.push(allData[0], allData[1], allData[2]);
    
    // Add filtered data rows (starting from row 4, index 3)
    for (let i = 3; i < allData.length; i++) {
      if (allData[i][typeColumnIndex] === FILTER_VALUE) {
        filteredData.push(allData[i]);
      }
    }
    
    // Clear the destination range first
    hostSheet.clear();
    
    // Paste filtered data
    if (filteredData.length > 0) {
      const pasteRange = hostSheet.getRange(1, 1, filteredData.length, lastCol);
      pasteRange.setValues(filteredData);
      
      console.log(`Successfully copied ${filteredData.length - 3} filtered rows (plus 3 header rows)`);
    } else {
      console.log('No data to copy after filtering');
    }
    
  } catch (error) {
    console.error('Error in copyFilteredData:', error.toString());
    
    // Optional: Send email notification on error
    // MailApp.sendEmail('your-email@domain.com', 'Script Error', error.toString());
  }
}

function setupHourlyTrigger() {
  // Delete existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'copyFilteredData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new hourly trigger
  ScriptApp.newTrigger('copyFilteredData')
    .timeBased()
    .everyHours(1)
    .create();
  
  console.log('Hourly trigger set up successfully');
}

function testCopy() {
  // Run this once to test the copy function
  copyFilteredData();
}