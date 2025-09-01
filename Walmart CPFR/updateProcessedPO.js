function updateProcessedPO() {
  // File IDs and configuration
  const DATA_FILE_ID = '1Fe6EU1s-uMFosXI2NHmN4pY-WvYO4skLXu5T-yxwC0Q';
  const HOST_FILE_ID = '1SYIFq0_AN9ziuPe4N59V1tDdHylhtbpSLUwVZMvyCsU';
  const DATA_TAB_NAME = 'Order Entry(CW)pivot for Daniel';
  const HOST_TAB_NAME = 'Processed PO';
  
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
    
    // Get all data from A1:C (up to last row with data)
    const lastRow = dataSheet.getLastRow();
    const lastCol = 3; // Column C
    
    if (lastRow < 1) {
      console.log('No data to copy');
      return;
    }
    
    const allData = dataSheet.getRange(1, 1, lastRow, lastCol).getValues();
    
    // Clear only columns A-C (preserve everything from D onward)
    const lastRowInHost = hostSheet.getLastRow();
    if (lastRowInHost > 0) {
      const clearRange = hostSheet.getRange(1, 1, lastRowInHost, 3); // Only clear A-C
      clearRange.clear();
    }
    
    // Paste all data only in columns A-C
    if (allData.length > 0) {
      const pasteRange = hostSheet.getRange(1, 1, allData.length, lastCol);
      pasteRange.setValues(allData);
      console.log(`Successfully copied ${allData.length} rows to columns A-C only`);
    }
    
  } catch (error) {
    console.error('Error in updateProcessedPO:', error.toString());
    // Send email notification on error
    MailApp.sendEmail('kyle.bartlett@Anker.com', 'Script Error', error.toString());
  }
}

function setup12hrTrigger() {
  // Delete existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'updateProcessedPO') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new 12-hour trigger
  ScriptApp.newTrigger('updateProcessedPO')
    .timeBased()
    .everyHours(12)
    .create();
    
  console.log('12-hour trigger set up successfully');
}

function testCopy() {
  // Run this once to test the copy function
  updateProcessedPO();
}