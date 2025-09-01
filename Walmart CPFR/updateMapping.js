function updateMapping() {
  // Sheet IDs extracted from URLs
  const HOST_SHEET_ID = '1SYIFq0_AN9ziuPe4N59V1tDdHylhtbpSLUwVZMvyCsU';
  const DATA_SHEET_ID = '11VOdGH66QwP33g7jugReb2-ZBxs3YJU6sKYqauwsphs';
  
  // Configuration
  const DATA_TAB_NAME = 'SKU mapping';
  const DATA_RANGE = 'A1:AE';
  const HOST_TAB_NAME = 'Mapping';
  const START_ROW = 15; // Clear and paste starting from row 15
  const FILTER_COLUMN = 'B'; // IPMT column
  const FILTER_VALUE = 'Mobile Charging IPMT';
  
  try {
    // Open the data source sheet
    const dataSheet = SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName(DATA_TAB_NAME);
    if (!dataSheet) {
      throw new Error(`Tab "${DATA_TAB_NAME}" not found in data sheet`);
    }
    
    // Get all data from the specified range
    const allData = dataSheet.getRange(DATA_RANGE).getValues();
    
    if (allData.length === 0) {
      throw new Error('No data found in the specified range');
    }
    
    // Find the header row and column B index
    const headers = allData[0];
    const columnBIndex = headers.findIndex(header => header === 'IPMT' || header.toString().toLowerCase().includes('ipmt'));
    
    if (columnBIndex === -1) {
      throw new Error('IPMT column not found in headers');
    }
    
    console.log(`Found IPMT column at index: ${columnBIndex}`);
    
    // Filter data: keep header row + rows where column B matches filter value
    const filteredData = allData.filter((row, index) => {
      // Keep header row (index 0)
      if (index === 0) return true;
      
      // Keep rows where column B matches the filter value
      return row[columnBIndex] === FILTER_VALUE;
    });
    
    console.log(`Filtered from ${allData.length} to ${filteredData.length} rows`);
    
    if (filteredData.length <= 1) {
      console.log('Warning: Only header row found after filtering. No matching data.');
    }
    
    // Open the host sheet and get the specific tab
    const hostSpreadsheet = SpreadsheetApp.openById(HOST_SHEET_ID);
    const hostSheet = hostSpreadsheet.getSheetByName(HOST_TAB_NAME);
    
    if (!hostSheet) {
      throw new Error(`Tab "${HOST_TAB_NAME}" not found in host sheet`);
    }
    
    // Clear everything from row 15 onwards (preserve rows 1-14)
    const lastRow = hostSheet.getLastRow();
    const lastColumn = hostSheet.getLastColumn();
    
    if (lastRow >= START_ROW) {
      const clearRange = hostSheet.getRange(START_ROW, 1, lastRow - START_ROW + 1, lastColumn);
      clearRange.clear();
    }
    
    // Paste the filtered data starting at row 15
    if (filteredData.length > 0) {
      const targetRange = hostSheet.getRange(START_ROW, 1, filteredData.length, filteredData[0].length);
      targetRange.setValues(filteredData);
      
      console.log(`Successfully pasted ${filteredData.length} rows starting at row ${START_ROW}`);
    }
    
    return {
      success: true,
      message: `Filtered and pasted ${filteredData.length} rows (including header) to host sheet`,
      originalRows: allData.length,
      filteredRows: filteredData.length
    };
    
  } catch (error) {
    console.error('Error in updateMapping:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

// Optional: Function to run this automatically on a schedule
function createTrigger() {
  // Delete existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'updateMapping') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger - runs every hour (adjust as needed)
  ScriptApp.newTrigger('updateMapping')
    .timeBased()
    .everyHours(1)
    .create();
    
  console.log('Trigger created to run every hour');
}

// Test function to verify the filtering works
function testFilter() {
  const result = updateMapping();
  console.log('Test result:', result);
}