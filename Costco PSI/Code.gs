/**
 * Costco PSI Weekly Sales Update Automation
 * This script automates the process of pulling data from multiple sheets
 * and pasting it into specific cells in a target sheet.
 */

// Configuration object - modify these values for your specific setup
const CONFIG = {
  // Target sheet where data will be pasted
  TARGET_SHEET_NAME: 'Weekly Summary',
  
  // Source sheets and their data ranges
  SOURCE_SHEETS: [
    {
      name: 'Sales Data',
      ranges: [
        { source: 'A1:D10', target: 'B2:E11', description: 'Weekly sales figures' },
        { source: 'F1:H15', target: 'G2:I16', description: 'Product performance' }
      ]
    },
    {
      name: 'Inventory Data',
      ranges: [
        { source: 'A1:C20', target: 'K2:M21', description: 'Stock levels' }
      ]
    },
    {
      name: 'Customer Data',
      ranges: [
        { source: 'B1:E25', target: 'O2:R26', description: 'Customer metrics' }
      ]
    }
  ],
  
  // Date formatting for the update
  DATE_FORMAT: 'MM/dd/yyyy',
  
  // Logging options
  ENABLE_LOGGING: true
};

/**
 * Main function to run the weekly update automation
 */
function runWeeklyUpdate() {
  try {
    if (CONFIG.ENABLE_LOGGING) {
      console.log('Starting weekly sales update automation...');
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = getOrCreateSheet(spreadsheet, CONFIG.TARGET_SHEET_NAME);
    
    // Add timestamp header
    addTimestampHeader(targetSheet);
    
    // Process each source sheet
    CONFIG.SOURCE_SHEETS.forEach(sourceConfig => {
      processSourceSheet(spreadsheet, sourceConfig, targetSheet);
    });
    
    // Format the target sheet
    formatTargetSheet(targetSheet);
    
    if (CONFIG.ENABLE_LOGGING) {
      console.log('Weekly update completed successfully!');
    }
    
    // Show success message
    SpreadsheetApp.getUi().alert('Weekly Update Complete', 
      'All data has been successfully copied and pasted to the target sheet.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
      
  } catch (error) {
    console.error('Error in weekly update:', error);
    SpreadsheetApp.getUi().alert('Error', 
      'An error occurred during the update: ' + error.message, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Process a single source sheet and copy its data to the target
 */
function processSourceSheet(spreadsheet, sourceConfig, targetSheet) {
  try {
    const sourceSheet = spreadsheet.getSheetByName(sourceConfig.name);
    if (!sourceSheet) {
      throw new Error(`Source sheet "${sourceConfig.name}" not found`);
    }
    
    if (CONFIG.ENABLE_LOGGING) {
      console.log(`Processing source sheet: ${sourceConfig.name}`);
    }
    
    // Process each range for this source sheet
    sourceConfig.ranges.forEach(rangeConfig => {
      copyRangeData(sourceSheet, rangeConfig, targetSheet);
    });
    
  } catch (error) {
    console.error(`Error processing source sheet ${sourceConfig.name}:`, error);
    throw error;
  }
}

/**
 * Copy data from a specific range in source sheet to target sheet
 */
function copyRangeData(sourceSheet, rangeConfig, targetSheet) {
  try {
    const sourceRange = sourceSheet.getRange(rangeConfig.source);
    const targetRange = targetSheet.getRange(rangeConfig.target);
    
    // Get the data from source
    const data = sourceRange.getValues();
    
    // Paste to target
    targetRange.setValues(data);
    
    if (CONFIG.ENABLE_LOGGING) {
      console.log(`Copied ${rangeConfig.description} from ${rangeConfig.source} to ${rangeConfig.target}`);
    }
    
  } catch (error) {
    console.error(`Error copying range ${rangeConfig.source}:`, error);
    throw error;
  }
}

/**
 * Get or create the target sheet if it doesn't exist
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    if (CONFIG.ENABLE_LOGGING) {
      console.log(`Created new sheet: ${sheetName}`);
    }
  }
  
  return sheet;
}

/**
 * Add timestamp header to the target sheet
 */
function addTimestampHeader(targetSheet) {
  const timestamp = new Date();
  const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), CONFIG.DATE_FORMAT);
  
  targetSheet.getRange('A1').setValue('Weekly Sales Update - ' + formattedDate);
  targetSheet.getRange('A1').setFontWeight('bold');
  targetSheet.getRange('A1').setFontSize(14);
  
  if (CONFIG.ENABLE_LOGGING) {
    console.log(`Added timestamp header: ${formattedDate}`);
  }
}

/**
 * Format the target sheet for better readability
 */
function formatTargetSheet(targetSheet) {
  try {
    // Auto-resize columns
    targetSheet.autoResizeColumns(1, targetSheet.getLastColumn());
    
    // Add borders to data ranges
    const lastRow = targetSheet.getLastColumn();
    const lastCol = targetSheet.getLastColumn();
    
    if (lastRow > 1 && lastCol > 1) {
      const dataRange = targetSheet.getRange(2, 1, lastRow - 1, lastCol);
      dataRange.setBorder(true, true, true, true, true, true);
    }
    
    // Format header row
    if (lastRow > 1) {
      const headerRow = targetSheet.getRange(1, 1, 1, lastCol);
      headerRow.setFontWeight('bold');
      headerRow.setBackground('#f3f3f3');
    }
    
    if (CONFIG.ENABLE_LOGGING) {
      console.log('Target sheet formatting completed');
    }
    
  } catch (error) {
    console.error('Error formatting target sheet:', error);
  }
}

/**
 * Test function to verify the configuration
 */
function testConfiguration() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if all source sheets exist
    const missingSheets = [];
    CONFIG.SOURCE_SHEETS.forEach(sourceConfig => {
      const sheet = spreadsheet.getSheetByName(sourceConfig.name);
      if (!sheet) {
        missingSheets.push(sourceConfig.name);
      }
    });
    
    if (missingSheets.length > 0) {
      SpreadsheetApp.getUi().alert('Configuration Error', 
        `The following source sheets are missing:\n${missingSheets.join('\n')}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    SpreadsheetApp.getUi().alert('Configuration Valid', 
      'All source sheets are present and configuration is valid.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
      
  } catch (error) {
    console.error('Error testing configuration:', error);
    SpreadsheetApp.getUi().alert('Error', 
      'Error testing configuration: ' + error.message, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Create a custom menu in the spreadsheet
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Costco PSI Automation')
    .addItem('Run Weekly Update', 'runWeeklyUpdate')
    .addItem('Test Configuration', 'testConfiguration')
    .addSeparator()
    .addItem('View Configuration', 'showConfiguration')
    .addToUi();
}

/**
 * Show current configuration
 */
function showConfiguration() {
  const configText = JSON.stringify(CONFIG, null, 2);
  SpreadsheetApp.getUi().alert('Current Configuration', 
    `Configuration:\n\n${configText}`, 
    SpreadsheetApp.getUi().ButtonSet.OK);
}
