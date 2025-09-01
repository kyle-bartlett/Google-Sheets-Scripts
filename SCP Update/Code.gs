// Google Apps Script for Two-Way Live Editing Between Tabs
// Report ID: 1I-O6PF3LyK0CQ3SrU2ygmd1uxOVNE920v_TpmLUvG6k

// Configuration
const SPREADSHEET_ID = "1I-O6PF3LyK0CQ3SrU2ygmd1uxOVNE920v_TpmLUvG6k";
const ALL_SKU_TAB_NAME = "All SKU Rollup WoW - FY'2025";
const TOP_SKU_TAB_NAME = "TOP SKU Level Summary";

// Column configurations
const ALL_SKU_EDIT_COL = "T"; // Column T for editing
const ALL_SKU_HELPER_COL = "W"; // Column W for helper lookup
const TOP_SKU_EDIT_COL = "T"; // Column T for editing  
const TOP_SKU_HELPER_COL = "A"; // Column A for helper lookup

// Global variables to prevent infinite loops
let isUpdating = false;
let lastUpdatedValue = "";
let lastUpdatedRange = "";

/**
 * Main function to set up the two-way editing system
 */
function setupTwoWayEditing() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Set up triggers for both tabs
    setupTriggers(ss);
    
    // Create a custom menu
    createCustomMenu();
    
    Logger.log("Two-way editing system setup complete!");
    
  } catch (error) {
    Logger.log("Error setting up two-way editing: " + error.toString());
  }
}

/**
 * Set up triggers for both tabs
 */
function setupTriggers(ss) {
  // Remove existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction().includes('onEdit')) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create a single onEdit trigger that handles both tabs
  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(SPREADSHEET_ID)
    .onEdit()
    .create();
}

/**
 * Unified edit handler for both tabs
 */
function onEditHandler(e) {
  if (isUpdating) return;
  
  try {
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    
    // Skip if not editing in column T or if row is less than 7
    if (range.getColumn() !== getColumnNumber(ALL_SKU_EDIT_COL)) return;
    if (range.getRow() < 7) return;
    
    const value = range.getValue();
    const row = range.getRow();
    
    // Handle edits on "All SKU Rollup WoW - FY'2025" tab
    if (sheetName === ALL_SKU_TAB_NAME) {
      // Get helper value from column W
      const helperValue = sheet.getRange(row, getColumnNumber(ALL_SKU_HELPER_COL)).getValue();
      
      if (helperValue && helperValue.toString().trim() !== "") {
        // Update the corresponding row in TOP SKU tab
        updateTopSKUTab(helperValue, value);
      }
    }
    
    // Handle edits on "TOP SKU Level Summary" tab
    else if (sheetName === TOP_SKU_TAB_NAME) {
      // Get helper value from column A
      const helperValue = sheet.getRange(row, getColumnNumber(TOP_SKU_HELPER_COL)).getValue();
      
      if (helperValue && helperValue.toString().trim() !== "") {
        // Update the corresponding row in ALL SKU tab
        updateAllSKUTab(helperValue, value);
      }
    }
    
  } catch (error) {
    Logger.log("Error in onEditHandler: " + error.toString());
  }
}

/**
 * Update the TOP SKU tab when ALL SKU tab is edited
 */
function updateTopSKUTab(helperValue, newValue) {
  if (isUpdating) return;
  
  try {
    isUpdating = true;
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const topSkuSheet = ss.getSheetByName(TOP_SKU_TAB_NAME);
    
    // Find the row with matching helper value in column A
    const helperRange = topSkuSheet.getRange(7, getColumnNumber(TOP_SKU_HELPER_COL), 
                                            topSkuSheet.getLastRow() - 6);
    const helperValues = helperRange.getValues();
    
    for (let i = 0; i < helperValues.length; i++) {
      if (helperValues[i][0] && helperValues[i][0].toString().trim() === helperValue.toString().trim()) {
        const targetRow = i + 7; // Adjust for 1-based indexing and starting at row 7
        const targetRange = topSkuSheet.getRange(targetRow, getColumnNumber(TOP_SKU_EDIT_COL));
        
        // Only update if the value is different to avoid infinite loops
        if (targetRange.getValue() !== newValue) {
          targetRange.setValue(newValue);
          Logger.log(`Updated TOP SKU tab: Row ${targetRow}, Value: ${newValue}`);
        }
        break;
      }
    }
    
  } catch (error) {
    Logger.log("Error updating TOP SKU tab: " + error.toString());
  } finally {
    isUpdating = false;
  }
}

/**
 * Update the ALL SKU tab when TOP SKU tab is edited
 */
function updateAllSKUTab(helperValue, newValue) {
  if (isUpdating) return;
  
  try {
    isUpdating = true;
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSkuSheet = ss.getSheetByName(ALL_SKU_TAB_NAME);
    
    // Find the row with matching helper value in column W
    const helperRange = allSkuSheet.getRange(7, getColumnNumber(ALL_SKU_HELPER_COL), 
                                           allSkuSheet.getLastRow() - 6);
    const helperValues = helperRange.getValues();
    
    for (let i = 0; i < helperValues.length; i++) {
      if (helperValues[i][0] && helperValues[i][0].toString().trim() === helperValue.toString().trim()) {
        const targetRow = i + 7; // Adjust for 1-based indexing and starting at row 7
        const targetRange = allSkuSheet.getRange(targetRow, getColumnNumber(ALL_SKU_EDIT_COL));
        
        // Only update if the value is different to avoid infinite loops
        if (targetRange.getValue() !== newValue) {
          targetRange.setValue(newValue);
          Logger.log(`Updated ALL SKU tab: Row ${targetRow}, Value: ${newValue}`);
        }
        break;
      }
    }
    
  } catch (error) {
    Logger.log("Error updating ALL SKU tab: " + error.toString());
  } finally {
    isUpdating = false;
  }
}

/**
 * Convert column letter to column number
 */
function getColumnNumber(columnLetter) {
  let column = 0;
  const length = columnLetter.length;
  for (let i = 0; i < length; i++) {
    column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

/**
 * Create a custom menu for easy access to functions
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Two-Way Editing')
    .addItem('Setup System', 'setupTwoWayEditing')
    .addItem('Test Connection', 'testConnection')
    .addItem('Manual Sync', 'manualSync')
    .addToUi();
}

/**
 * Test the connection between tabs
 */
function testConnection() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSkuSheet = ss.getSheetByName(ALL_SKU_TAB_NAME);
    const topSkuSheet = ss.getSheetByName(TOP_SKU_TAB_NAME);
    
    if (allSkuSheet && topSkuSheet) {
      SpreadsheetApp.getUi().alert('✅ Connection Test Successful!\n\nBoth tabs are accessible and ready for two-way editing.');
    } else {
      SpreadsheetApp.getUi().alert('❌ Connection Test Failed!\n\nPlease check tab names and permissions.');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.toString());
  }
}

/**
 * Manual sync function for testing or bulk updates
 */
function manualSync() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const allSkuSheet = ss.getSheetByName(ALL_SKU_TAB_NAME);
    const topSkuSheet = ss.getSheetByName(TOP_SKU_TAB_NAME);
    
    // Get all helper values and corresponding T column values from both sheets
    const allSkuData = getAllSkuData(allSkuSheet);
    const topSkuData = getTopSkuData(topSkuSheet);
    
    // Sync from ALL SKU to TOP SKU
    syncDataToTopSKU(allSkuData, topSkuSheet);
    
    // Sync from TOP SKU to ALL SKU
    syncDataToAllSKU(topSkuData, allSkuSheet);
    
    SpreadsheetApp.getUi().alert('✅ Manual sync completed successfully!');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Manual sync failed: ' + error.toString());
  }
}

/**
 * Get data from ALL SKU sheet
 */
function getAllSkuData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 7) return {};
  
  const helperRange = sheet.getRange(7, getColumnNumber(ALL_SKU_HELPER_COL), lastRow - 6);
  const editRange = sheet.getRange(7, getColumnNumber(ALL_SKU_EDIT_COL), lastRow - 6);
  
  const helperValues = helperRange.getValues();
  const editValues = editRange.getValues();
  
  const data = {};
  for (let i = 0; i < helperValues.length; i++) {
    if (helperValues[i][0] && helperValues[i][0].toString().trim() !== "") {
      data[helperValues[i][0].toString().trim()] = editValues[i][0];
    }
  }
  
  return data;
}

/**
 * Get data from TOP SKU sheet
 */
function getTopSkuData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 7) return {};
  
  const helperRange = sheet.getRange(7, getColumnNumber(TOP_SKU_HELPER_COL), lastRow - 6);
  const editRange = sheet.getRange(7, getColumnNumber(TOP_SKU_EDIT_COL), lastRow - 6);
  
  const helperValues = helperRange.getValues();
  const editValues = editRange.getValues();
  
  const data = {};
  for (let i = 0; i < helperValues.length; i++) {
    if (helperValues[i][0] && helperValues[i][0].toString().trim() !== "") {
      data[helperValues[i][0].toString().trim()] = editValues[i][0];
    }
  }
  
  return data;
}

/**
 * Sync data to TOP SKU sheet
 */
function syncDataToTopSKU(allSkuData, topSkuSheet) {
  const lastRow = topSkuSheet.getLastRow();
  if (lastRow < 7) return;
  
  const helperRange = topSkuSheet.getRange(7, getColumnNumber(TOP_SKU_HELPER_COL), lastRow - 6);
  const helperValues = helperRange.getValues();
  
  for (let i = 0; i < helperValues.length; i++) {
    const helperValue = helperValues[i][0];
    if (helperValue && helperValue.toString().trim() !== "") {
      const helperKey = helperValue.toString().trim();
      if (allSkuData[helperKey] !== undefined) {
        const targetRow = i + 7;
        const targetRange = topSkuSheet.getRange(targetRow, getColumnNumber(TOP_SKU_EDIT_COL));
        targetRange.setValue(allSkuData[helperKey]);
      }
    }
  }
}

/**
 * Sync data to ALL SKU sheet
 */
function syncDataToAllSKU(topSkuData, allSkuSheet) {
  const lastRow = allSkuSheet.getLastRow();
  if (lastRow < 7) return;
  
  const helperRange = allSkuSheet.getRange(7, getColumnNumber(ALL_SKU_HELPER_COL), lastRow - 6);
  const helperValues = helperRange.getValues();
  
  for (let i = 0; i < helperValues.length; i++) {
    const helperValue = helperValues[i][0];
    if (helperValue && helperValue.toString().trim() !== "") {
      const helperKey = helperValue.toString().trim();
      if (topSkuData[helperKey] !== undefined) {
        const targetRow = i + 7;
        const targetRange = allSkuSheet.getRange(targetRow, getColumnNumber(ALL_SKU_EDIT_COL));
        targetRange.setValue(topSkuData[helperKey]);
      }
    }
  }
}

/**
 * Clean up function to remove triggers
 */
function cleanup() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction().includes('onEdit')) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  Logger.log("Triggers cleaned up successfully!");
}
