/**
 * Walmart CPFR Master Automation Controller
 * Replaces manual execution of 14+ individual scripts with orchestrated automation
 */

class WalmartCPFRController {
  constructor() {
    this.config = {
      // Copy/Paste operations (your 4 manual steps)
      copyPasteOperations: [
        { source: 'CW QTD', target: 'LW QTD', function: 'copyToLWQTD' },
        { source: 'ReportUpload', target: 'LW ReportUpload', function: 'copyToLWReportUpload' },
        { source: 'CW FC', target: 'LW FC', function: 'copyToLWFC' },
        { source: 'Raw_SP Ladder', target: 'LW Raw_SP Ladder', function: 'copyToLWRawSPLadder' }
      ],
      
      // Execution schedule organized by timing
      executionSchedule: {
        // Sunday operations (CRITICAL ORDER: Backup first, then refresh)
        sunday: [
          // PHASE 1: Backup current data to "Last Week" tabs BEFORE updating
          'copyToLWQTD',               // CW QTD → LW QTD (preserve current week)
          'copyToLWReportUpload',      // ReportUpload → LW ReportUpload  
          'copyToLWFC',                // CW FC → LW FC
          'copyToLWRawSPLadder',       // Raw_SP Ladder → LW Raw_SP Ladder
          // PHASE 2: Now safe to update current week data
          'duplicateColumns',           // Notes tab
          'updateDashboard',            // Dashboard tab  
          'updateSellinHistory',        // Sellin History tab
          'updateActualOrders',         // Actual Orders tab
          'copyPasteTotalPipe',         // Total Pipeline tab
          'copyPasteSupplyLadder'       // Raw_SP Ladder tab
        ],
        
        // Monday operations
        monday: [
          'snapshotSelloutHistory'      // Sellout History (Monday afternoon)
        ],
        
        // Thursday operations  
        thursday: [
          'copyPasteQTD'               // CW QTD import (Thursday morning)
        ],
        
        // Continuous operations (already automated)
        continuous: [
          'updateSellinPrice',          // Every hour
          'updateDailyInv',            // Every hour  
          'updateProcessedPO',         // Every 12 hours
          'processWeeklyCsvEmail'      // Email trigger
        ]
      },
      
      // Execution status tracking
      status: {
        lastSundayRun: null,
        lastMondayRun: null, 
        lastThursdayRun: null,
        currentOperation: null,
        completedOperations: [],
        errors: [],
        totalDuration: 0
      }
    };
    
    this.logger = new CPFRLogger();
  }

  /**
   * MAIN EXECUTION FUNCTIONS
   */
  
  // Sunday: Full weekly refresh
  async runSundayAutomation() {
    this.logger.log('=== STARTING SUNDAY AUTOMATION ===');
    const startTime = new Date();
    
    try {
      // Phase 1: BACKUP CURRENT DATA TO "LAST WEEK" TABS (MUST BE FIRST!)
      await this.executeOperations([
        'copyToLWQTD',
        'copyToLWReportUpload', 
        'copyToLWFC',
        'copyToLWRawSPLadder'
      ], 'CRITICAL: Backup Current Data to Last Week');
      
      // Phase 2: Now safe to update current data
      await this.executeOperations([
        'duplicateColumns',
        'updateDashboard', 
        'updateSellinHistory',
        'updateActualOrders',
        'copyPasteTotalPipe',
        'copyPasteSupplyLadder'
      ], 'Update Current Week Data');
      
      this.config.status.lastSundayRun = new Date();
      this.config.status.totalDuration = new Date() - startTime;
      
      this.logger.log(`=== SUNDAY AUTOMATION COMPLETED (${this.config.status.totalDuration}ms) ===`);
      return { success: true, duration: this.config.status.totalDuration };
      
    } catch (error) {
      this.logger.error('Sunday automation failed', error);
      throw error;
    }
  }
  
  // Monday: Sellout history update
  async runMondayAutomation() {
    this.logger.log('=== STARTING MONDAY AUTOMATION ===');
    
    try {
      await this.executeOperation('snapshotSelloutHistory');
      this.config.status.lastMondayRun = new Date();
      this.logger.log('=== MONDAY AUTOMATION COMPLETED ===');
      return { success: true };
      
    } catch (error) {
      this.logger.error('Monday automation failed', error);
      throw error;
    }
  }
  
  // Thursday: QTD data import
  async runThursdayAutomation() {
    this.logger.log('=== STARTING THURSDAY AUTOMATION ===');
    
    try {
      await this.executeOperation('copyPasteQTD');
      this.config.status.lastThursdayRun = new Date();
      this.logger.log('=== THURSDAY AUTOMATION COMPLETED ===');
      return { success: true };
      
    } catch (error) {
      this.logger.error('Thursday automation failed', error);
      throw error;
    }
  }

  /**
   * EXECUTION HELPERS
   */
  
  async executeOperations(operationNames, phaseName) {
    this.logger.log(`Starting phase: ${phaseName}`);
    
    for (const operationName of operationNames) {
      await this.executeOperation(operationName);
    }
    
    this.logger.log(`Completed phase: ${phaseName}`);
  }
  
  async executeOperation(operationName) {
    this.config.status.currentOperation = operationName;
    this.logger.log(`Executing: ${operationName}`);
    
    try {
      // Execute the actual function
      await this.callFunction(operationName);
      
      this.config.status.completedOperations.push({
        name: operationName,
        timestamp: new Date(),
        success: true
      });
      
      this.logger.log(`Completed: ${operationName}`);
      
    } catch (error) {
      this.config.status.errors.push({
        operation: operationName,
        error: error.message,
        timestamp: new Date()
      });
      
      this.logger.error(`Failed: ${operationName}`, error);
      throw error;
    }
  }
  
  async callFunction(functionName) {
    // Call your existing functions
    switch (functionName) {
      // Notes tab
      case 'duplicateColumns':
        return duplicateColumns();
        
      // Dashboard tab  
      case 'updateDashboard':
        return updateDashboard();
        
      // Data import functions
      case 'updateSellinHistory':
        return updateSellinHistory();
      case 'updateActualOrders':
        return updateActualOrders();
      case 'copyPasteTotalPipe':
        return copyPasteTotalPipe();
      case 'copyPasteSupplyLadder':
        return copyPasteSupplyLadder();
      case 'snapshotSelloutHistory':
        return snapshotSelloutHistory();
      case 'copyPasteQTD':
        return copyPasteQTD();
        
      // Continuous operations
      case 'updateSellinPrice':
        return updateSellinPrice();
      case 'updateDailyInv':
        return updateDailyInv();
      case 'updateProcessedPO':
        return updateProcessedPO();
        
      // CPFR operations
      case 'updateCPFR':
        return updateCPFR();
        
      // Copy/paste functions
      case 'copyToLWQTD':
        return copyToLWQTD();
      case 'copyToLWReportUpload':
        return copyToLWReportUpload();
      case 'copyToLWFC':
        return copyToLWFC();
      case 'copyToLWRawSPLadder':
        return copyToLWRawSPLadder();
        
      default:
        throw new Error(`Unknown function: ${functionName}`);
    }
  }

  /**
   * MANUAL TRIGGER FUNCTIONS (for testing/troubleshooting)
   */
  
  async runCopyPasteOnly() {
    this.logger.log('Running copy/paste operations only - BACKUP CURRENT TO LAST WEEK');
    await this.executeOperations([
      'copyToLWQTD',
      'copyToLWReportUpload',
      'copyToLWFC', 
      'copyToLWRawSPLadder'
    ], 'Manual Backup to Last Week');
  }
  
  async runDataImportsOnly() {
    this.logger.log('Running data import operations only');
    await this.executeOperations([
      'updateSellinHistory',
      'updateActualOrders',
      'copyPasteTotalPipe',
      'copyPasteSupplyLadder'
    ], 'Manual Data Imports');
  }

  async runContinuousOperations() {
    this.logger.log('Running continuous operations (hourly)');
    await this.executeOperations([
      'updateSellinPrice',
      'updateDailyInv',
      'updateProcessedPO'
    ], 'Continuous Operations');
  }
  
  async runSpecificOperation(operationName) {
    await this.executeOperation(operationName);
  }

  /**
   * STATUS AND MONITORING
   */
  
  getStatus() {
    return {
      ...this.config.status,
      nextScheduled: this.getNextScheduledRun(),
      isHealthy: this.config.status.errors.length === 0
    };
  }
  
  getNextScheduledRun() {
    const now = new Date();
    const sunday = this.getNextWeekday(0); // Sunday = 0
    const monday = this.getNextWeekday(1); // Monday = 1  
    const thursday = this.getNextWeekday(4); // Thursday = 4
    
    return {
      nextSunday: sunday,
      nextMonday: monday,
      nextThursday: thursday
    };
  }
  
  getNextWeekday(targetDay) {
    const now = new Date();
    const currentDay = now.getDay();
    const daysUntilTarget = (targetDay + 7 - currentDay) % 7;
    const nextDate = new Date(now);
    nextDate.setDate(now.getDate() + (daysUntilTarget === 0 ? 7 : daysUntilTarget));
    return nextDate;
  }
  
  clearErrors() {
    this.config.status.errors = [];
    this.logger.log('Error log cleared');
  }
}

/**
 * LOGGING UTILITY
 */
class CPFRLogger {
  log(message) {
    const timestamp = new Date().toISOString();
    console.log(`[${timestamp}] ${message}`);
    Logger.log(`[${timestamp}] ${message}`);
  }
  
  error(message, error) {
    const timestamp = new Date().toISOString();
    console.error(`[${timestamp}] ERROR: ${message}`, error);
    Logger.log(`[${timestamp}] ERROR: ${message} - ${error.toString()}`);
  }
}

/**
 * GOOGLE SHEETS TRIGGER FUNCTIONS
 */

// Main automation functions (called by triggers)
function runSundayAutomation() {
  const controller = new WalmartCPFRController();
  return controller.runSundayAutomation();
}

function runMondayAutomation() {
  const controller = new WalmartCPFRController();
  return controller.runMondayAutomation();
}

function runThursdayAutomation() {
  const controller = new WalmartCPFRController();
  return controller.runThursdayAutomation();
}

// Manual trigger functions  
function runCopyPasteOnly() {
  const controller = new WalmartCPFRController();
  return controller.runCopyPasteOnly();
}

function runDataImportsOnly() {
  const controller = new WalmartCPFRController();
  return controller.runDataImportsOnly();
}

function runContinuousOperations() {
  const controller = new WalmartCPFRController();
  return controller.runContinuousOperations();
}



function runSpecificOperation(operationName) {
  const controller = new WalmartCPFRController();
  return controller.runSpecificOperation(operationName);
}

// Status functions
function getAutomationStatus() {
  const controller = new WalmartCPFRController();
  return controller.getStatus();
}

function clearAutomationErrors() {
  const controller = new WalmartCPFRController();
  return controller.clearErrors();
}

/**
 * TRIGGER SETUP FUNCTIONS
 */
function setupWeeklyTriggers() {
  // Clear existing triggers
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (['runSundayAutomation', 'runMondayAutomation', 'runThursdayAutomation'].includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Sunday trigger (6 AM)
  ScriptApp.newTrigger('runSundayAutomation')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(6)
    .create();
  
  // Monday trigger (2 PM)  
  ScriptApp.newTrigger('runMondayAutomation')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(14)
    .create();
    
  // Thursday trigger (8 AM)
  ScriptApp.newTrigger('runThursdayAutomation')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(8)
    .create();
    
  Logger.log('Weekly triggers setup complete');
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  Logger.log('All triggers removed');
}
function copyPasteQTD() {
  var sourceSpreadsheetId = '1bQxaNJwspIYmHGhRemKHOp0Y21fCu9Y-s3nLyFqxbiU'; // ID of the source spreadsheet
  var sourceSheetName = 'SKU level WoW delta -25Q3'; // Name of the source sheet
  Logger.log('Starting the copyPasteQTD function.');

  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet

  // Specify the target sheet by name
  var targetSheetName = "CW QTD"; // Replace with your actual target sheet name
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName); // Get the target sheet by name

  // Check if the target sheet exists
  if (!targetSheet) {
    Logger.log('"' + targetSheetName + '" sheet not found. Exiting the function.');
    return; // Exit the function if the sheet does not exist
  }

  Logger.log('The target sheet is: ' + targetSheet.getName());

  try {
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId); // Open the source spreadsheet by ID
    Logger.log('Accessing source sheet: ' + sourceSheetName);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName); // Get the source sheet by name

    // Check if the source sheet exists
    if (sourceSheet === null) throw new Error('Sheet not found: ' + sourceSheetName);

    var manualRange = sourceSheet.getRange('A1:AU3000'); // Define the range to copy from the source sheet
    var valuesToCopy = manualRange.getValues(); // Get the values from the defined range

    Logger.log('Clearing content in the target sheet range A3:AU3002.');
    // Clear the target sheet contents from A3 to AU3002
    targetSheet.getRange('A3:AU3002').clearContent();

    Logger.log('Preparing to paste values into "' + targetSheetName + '" sheet starting at A3.');
    // Paste the copied values into the target sheet starting at A3
    targetSheet.getRange('A3').offset(0, 0, 3000, manualRange.getNumColumns()).setValues(valuesToCopy);
    Logger.log('Values pasted successfully.');

    // Get the current date and time for the log message
    var currentDate = new Date();
    var timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy HH:mm:ss");

    // Calculate the end row and end column for logging purposes
    var endRow = 2 + manualRange.getNumRows() - 1; // Adjust based on your start row and the number of rows pasted
    var endColumn = columnToLetter(columnToLetterToNumber('A') + manualRange.getNumColumns() - 1);

    // Construct and log the final message with details
    var finalLogMessage = "Pasted rows from 3 to " + endRow +
                          " and columns from A to " + endColumn +
                          " on " + formattedDate;
    Logger.log(finalLogMessage);
    targetSheet.getRange('D1').setValue(finalLogMessage); // Set the final log message in the target sheet

  } catch (e) {
    // Log and display the error message with a timestamp
    var errorMessage = "Error: " + e.message;
    var currentDate = new Date();
    var timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy HH:mm:ss");
    Logger.log(formattedDate + " - " + errorMessage);
    targetSheet.getRange('D1').setValue(formattedDate + " - " + errorMessage); // Display the error message in the target sheet
  }
}

// Function to convert a column number to a letter
function columnToLetter(columnNum) {
  let temp, letter = '';
  while (columnNum > 0) {
    temp = (columnNum - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNum = (columnNum - temp - 1) / 26;
  }
  return letter;
}

// Function to convert a column letter to a number
function columnToLetterToNumber(columnLetter) {
  let column = 0, length = columnLetter.length;
  for (let i = 0; i < length; i++) {
    column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}
function copyPasteSupplyLadder() {
  var sourceSpreadsheetId = '1K2yzFBkPgeb_2L5IfZ9wKHPY7fW2gkycRy_4RJAiwKg';
  var sourceSheetName = 'WM-Charging';
  Logger.log('Starting the copyPasteSupplyLadder function.');

  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Specify the target sheet by name - Raw_SP Ladder tab
  var targetSheetName = "Raw_SP Ladder"; 
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

  if (!targetSheet) {
    Logger.log('"' + targetSheetName + '" sheet not found. Exiting the function.');
    return;
  }

  Logger.log('The target sheet is: ' + targetSheet.getName());

  try {
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    Logger.log('Accessing source sheet: ' + sourceSheetName);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

    if (sourceSheet === null) {
      throw new Error('Sheet not found: ' + sourceSheetName);
    }

    var sourceRange = sourceSheet.getDataRange();
    var valuesToCopy = sourceRange.getValues();

    // Adjust the targetRangeA1Notation if the starting point for pasting changes
    var targetRangeA1Notation = 'B2'; 
    targetSheet.getRange(targetRangeA1Notation)
               .offset(0, 0, sourceRange.getNumRows(), sourceRange.getNumColumns())
               .setValues(valuesToCopy);
    Logger.log('Values pasted successfully.');

    var currentDate = new Date();
    var timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy HH:mm:ss");

    var endRow = 1 + sourceRange.getNumRows(); 
    var endColumn = columnToLetter(columnToLetterToNumber('A') + sourceRange.getNumColumns() - 1);

    var finalLogMessage = "Pasted rows from 2 to " + endRow + 
                          " and columns from B to " + endColumn + 
                          " on " + formattedDate;
    Logger.log(finalLogMessage);
    targetSheet.getRange('G1').setValue(finalLogMessage);

  } catch (e) {
    var errorMessage = "Error: " + e.message;
    Logger.log(errorMessage);
    targetSheet.getRange('G1').setValue(errorMessage);
  }
}

function copyPasteTotalPipe() {
  var sourceSpreadsheetId = '1-VsBCoShVk106dSNali-4TuCIAanKnVpOW7n7r6U4b8';
  var sourceSheetName = 'Pipeline Overview';
  Logger.log('Starting the copyPasteTotalPipe function.');

  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Specify the target sheet by name
  var targetSheetName = "Total Pipeline"; 
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

  if (!targetSheet) {
    var errorMsg = 'Target sheet "' + targetSheetName + '" not found. Exiting the function.';
    Logger.log(errorMsg);
    throw new Error(errorMsg);
  }

  Logger.log('Target sheet found: ' + targetSheet.getName());
  Logger.log('Target spreadsheet: ' + targetSpreadsheet.getName());

  try {
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    Logger.log('Source spreadsheet opened: ' + sourceSpreadsheet.getName());
    
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

    if (sourceSheet === null) {
      throw new Error('Source sheet "' + sourceSheetName + '" not found in source spreadsheet');
    }

    Logger.log('Source sheet found: ' + sourceSheet.getName());
    
    var sourceRange = sourceSheet.getDataRange();
    var valuesToCopy = sourceRange.getValues();
    
    Logger.log('Source data range: ' + sourceRange.getA1Notation());
    Logger.log('Source rows: ' + sourceRange.getNumRows() + ', Source columns: ' + sourceRange.getNumColumns());

    // Clear the target sheet first (optional - remove if you don't want this)
    Logger.log('Clearing target sheet before pasting...');
    targetSheet.clear();

    // Adjust the targetRangeA1Notation if the starting point for pasting changes
    var targetRangeA1Notation = 'A1'; // Changed from A2 to A1 since we're clearing first
    Logger.log('Pasting data starting at: ' + targetRangeA1Notation);
    
    targetSheet.getRange(targetRangeA1Notation)
               .offset(0, 0, sourceRange.getNumRows(), sourceRange.getNumColumns())
               .setValues(valuesToCopy);
               
    Logger.log('Values pasted successfully.');

    var currentDate = new Date();
    var timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy HH:mm:ss");

    var endRow = sourceRange.getNumRows(); 
    var endColumn = columnToLetter(columnToLetterToNumber('A') + sourceRange.getNumColumns() - 1);

    var finalLogMessage = "Pasted rows from 1 to " + endRow + 
                          " and columns from A to " + endColumn + 
                          " on " + formattedDate;
    Logger.log(finalLogMessage);
    targetSheet.getRange('G1').setValue(finalLogMessage);
    
    Logger.log('Total Pipeline update completed successfully');

  } catch (e) {
    var errorMessage = "Error in copyPasteTotalPipe: " + e.message;
    Logger.log(errorMessage);
    Logger.log('Stack trace: ' + e.stack);
    
    if (targetSheet) {
      targetSheet.getRange('G1').setValue(errorMessage);
    }
    
    throw e; // Re-throw to let the calling function handle it
  }
}

function columnToLetter(columnNum) {
  let temp, letter = '';
  while (columnNum > 0) {
    temp = (columnNum - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNum = (columnNum - temp - 1) / 26;
  }
  return letter;
}

function columnToLetterToNumber(columnLetter) {
  let column = 0, length = columnLetter.length;
  for (let i = 0; i < length; i++) {
    column += (columnLetter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

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
// Script Summary:
// This script shifts cell contents to the left by one column in rows where specific text strings are found.
// It's designed to work on the active sheet of a Google Spreadsheet.
// The script processes multiple search strings and handles each row where these strings are found.

var sheet = SpreadsheetApp.getActiveSheet();

// Function to find the target column where the text "Shift" is found starting from row 2
function findTargetCol() {
  var startRow = 2;  // Start searching from row 2
  var range = sheet.getRange(startRow, 1, 1, sheet.getLastColumn());
  var values = range.getValues()[0]; // Get the first (and only) row of values

  var colToShift;
  values.forEach((value, index) => {
    if (value === "Shift") {
      colToShift = index + 1; // +1 because array indexes start at 0, but spreadsheet columns start at 1
      return;
    }
  });

  return colToShift || null; // Return the found column or null if not found
}



// Function to process shifting for multiple search texts
function _shiftAll(searchTexts) {
  // Ensure searchTexts is an array
  if (!Array.isArray(searchTexts)) {
    searchTexts = [searchTexts];
  }

  Logger.log("Search texts: " + searchTexts.join(", "));

  searchTexts.forEach(searchTxt => {
    var tf = sheet.createTextFinder(searchTxt);
    var cellsFound = tf.findAll();
    Logger.log("Found " + cellsFound.length + " cells with text: " + searchTxt);

    var rowsToProcess = cellsFound.filter(v => v.getValue() == searchTxt).map(v => v.getRow());

    if (rowsToProcess.length == 0) {
      Logger.log("No rows found with the text: " + searchTxt);
      return;
    }

    var col = findTargetCol();
    rowsToProcess.forEach((row, i) => {
      if (i > 0 && row == rowsToProcess[i - 1]) return; // Skip duplicate rows
      shiftLeftByOne(row, col);
    });
  });
}

// Function to initiate the shift process for multiple texts
function shiftAll() {
  _shiftAll(["Sellin FC (Uncon.)", "Sellin FC (Con.)", "Seasonality index"]);
}

// Function to shift cells left by one column
function shiftLeftByOne(row, col) {
  var range = sheet.getRange(row, col + 1, 1, sheet.getLastColumn() - col);
  var dst = sheet.getRange(row, col, 1, sheet.getLastColumn() - col);
  range.copyTo(dst);
}

// Run the script
function run() {
  shiftAll();
}

function snapshotSelloutHistory() {
  const sourceFileId = '1KXmfnX5dUfDfRwQUhDhvfngOb52o35a4Fnk9zoR78xM';
  const sourceSheetName = 'WMT';
  const sourceRange = 'A1:NZ1000'; // adjust if needed

  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sellout History');
  if (!targetSheet) {
    Logger.log('Target sheet "Sellout History" not found.');
    return;
  }

  const sourceSheet = SpreadsheetApp.openById(sourceFileId).getSheetByName(sourceSheetName);
  const values = sourceSheet.getRange(sourceRange).getValues();

  targetSheet.getRange('A5').offset(0, 0, values.length, values[0].length).setValues(values);

  const now = new Date();
  const formatted = Utilities.formatDate(now, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  targetSheet.getRange('B2').setValue(`Last Snapshot: ${formatted}`);
}
function updateActualOrders() {
  var sourceSpreadsheetId = '18B7eX7p_fQXyDXi_lwdf13gFNyJOjGWMX9nQB7Ak-xE';
  var sourceSheetName = 'WM Pivot Table';
  Logger.log('Starting the updateActualOrders function.');

  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Target a specific sheet by name for pasting the values
  var targetSheetName = "Actual Orders"; // Update this to your actual target sheet name
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

  // Verify that the target sheet exists
  if (!targetSheet) {
    Logger.log('"' + targetSheetName + '" sheet not found. Exiting the function.');
    return; // Exit the function if the sheet does not exist
  }

  Logger.log('The target sheet is: ' + targetSheet.getName());

  try {
    Logger.log('Attempting to open source spreadsheet.');
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    Logger.log('Accessing source sheet: ' + sourceSheetName);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

    if (sourceSheet === null) {
      throw new Error('Sheet not found: ' + sourceSheetName);
    }

    Logger.log('Retrieving data range from source sheet.');
    var sourceRange = sourceSheet.getDataRange();
    var valuesToCopy = sourceRange.getValues();

    Logger.log('Preparing to paste values into "' + targetSheetName + '" sheet.');
    var targetRangeA1Notation = 'B3'; // Adjust as needed for this script's target start cell

    // Execute the paste operation
    targetSheet.getRange(targetRangeA1Notation)
                .offset(0, 0, sourceRange.getNumRows(), sourceRange.getNumColumns())
                .setValues(valuesToCopy);
    Logger.log('Values pasted successfully.');

    // Get the current date and time for the log message
    var currentDate = new Date();
    var timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy HH:mm:ss");

    // Construct and log the final message with details
    var endRow = 1 + sourceRange.getNumRows(); // Adjusted for clarity
    var endColumn = columnToLetter(1 + sourceRange.getNumColumns()); // Assuming B2 start, hence +1
    var finalLogMessage = "Pasted rows from 2 to " + endRow + 
                          " and columns from B to " + endColumn + 
                          " on " + formattedDate;
    Logger.log(finalLogMessage);
    targetSheet.getRange('B1').setValue(finalLogMessage); // Set the final log message in the target sheet

    Logger.log('Last row and column highlighted in yellow.');
  } catch (e) {
    var errorMessage = "Error: " + e.message;
    Logger.log(errorMessage);
    targetSheet.getRange('B1').setValue(errorMessage); // Display the error message in the target sheet
  }
}

// Reuse the columnToLetter function as is
function columnToLetter(columnNum) {
  let temp, letter = '';
  while (columnNum > 0) {
    temp = (columnNum - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNum = (columnNum - temp - 1) / 26;
  }
  return letter;
}

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
    if (trigger.getHandlerFunction() === 'runContinuousOperations') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new hourly trigger
  ScriptApp.newTrigger('runContinuousOperations')
    .timeBased()
    .everyHours(1)
    .create();
  
  console.log('Hourly trigger set up successfully for continuous operations');
}

function testCopy() {
  // Run this once to test the copy function
  updateDailyInv();
}

function testTotalPipelineUpdate() {
  try {
    Logger.log('Testing Total Pipeline update...');
    copyPasteTotalPipe();
    SpreadsheetApp.getUi().alert('Total Pipeline update test completed. Check logs for details.');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Total Pipeline update test failed: ' + error.toString());
  }
}

function testCPFRUpdate() {
  try {
    Logger.log('Testing CPFR update...');
    updateCPFR();
    SpreadsheetApp.getUi().alert('CPFR update test completed. Check logs for details.');
  } catch (error) {
    SpreadsheetApp.getUi().alert('CPFR update test failed: ' + error.toString());
  }
}

function checkTotalPipelineStatus() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = spreadsheet.getSheetByName("Total Pipeline");
    
    if (!targetSheet) {
      SpreadsheetApp.getUi().alert('ERROR: "Total Pipeline" tab not found in the current spreadsheet!');
      return;
    }
    
    const lastRow = targetSheet.getLastRow();
    const lastCol = targetSheet.getLastColumn();
    const lastUpdate = targetSheet.getRange('G1').getValue();
    
    let statusMessage = 'Total Pipeline Tab Status:\n\n';
    statusMessage += 'Tab exists: ✓\n';
    statusMessage += 'Last row with data: ' + lastRow + '\n';
    statusMessage += 'Last column with data: ' + lastCol + '\n';
    statusMessage += 'Last update: ' + (lastUpdate || 'No update recorded') + '\n';
    
    SpreadsheetApp.getUi().alert(statusMessage);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error checking Total Pipeline status: ' + error.toString());
  }
}

function updateCPFR() {
  Logger.log('Starting CPFR update function');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const cpfrSheet = spreadsheet.getSheetByName("CPFR");
    
    if (!cpfrSheet) {
      throw new Error('CPFR tab not found in the current spreadsheet');
    }
    
    Logger.log('CPFR tab found, starting update process');
    
    // Step 1: Update cell J8 to today's date
    const today = new Date();
    const timeZone = spreadsheet.getSpreadsheetTimeZone();
    const formattedDate = Utilities.formatDate(today, timeZone, "MM/dd/yyyy");
    
    cpfrSheet.getRange('J8').setValue(formattedDate);
    Logger.log('Updated J8 to today\'s date: ' + formattedDate);
    
    // Step 2: Wait 30 seconds
    Logger.log('Waiting 30 seconds before proceeding with data shifts...');
    Utilities.sleep(30000); // 30 seconds = 30000 milliseconds
    Logger.log('30 second wait completed');
    
    // Step 3: Find rows where column E = 15, 22, or 24 and shift data
    const lastRow = cpfrSheet.getLastRow();
    Logger.log('Scanning rows 1 to ' + lastRow + ' for column E = 15, 22, or 24');
    
    let rowsProcessed = 0;
    let rowsShifted = 0;
    
    // Get all values from column E to find rows with value 15, 22, or 24
    const columnEValues = cpfrSheet.getRange('E1:E' + lastRow).getValues();
    
    for (let row = 1; row <= lastRow; row++) {
      const columnEValue = columnEValues[row - 1][0]; // Arrays are 0-indexed
      
      if (columnEValue === 15 || columnEValue === 22 || columnEValue === 24) {
        rowsProcessed++;
        Logger.log('Row ' + row + ' has E=' + columnEValue + ', processing data shift');
        
        try {
          let sourceRange, targetRange, sourceColumns, targetColumns, shiftDescription;
          
          if (columnEValue === 15) {
            // E = 15: Q:DC → P:DB (Q=17, DC=74, so 58 columns)
            sourceRange = cpfrSheet.getRange(row, 17, 1, 58);
            targetRange = cpfrSheet.getRange(row, 16, 1, 58); // P=16, DB=73, so 58 columns
            shiftDescription = 'Q:DC to P:DB';
          } else {
            // E = 22 or 24: AR:DC → AQ:DB (AR=44, DC=74, so 31 columns)
            sourceRange = cpfrSheet.getRange(row, 44, 1, 31);
            targetRange = cpfrSheet.getRange(row, 43, 1, 31); // AQ=43, DB=73, so 31 columns
            shiftDescription = 'AR:DC to AQ:DB';
          }
          
          const sourceValues = sourceRange.getValues();
          targetRange.setValues(sourceValues);
          
          rowsShifted++;
          Logger.log('Successfully shifted data for row ' + row + ' from ' + shiftDescription);
          
        } catch (shiftError) {
          Logger.log('Error shifting data for row ' + row + ': ' + shiftError.message);
        }
      }
    }
    
    Logger.log('CPFR update completed. Rows processed: ' + rowsProcessed + ', Rows shifted: ' + rowsShifted);
    
    // Update status in the sheet
    const statusMessage = 'CPFR updated on ' + formattedDate + ' - Rows shifted: ' + rowsShifted + ' (E=15/22/24, Q:DC→P:DB, AR:DC→AQ:DB)';
    cpfrSheet.getRange('G1').setValue(statusMessage);
    
    return {
      success: true,
      rowsProcessed: rowsProcessed,
      rowsShifted: rowsShifted,
      dateUpdated: formattedDate
    };
    
  } catch (error) {
    Logger.log('Error in updateCPFR: ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
    
    // Try to update status even if there was an error
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const cpfrSheet = spreadsheet.getSheetByName("CPFR");
      if (cpfrSheet) {
        cpfrSheet.getRange('G1').setValue('ERROR: ' + error.message);
      }
    } catch (statusError) {
      Logger.log('Could not update error status: ' + statusError.message);
    }
    
    throw error;
  }
}
function updateDashboard() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Dashboard");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "Dashboard" not found');
    return;
  }
  
  // Copy C3:I9 to C11:I17 as values only
  const sourceRange = sheet.getRange("C3:I9");
  const destRange = sheet.getRange("C11:I17");
  const values = sourceRange.getValues();
  destRange.setValues(values);
  
  // Set E1 to today's date
  const today = new Date();
  sheet.getRange("E1").setValue(today);
  
  SpreadsheetApp.getUi().alert('Dashboard updated successfully');
}

// Professional menu for Anker Demand Planning Team
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Anker CPFR Automation')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔄 Sunday Updates')
      .addItem('🔄 Run Full Sunday Automation', 'runSundayAutomation')
      .addSeparator()
      .addItem('📋 Backup Current Data to Last Week', 'runCopyPasteOnly')
      .addItem('🔄 Update Current Week Data', 'runDataImportsOnly')
      .addSeparator()
      .addItem('📊 Update Dashboard', 'updateDashboard')
      .addItem('📋 Duplicate Columns I:AH', 'duplicateColumns')
      .addItem('📈 Update Sellin History', 'updateSellinHistory')
      .addItem('📦 Update Actual Orders', 'updateActualOrders')
      .addItem('📊 Copy/Paste Total Pipeline', 'copyPasteTotalPipe')
      .addItem('📊 Copy/Paste Supply Ladder', 'copyPasteSupplyLadder'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Monday Updates')
      .addItem('📸 Snapshot Sellout History', 'snapshotSelloutHistory')
      .addSeparator()
      .addItem('📊 Update Sellin Price', 'updateSellinPrice')
      .addItem('📦 Update Processed PO', 'updateProcessedPO')
      .addItem('📊 Update Daily Inventory', 'updateDailyInv'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📅 Thursday Updates')
      .addItem('📊 Copy/Paste QTD Import', 'copyPasteQTD')
      .addSeparator()
      .addItem('🔄 Update CPFR (Original)', 'updateCPFR')
      .addItem('🚀 Complete CPFR Update (Two-Step)', 'updateCPFRComplete')
      .addItem('📋 CPFR Bulk Shift Only', 'shiftCPFRBulkColumns')
      .addItem('🎯 CPFR Complex Shift Only', 'updateCPFRComplexShift')
      .addSeparator()
      .addItem('🗺️ Update Mapping', 'updateMapping'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ System & Maintenance')
      .addItem('🔄 Run Continuous Operations', 'runContinuousOperations')
      .addItem('📧 Process Weekly CSV Email', 'processWeeklyCsvEmail')
      .addSeparator()
      .addItem('📊 Check Automation Status', 'getAutomationStatus')
      .addItem('🧹 Clear Automation Errors', 'clearAutomationErrors')
      .addSeparator()
      .addItem('🚀 Run All Weekly Updates', 'runAllWeeklyUpdates'))
    .addSeparator()
    .addItem('ℹ️ About This System', 'showSystemInfo')
    .addToUi();
}

// Master function to run all updates
function runAllWeeklyUpdates() {
  try {
    // Add your existing function names here
    updateDashboard();
    duplicateColumns(); // if this applies to weekly updates
    
    // Add calls to your existing functions like:
    // existingFunction1();
    // existingFunction2();
    // etc.
    
    SpreadsheetApp.getUi().alert('All weekly updates completed successfully');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error during updates: ' + error.toString());
  }
}

// Set up weekly schedule (run this once to establish the trigger)
function setupWeeklyTrigger() {
  // Delete existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runAllWeeklyUpdates') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new weekly trigger (runs every Monday at 9 AM)
  ScriptApp.newTrigger('runAllWeeklyUpdates')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
    
  SpreadsheetApp.getUi().alert('Weekly trigger set for Mondays at 9 AM');
}

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
    
    // Clear the destination range first
    hostSheet.clear();
    
    // Paste all data
    if (allData.length > 0) {
      const pasteRange = hostSheet.getRange(1, 1, allData.length, lastCol);
      pasteRange.setValues(allData);
      console.log(`Successfully copied ${allData.length} rows`);
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
function updateSellinHistory() {
  var sourceSpreadsheetId = '18B7eX7p_fQXyDXi_lwdf13gFNyJOjGWMX9nQB7Ak-xE';
  var sourceSheetName = 'WM Pivot Table';
  Logger.log('Starting the updateSellinHistory function.');

  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Target a specific sheet by name for pasting the values
  var targetSheetName = "Sellin History"; // Update this to your actual target sheet name
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

  // Verify that the target sheet exists
  if (!targetSheet) {
    Logger.log('"' + targetSheetName + '" sheet not found. Exiting the function.');
    return; // Exit the function if the sheet does not exist
  }

  Logger.log('The target sheet is: ' + targetSheet.getName());

  try {
    Logger.log('Attempting to open source spreadsheet.');
    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    Logger.log('Accessing source sheet: ' + sourceSheetName);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

    if (sourceSheet === null) {
      throw new Error('Sheet not found: ' + sourceSheetName);
    }

    Logger.log('Retrieving data range from source sheet.');
    var sourceRange = sourceSheet.getDataRange();
    var valuesToCopy = sourceRange.getValues();

    Logger.log('Preparing to paste values into "' + targetSheetName + '" sheet.');
    var targetRangeA1Notation = 'B2'; // Adjust as needed for this script's target start cell

    // Execute the paste operation
    targetSheet.getRange(targetRangeA1Notation)
                .offset(0, 0, sourceRange.getNumRows(), sourceRange.getNumColumns())
                .setValues(valuesToCopy);
    Logger.log('Values pasted successfully.');

    // Get the current date and time for the log message
    var currentDate = new Date();
    var timeZone = targetSpreadsheet.getSpreadsheetTimeZone();
    var formattedDate = Utilities.formatDate(currentDate, timeZone, "MM/dd/yyyy HH:mm:ss");

    // Construct and log the final message with details
    var endRow = 1 + sourceRange.getNumRows(); // Adjusted for clarity
    var endColumn = columnToLetter(1 + sourceRange.getNumColumns()); // Assuming B2 start, hence +1
    var finalLogMessage = "Pasted rows from 2 to " + endRow + 
                          " and columns from B to " + endColumn + 
                          " on " + formattedDate;
    Logger.log(finalLogMessage);
    targetSheet.getRange('B1').setValue(finalLogMessage); // Set the final log message in the target sheet

    Logger.log('Last row and column highlighted in yellow.');
  } catch (e) {
    var errorMessage = "Error: " + e.message;
    Logger.log(errorMessage);
    targetSheet.getRange('B1').setValue(errorMessage); // Display the error message in the target sheet
  }
}

function updateSellinPrice() {
  // Sheet IDs extracted from URLs
  const HOST_SHEET_ID = '1SYIFq0_AN9ziuPe4N59V1tDdHylhtbpSLUwVZMvyCsU';
  const DATA_SHEET_ID = '11iZYly0LkpllmOyUL-5zfwghQMZ4BVW_Xbrj6KvOeW0';
  
  // Configuration
  const DATA_TAB_NAME = 'Sell In Price';
  const DATA_RANGE = 'A1:V';
  const HOST_TAB_NAME = 'Sell In Price';
  const START_ROW = 15; // Clear and paste starting from row 15
  const FILTER_COLUMN = 'B'; // IPMT column
  const FILTER_VALUE = '657';
  
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
    const columnBIndex = headers.findIndex(header => header === 'Cust ID' || header.toString().toLowerCase().includes('Cust ID'));
    
    if (columnBIndex === -1) {
      throw new Error('Cust ID column not found in headers');
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
    
    // Paste the filtered data starting at row 9
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
    console.error('Error in updateSellinPrice:', error);
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
    if (trigger.getHandlerFunction() === 'updateSellinPrice') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger - runs every hour (adjust as needed)
  ScriptApp.newTrigger('updateSellinPrice')
    .timeBased()
    .everyHours(1)
    .create();
    
  console.log('Trigger created to run every hour');
}

// Test function to verify the filtering works
function testFilter() {
  const result = updateSellinPrice();
  console.log('Test result:', result);
}

function processWeeklyCsvEmail() {
  // Configuration - UPDATE THESE VALUES
  const SHEET_ID = '1SYIFq0_AN9ziuPe4N59V1tDdHylhtbpSLUwVZMvyCsU'; // Get from your Google Sheet URL
  const SHEET_NAME = 'ReportUpload'; // Name of the tab where you want data pasted
  const START_CELL = 'M2'; // Where to start pasting (A1 = top-left corner)
  const EMAIL_SUBJECT_CONTAINS = 'FANTASIA TRADING LLC Vendor# 54205587'; // Part of subject line to search for
  const SENDER_EMAIL = 'rso-am-dp.groups@anker.com'; // Email address of sender (optional filter)
  
  try {
    // Search for unread emails with CSV attachments
    let searchQuery = `has:attachment filename:csv is:unread`;
    if (EMAIL_SUBJECT_CONTAINS) {
      searchQuery += ` subject:"${EMAIL_SUBJECT_CONTAINS}"`;
    }
    if (SENDER_EMAIL) {
      searchQuery += ` from:${SENDER_EMAIL}`;
    }
    
    const threads = GmailApp.search(searchQuery, 0, 10);
    
    if (threads.length === 0) {
      console.log('No new emails with CSV attachments found');
      return;
    }
    
    // Process the most recent email
    const messages = threads[0].getMessages();
    const latestMessage = messages[messages.length - 1];
    
    console.log(`Processing email: ${latestMessage.getSubject()}`);
    
    // Find CSV attachment
    const attachments = latestMessage.getAttachments();
    let csvAttachment = null;
    
    for (let attachment of attachments) {
      if (attachment.getName() === 'data_dump.csv') {
        csvAttachment = attachment;
        break;
      }
    }
    
    if (!csvAttachment) {
      console.log('No CSV attachment found in email');
      return;
    }
    
    // Parse CSV data
    const csvContent = csvAttachment.getDataAsString();
    const csvData = parseCSV(csvContent);
    
    if (csvData.length === 0) {
      console.log('CSV file appears to be empty');
      return;
    }
    
    // Open target Google Sheet
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`Sheet named "${SHEET_NAME}" not found`);
    }
    
    // Clear only the data area where we'll paste new CSV data
    // This calculates the range based on START_CELL and previous CSV dimensions
    const startRange = sheet.getRange(START_CELL);
    const startRow = startRange.getRow();
    const startCol = startRange.getColumn();
    
    // Determine how much area to clear (50 rows x 25 columns based on your data size)
    // Adjust these numbers if your CSV dimensions change
    const maxRowsToClear = 75; // A few extra rows for safety
    const maxColsToClear = 300; // A few extra columns for safety
    
    // Clear only the specific range where CSV data goes
    const clearRange = sheet.getRange(startRow, startCol, maxRowsToClear, maxColsToClear);
    clearRange.clear();
    
    // Paste new data
    const range = sheet.getRange(START_CELL);
    const targetRange = sheet.getRange(range.getRow(), range.getColumn(), csvData.length, csvData[0].length);
    targetRange.setValues(csvData);
    
    // Sort by column AH (LSTWKPOS) (Z to A - descending order)
    // CRITICAL: Sort only the DATA rows, NOT the header row
    const sortStartRow = startRow + 1; // Start from row 3 (skip header row 2)
    const sortStartCol = startCol; // Start from column M (our pasted data)
    const sortEndRow = startRow + csvData.length - 1; // Last row of data
    const sortEndCol = startCol + csvData[0].length - 1; // End of pasted data
    
    const sortRange = sheet.getRange(sortStartRow, sortStartCol, 
                                    sortEndRow - sortStartRow + 1, 
                                    sortEndCol - sortStartCol + 1);
    
    // Column LSTWKPOS is position 22 in our range (relative to column M)
    sortRange.sort({column: 22, ascending: false}); // Z to A = descending
    
    // Mark email as read
    latestMessage.markRead();
    
    console.log(`Successfully imported ${csvData.length} rows and ${csvData[0].length} columns`);
    console.log(`Data pasted starting at ${START_CELL} in sheet "${SHEET_NAME}"`);
    console.log('Data sorted by column AH (Z to A)');
    
  } catch (error) {
    console.error('Error processing email:', error.toString());
    
    // Optional: Send yourself an error notification
    // GmailApp.sendEmail('your@email.com', 'CSV Import Error', error.toString());
  }
}

function parseCSV(csvText) {
  const lines = csvText.split('\n');
  const result = [];
  
  for (let line of lines) {
    line = line.trim();
    if (line.length === 0) continue;
    
    // Simple CSV parsing (handles basic cases)
    const row = [];
    let currentField = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === ',' && !inQuotes) {
        row.push(currentField.trim());
        currentField = '';
      } else {
        currentField += char;
      }
    }
    
    // Add the last field
    row.push(currentField.trim());
    result.push(row);
  }
  
  return result;
}

function createTrigger() {
  // Delete existing triggers
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'processWeeklyCsvEmail')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new trigger
  const trigger = ScriptApp.newTrigger('processWeeklyCsvEmail')
    .timeBased()
    .everyHours(1)
    .create();
    
  console.log('Trigger ID:', trigger.getUniqueId());
  console.log('Handler function:', trigger.getHandlerFunction());
}

// Helper function to convert column number to letter
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// Test function to verify the controller works
function testController() {
  console.log('Starting controller test...');
  
  try {
    // Test the main processing function
    processWeeklyCsvEmail();
    console.log('Controller test completed successfully');
  } catch (error) {
    console.error('Controller test failed:', error);
  }
}

// Controller class to manage all operations
class MasterController {
  static callFunction(functionName, ...args) {
    const functions = {
      'parseCSV': parseCSV,
      'createTrigger': createTrigger,
      'columnToLetter': columnToLetter,
      'testController': testController,
      'processWeeklyCsvEmail': processWeeklyCsvEmail,
      'copyToLWQTD': copyToLWQTD,
      'copyToLWReportUpload': copyToLWReportUpload,
      'copyToLWFC': copyToLWFC,
      'copyToLWRawSPLadder': copyToLWRawSPLadder,
      'duplicateColumns': duplicateColumns,
      'updateDashboard': updateDashboard,
      'updateSellinHistory': updateSellinHistory,
      'updateActualOrders': updateActualOrders,
      'copyPasteTotalPipe': copyPasteTotalPipe,
      'copyPasteSupplyLadder': copyPasteSupplyLadder,
      'snapshotSelloutHistory': snapshotSelloutHistory,
      'copyPasteQTD': copyPasteQTD,
      'updateSellinPrice': updateSellinPrice,
      'updateDailyInv': updateDailyInv,
      'updateProcessedPO': updateProcessedPO
    };
    
    if (functions[functionName]) {
      return functions[functionName](...args);
    } else {
      throw new Error(`Function ${functionName} not found`);
    }
  }
}




// Copy CW QTD data to LW QTD (Last Week QTD) tab
function copyToLWQTD() {
  console.log('copyToLWQTD function called - copying CW QTD to LW QTD');
  
  try {
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSpreadsheet.getSheetByName('CW QTD');
    const targetSheet = sourceSpreadsheet.getSheetByName('LW QTD');
    
    if (!sourceSheet) {
      throw new Error('Source sheet "CW QTD" not found');
    }
    
    if (!targetSheet) {
      throw new Error('Target sheet "LW QTD" not found');
    }
    
    // Get all data from source sheet
    const sourceRange = sourceSheet.getDataRange();
    const valuesToCopy = sourceRange.getValues();
    
    // Clear target sheet and paste new data
    targetSheet.clear();
    targetSheet.getRange(1, 1, valuesToCopy.length, valuesToCopy[0].length).setValues(valuesToCopy);
    
    console.log(`Successfully copied ${valuesToCopy.length} rows from CW QTD to LW QTD`);
    
    // Add timestamp
    const timestamp = new Date().toLocaleString();
    targetSheet.getRange('A1').setValue(`Last updated: ${timestamp}`);
    
    return {
      success: true,
      message: `Copied ${valuesToCopy.length} rows from CW QTD to LW QTD`,
      timestamp: timestamp
    };
    
  } catch (error) {
    console.error('Error in copyToLWQTD:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}
// Copy CW Report Upload data to LW Report Upload (Last Week Report Upload) tab
function copyToLWReportUpload() {
  console.log('copyToLWReportUpload function called - copying ReportUpload to LW ReportUpload');
  
  try {
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSpreadsheet.getSheetByName('ReportUpload');
    const targetSheet = sourceSpreadsheet.getSheetByName('LW ReportUpload');
    
    if (!sourceSheet) {
      throw new Error('Source sheet "ReportUpload" not found');
    }
    
    if (!targetSheet) {
      throw new Error('Target sheet "LW ReportUpload" not found');
    }
    
    // Get all data from source sheet
    const sourceRange = sourceSheet.getDataRange();
    const valuesToCopy = sourceRange.getValues();
    
    // Clear target sheet and paste new data
    targetSheet.clear();
    targetSheet.getRange(1, 1, valuesToCopy.length, valuesToCopy[0].length).setValues(valuesToCopy);
    
    console.log(`Successfully copied ${valuesToCopy.length} rows from ReportUpload to LW ReportUpload`);
    
    // Add timestamp
    const timestamp = new Date().toLocaleString();
    targetSheet.getRange('A1').setValue(`Last updated: ${timestamp}`);
    
    return {
      success: true,
      message: `Copied ${valuesToCopy.length} rows from ReportUpload to LW ReportUpload`,
      timestamp: timestamp
    };
    
  } catch (error) {
    console.error('Error in copyToLWReportUpload:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

// Copy Raw_SP Ladder data to LW Raw_SP Ladder (Last Week Raw_SP Ladder) tab
function copyToLWRawSPLadder() {
  console.log('copyToLWRawSPLadder function called - copying Raw_SP Ladder to LW Raw_SP Ladder');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = spreadsheet.getSheetByName('Raw_SP Ladder');
    const destSheet = spreadsheet.getSheetByName('LW Raw_SP Ladder');
    
    if (!sourceSheet) {
      throw new Error('Source sheet "Raw_SP Ladder" not found');
    }
    
    if (!destSheet) {
      throw new Error('Destination sheet "LW Raw_SP Ladder" not found');
    }
    
    // Get all data from A1:AW (columns A through AW)
    const sourceRange = sourceSheet.getRange("A1:AW");
    const sourceValues = sourceRange.getValues();
    
    // Clear destination sheet first
    destSheet.clear();
    
    // Paste values only to destination sheet starting at A1
    const destRange = destSheet.getRange(1, 1, sourceValues.length, sourceValues[0].length);
    destRange.setValues(sourceValues);
    
    console.log(`Successfully copied ${sourceValues.length} rows from Raw_SP Ladder to LW Raw_SP Ladder`);
    
    // Add timestamp
    const timestamp = new Date().toLocaleString();
    destSheet.getRange('A1').setValue(`Last updated: ${timestamp}`);
    
    return {
      success: true,
      message: `Copied ${sourceValues.length} rows from Raw_SP Ladder to LW Raw_SP Ladder`,
      timestamp: timestamp
    };
    
  } catch (error) {
    console.error('Error in copyToLWRawSPLadder:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

function callFunction(functionName) {
  console.log(`[${new Date().toISOString()}] Executing: ${functionName}`);
  
  try {
    let result;
    
    switch (functionName) {
      case 'copyToLWQTD':
        result = copyToLWQTD();
        break;
      case 'copyToLWReportUpload':
        result = copyToLWReportUpload();
        break;
      case 'copyToLWRawSPLadder':
        result = copyToLWRawSPLadder();
        break;
      default:
        throw new Error(`Unknown function: ${functionName}`);
    }
    
    console.log(`[${new Date().toISOString()}] Completed: ${functionName}`);
    return result;
    
  } catch (error) {
    console.error(`[${new Date().toISOString()}] ERROR: Failed: ${functionName} [${error}]`);
    console.log(`[${new Date().toISOString()}] ERROR: Failed: ${functionName} - ${error}`);
    throw error;
  }
}

// ============================================================================
// CPFR COLUMN SHIFTING FUNCTIONS (Two-Step Process)
// ============================================================================

/**
 * STEP 1: Bulk shift for standard rows (from ShiftInOneColumn.js)
 * This shifts all rows with specific text patterns but SKIPS the complex SKU rows
 */
function shiftCPFRBulkColumns() {
  Logger.log('Starting CPFR Bulk Column Shift (Step 1)');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const cpfrSheet = spreadsheet.getSheetByName('CPFR');
    
    if (!cpfrSheet) {
      throw new Error('CPFR sheet not found');
    }
    
    // Find the target column where "Shift" is found starting from row 2
    const startRow = 2;
    const range = cpfrSheet.getRange(startRow, 1, 1, cpfrSheet.getLastColumn());
    const values = range.getValues()[0];
    
    let colToShift;
    values.forEach((value, index) => {
      if (value === "Shift") {
        colToShift = index + 1; // +1 because array indexes start at 0
      }
    });
    
    if (!colToShift) {
      throw new Error('Could not find "Shift" column marker in row 2');
    }
    
    Logger.log('Found Shift column at position: ' + colToShift);
    
    // Process shifting for standard text patterns
    const searchTexts = ["Sellin FC (Uncon.)", "Sellin FC (Con.)", "Seasonality index"];
    let totalRowsShifted = 0;
    
    searchTexts.forEach(searchTxt => {
      const tf = cpfrSheet.createTextFinder(searchTxt);
      const cellsFound = tf.findAll();
      Logger.log("Found " + cellsFound.length + " cells with text: " + searchTxt);
      
      const rowsToProcess = cellsFound
        .filter(v => v.getValue() == searchTxt)
        .map(v => v.getRow());
      
      if (rowsToProcess.length == 0) {
        Logger.log("No rows found with the text: " + searchTxt);
        return;
      }
      
      // Remove duplicates and shift each row
      const uniqueRows = [...new Set(rowsToProcess)];
      uniqueRows.forEach((row, i) => {
        if (i > 0 && row == uniqueRows[i - 1]) return; // Skip duplicate rows
        shiftCPFRRowLeftByOne(cpfrSheet, row, colToShift);
        totalRowsShifted++;
      });
    });
    
    Logger.log(`CPFR Bulk Shift completed. ${totalRowsShifted} rows shifted.`);
    return {
      success: true,
      rowsShifted: totalRowsShifted,
      step: 'Bulk Shift (Step 1)'
    };
    
  } catch (error) {
    Logger.log('Error in shiftCPFRBulkColumns: ' + error.message);
    throw error;
  }
}

/**
 * Helper function to shift a single row left by one column (for bulk shifting)
 */
function shiftCPFRRowLeftByOne(sheet, row, col) {
  const sourceRange = sheet.getRange(row, col + 1, 1, sheet.getLastColumn() - col);
  const destRange = sheet.getRange(row, col, 1, sheet.getLastColumn() - col);
  
  // Use copyTo to maintain original behavior from ShiftInOneColumn.js
  sourceRange.copyTo(destRange);
}

/**
 * COMPLETE CPFR UPDATE: Two-step process combining bulk shift + complex shift
 */
function updateCPFRComplete() {
  Logger.log('Starting COMPLETE CPFR Update (Two-Step Process)');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const cpfrSheet = spreadsheet.getSheetByName("CPFR");
    
    if (!cpfrSheet) {
      throw new Error('CPFR tab not found in the current spreadsheet');
    }
    
    // Step 1: Update cell J8 to today's date
    const today = new Date();
    const timeZone = spreadsheet.getSpreadsheetTimeZone();
    const formattedDate = Utilities.formatDate(today, timeZone, "MM/dd/yyyy");
    
    cpfrSheet.getRange('J8').setValue(formattedDate);
    Logger.log('Updated J8 to today\'s date: ' + formattedDate);
    
    // Step 2: Wait 30 seconds
    Logger.log('Waiting 30 seconds before proceeding with shifts...');
    Utilities.sleep(30000);
    
    // Step 3: BULK SHIFT (handles most rows automatically)
    const bulkResult = shiftCPFRBulkColumns();
    Logger.log('Bulk shift completed: ' + bulkResult.rowsShifted + ' rows');
    
    // Step 4: COMPLEX SHIFT (handles special SKU rows with E=15,22,24)
    const complexResult = updateCPFRComplexShift();
    Logger.log('Complex shift completed: ' + complexResult.rowsShifted + ' rows');
    
    // Step 5: Update status
    const totalRows = bulkResult.rowsShifted + complexResult.rowsShifted;
    const statusMessage = `CPFR Complete Update on ${formattedDate} - Bulk: ${bulkResult.rowsShifted}, Complex: ${complexResult.rowsShifted}, Total: ${totalRows}`;
    cpfrSheet.getRange('G1').setValue(statusMessage);
    
    Logger.log('COMPLETE CPFR Update finished successfully');
    
    return {
      success: true,
      dateUpdated: formattedDate,
      bulkRowsShifted: bulkResult.rowsShifted,
      complexRowsShifted: complexResult.rowsShifted,
      totalRowsShifted: totalRows
    };
    
  } catch (error) {
    Logger.log('Error in updateCPFRComplete: ' + error.message);
    
    // Try to update error status
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const cpfrSheet = spreadsheet.getSheetByName("CPFR");
      if (cpfrSheet) {
        cpfrSheet.getRange('G1').setValue('ERROR in Complete CPFR Update: ' + error.message);
      }
    } catch (statusError) {
      Logger.log('Could not update error status: ' + statusError.message);
    }
    
    throw error;
  }
}

/**
 * STEP 2: Complex shift for special SKU rows (your original logic)
 */
function updateCPFRComplexShift() {
  Logger.log('Starting CPFR Complex Shift (Step 2)');
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const cpfrSheet = spreadsheet.getSheetByName("CPFR");
    
    if (!cpfrSheet) {
      throw new Error('CPFR sheet not found');
    }
    
    // Find rows where column E = 15, 22, or 24 and shift data
    const lastRow = cpfrSheet.getLastRow();
    Logger.log('Scanning rows 1 to ' + lastRow + ' for column E = 15, 22, or 24');
    
    let rowsProcessed = 0;
    let rowsShifted = 0;
    
    // Get all values from column E to find rows with value 15, 22, or 24
    const columnEValues = cpfrSheet.getRange('E1:E' + lastRow).getValues();
    
    for (let row = 1; row <= lastRow; row++) {
      const columnEValue = columnEValues[row - 1][0];
      
      if (columnEValue === 15 || columnEValue === 22 || columnEValue === 24) {
        rowsProcessed++;
        Logger.log('Row ' + row + ' has E=' + columnEValue + ', processing complex shift');
        
        try {
          let sourceRange, targetRange, shiftDescription;
          
          if (columnEValue === 15) {
            // E = 15: Q:DC → P:DB (Q=17, DC=74, so 58 columns)
            sourceRange = cpfrSheet.getRange(row, 17, 1, 58);
            targetRange = cpfrSheet.getRange(row, 16, 1, 58);
            shiftDescription = 'Q:DC to P:DB';
          } else {
            // E = 22 or 24: AR:DC → AQ:DB (AR=44, DC=74, so 31 columns)
            sourceRange = cpfrSheet.getRange(row, 44, 1, 31);
            targetRange = cpfrSheet.getRange(row, 43, 1, 31);
            shiftDescription = 'AR:DC to AQ:DB';
          }
          
          const sourceValues = sourceRange.getValues();
          targetRange.setValues(sourceValues);
          
          rowsShifted++;
          Logger.log('Successfully shifted complex data for row ' + row + ' from ' + shiftDescription);
          
        } catch (shiftError) {
          Logger.log('Error shifting complex data for row ' + row + ': ' + shiftError.message);
        }
      }
    }
    
    Logger.log('CPFR Complex Shift completed. Rows processed: ' + rowsProcessed + ', Rows shifted: ' + rowsShifted);
    
    return {
      success: true,
      rowsProcessed: rowsProcessed,
      rowsShifted: rowsShifted,
      step: 'Complex Shift (Step 2)'
    };
    
  } catch (error) {
    Logger.log('Error in updateCPFRComplexShift: ' + error.message);
    throw error;
  }
}

/**
 * Display professional system information for the Anker team
 */
function showSystemInfo() {
  const ui = SpreadsheetApp.getUi();
  
  const infoMessage = 
    "📊 Anker CPFR Automation System\n\n" +
    "This system automates critical demand planning operations:\n\n" +
    "🔄 SUNDAY UPDATES:\n" +
    "• Backup current week data to 'Last Week' tabs\n" +
    "• Refresh all current week data and dashboards\n" +
    "• Update pipeline and supply ladder information\n\n" +
    "📅 MONDAY UPDATES:\n" +
    "• Capture sellout history snapshots\n" +
    "• Update pricing and inventory data\n\n" +
    "📅 THURSDAY UPDATES:\n" +
    "• Import QTD data and update CPFR\n" +
    "• Refresh mapping and reference data\n\n" +
    "⚙️ CONTINUOUS OPERATIONS:\n" +
    "• Hourly price and inventory updates\n" +
    "• Automated email processing\n\n" +
    "Developed for Anker Demand Planning Team\n" +
    "Contact: Your IT/Demand Planning Lead";
  
  ui.alert('System Information', infoMessage, ui.ButtonSet.OK);
}
