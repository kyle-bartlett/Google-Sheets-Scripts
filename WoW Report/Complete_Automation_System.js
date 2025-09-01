// COMPLETE AUTOMATION SYSTEM - Professional Weekly Data Processing
// Designed for: 1) Automated weekly runs  2) Beautiful manual overrides

const SOURCE_FILE_ID = '1bQxaNJwspIYmHGhRemKHOp0Y21fCu9Y-s3nLyFqxbiU';
const PARENT_FILE_ID = '1KEdrPKJwFYOCNIR80YItazraOGeOm0YlaXAoovicHN4';
const TAB_1_NAME = 'YTD rev';
const OUTPUT_TAB_NAME = 'QTD+FC Raw Data';

// Column mapping
const COLUMN_MAPPING = {
  0: 0, 12: 1, 4: 2, 3: 3, 5: 4, 7: 5, 8: 6, 9: 7, 
  10: 8, 17: 9, 11: 10, 15: 11, 13: 12, 2: 13
};

// ==========================================
// CONFIGURATION SYSTEM
// ==========================================

/**
 * DEFAULT AUTOMATION SETTINGS
 * 
 * This is what runs automatically every week without human interaction.
 * Modify these defaults based on your standard business process.
 */
const DEFAULT_SETTINGS = {
  // Thursday (YTD Receipts) Settings
  thursday: {
    enabled: true,
    quartersToProcess: "AUTO", // "AUTO" = current quarter, or specify ["2025Q3"]
    autoQuarterDetection: true,
    transitionPeriodDays: 21 // Include previous quarter if within 21 days of quarter end
  },
  
  // Friday (Forecast) Settings  
  friday: {
    enabled: true,
    quartersToProcess: "ALL", // "ALL" = all quarters in source, or specify ["2025Q3", "2025Q4"]
    cleanupBeforeProcessing: true,
    autoDetectWeeklyTab: true
  },
  
  // General Settings
  general: {
    runSequentially: true, // Thursday first, then Friday
    enableProgressLogging: true,
    enableChunkingForLargeDatasets: true,
    maxProcessingTimeMinutes: 5
  }
};

/**
 * MANUAL OVERRIDE CONFIGURATION
 * 
 * When users want custom settings, they modify this section.
 * Set manualMode = true to override defaults.
 */
const MANUAL_OVERRIDE = {
  manualMode: false, // Set to true to use manual settings below
  
  // Manual Thursday Settings
  thursday: {
    enabled: true,
    quartersToProcess: ["2025Q3"], // Specify exact quarters
    customDescription: "Normal Q3 processing"
  },
  
  // Manual Friday Settings
  friday: {
    enabled: true,
    quartersToProcess: ["2025Q3", "2025Q4"], // Specify exact quarters for cleanup
    cleanupQuarters: ["2025Q3", "2025Q4"], // Which quarters to clean up
    customDescription: "Q3/Q4 forecast processing"
  },
  
  // Manual General Settings
  general: {
    runThursdayOnly: false,
    runFridayOnly: false,
    skipCleanup: false
  }
};

// ==========================================
// MAIN AUTOMATION FUNCTIONS
// ==========================================

/**
 * MAIN WEEKLY AUTOMATION - Run this for hands-off processing
 * This is what gets triggered automatically every week
 */
function runWeeklyAutomation() {
  try {
    console.log('🤖 STARTING WEEKLY AUTOMATION');
    console.log('==========================================');
    
    const settings = getEffectiveSettings();
    displayProcessingPlan(settings);
    
    let thursdayResults = null;
    let fridayResults = null;
    
    // Thursday Processing
    if (settings.thursday.enabled) {
      console.log('\n🗓️  THURSDAY PROCESSING (YTD Receipts)');
      console.log('------------------------------------------');
      thursdayResults = runThursdayProcessing(settings.thursday);
    } else {
      console.log('\n🗓️  THURSDAY PROCESSING: SKIPPED');
    }
    
    // Friday Processing
    if (settings.friday.enabled) {
      console.log('\n🗓️  FRIDAY PROCESSING (Forecasts)');
      console.log('------------------------------------------');
      fridayResults = runFridayProcessing(settings.friday);
    } else {
      console.log('\n🗓️  FRIDAY PROCESSING: SKIPPED');
    }
    
    // Summary
    console.log('\n🎉 WEEKLY AUTOMATION COMPLETED');
    console.log('==========================================');
    if (thursdayResults) {
      console.log(`Thursday: ${thursdayResults.rowsProcessed} rows processed`);
    }
    if (fridayResults) {
      console.log(`Friday: ${fridayResults.rowsProcessed} rows processed`);
    }
    
    return { thursday: thursdayResults, friday: fridayResults };
    
  } catch (error) {
    console.error('❌ Weekly automation failed:', error);
    throw error;
  }
}

/**
 * Get effective settings (default or manual override)
 */
function getEffectiveSettings() {
  if (MANUAL_OVERRIDE.manualMode) {
    console.log('⚙️  Using MANUAL OVERRIDE settings');
    return processManualSettings();
  } else {
    console.log('🤖 Using DEFAULT AUTOMATION settings');
    return processDefaultSettings();
  }
}

/**
 * Process default automation settings
 */
function processDefaultSettings() {
  const settings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS)); // Deep copy
  
  // Auto-detect current quarter for Thursday
  if (settings.thursday.quartersToProcess === "AUTO") {
    const currentQuarter = getCurrentQuarter();
    const quarters = [currentQuarter];
    
    // Check if we're in transition period
    if (settings.thursday.transitionPeriodDays > 0) {
      const previousQuarter = getPreviousQuarter(currentQuarter);
      if (isInTransitionPeriod(settings.thursday.transitionPeriodDays)) {
        quarters.unshift(previousQuarter); // Add previous quarter first
        console.log(`📅 Transition period detected - including ${previousQuarter}`);
      }
    }
    
    settings.thursday.quartersToProcess = quarters;
  }
  
  // Auto-detect all quarters for Friday
  if (settings.friday.quartersToProcess === "ALL") {
    settings.friday.quartersToProcess = getAllAvailableQuarters();
  }
  
  return settings;
}

/**
 * Process manual override settings
 */
function processManualSettings() {
  const settings = JSON.parse(JSON.stringify(MANUAL_OVERRIDE)); // Deep copy
  
  // Validate manual settings
  if (!Array.isArray(settings.thursday.quartersToProcess)) {
    throw new Error('Manual Thursday quarters must be an array');
  }
  
  if (!Array.isArray(settings.friday.quartersToProcess)) {
    throw new Error('Manual Friday quarters must be an array');
  }
  
  return settings;
}

/**
 * Display the processing plan before execution
 */
function displayProcessingPlan(settings) {
  console.log('\n📋 PROCESSING PLAN');
  console.log('==========================================');
  
  if (settings.thursday.enabled) {
    console.log('🗓️  Thursday (YTD Receipts):');
    console.log(`   Quarters: ${settings.thursday.quartersToProcess.join(', ')}`);
    if (settings.thursday.customDescription) {
      console.log(`   Description: ${settings.thursday.customDescription}`);
    }
  }
  
  if (settings.friday.enabled) {
    console.log('🗓️  Friday (Forecasts):');
    console.log(`   Quarters: ${settings.friday.quartersToProcess.join(', ')}`);
    if (settings.friday.cleanupQuarters) {
      console.log(`   Cleanup: ${settings.friday.cleanupQuarters.join(', ')}`);
    }
    if (settings.friday.customDescription) {
      console.log(`   Description: ${settings.friday.customDescription}`);
    }
  }
  
  console.log('==========================================');
}

// ==========================================
// THURSDAY PROCESSING (Enhanced)
// ==========================================

/**
 * Run Thursday processing with given settings
 */
function runThursdayProcessing(thursdaySettings) {
  try {
    const quarters = thursdaySettings.quartersToProcess;
    console.log(`Processing Thursday quarters: ${quarters.join(', ')}`);
    
    // Step 1: Cleanup
    console.log('Step 1: Cleaning up existing Thursday data...');
    cleanupSelectedQuarters(quarters);
    
    // Step 2: Get and process data
    console.log('Step 2: Getting source data...');
    const sourceData = getSelectedQuartersSourceData(quarters);
    
    if (!sourceData || sourceData.data.length === 0) {
      console.log('⚠️  No Thursday data found for selected quarters');
      return { rowsProcessed: 0, quarters: quarters };
    }
    
    // Step 3: Process based on data volume
    const dataSize = sourceData.data.length;
    if (dataSize > 25000) {
      processLargeDataset(sourceData, 'Thursday');
    } else {
      processStandardDataset(sourceData, 'Thursday');
    }
    
    return { rowsProcessed: dataSize, quarters: quarters };
    
  } catch (error) {
    console.error('Error in Thursday processing:', error);
    throw error;
  }
}

// ==========================================
// FRIDAY PROCESSING (New Enhanced Version)
// ==========================================

/**
 * Run Friday processing with given settings
 */
function runFridayProcessing(fridaySettings) {
  try {
    console.log(`Processing Friday with settings:`, fridaySettings);
    
    // Step 1: Cleanup Friday data if requested
    if (fridaySettings.cleanupBeforeProcessing) {
      const cleanupQuarters = fridaySettings.cleanupQuarters || fridaySettings.quartersToProcess;
      console.log('Step 1: Cleaning up existing Friday data...');
      cleanupFridayData(cleanupQuarters);
    }
    
    // Step 2: Find current weekly tab
    console.log('Step 2: Finding current weekly forecast tab...');
    const currentWeeklyTab = findCurrentWeeklyTab();
    if (!currentWeeklyTab) {
      throw new Error('No current weekly forecast tab found');
    }
    
    console.log(`Found weekly tab: ${currentWeeklyTab}`);
    
    // Step 3: Get forecast data
    console.log('Step 3: Getting forecast data...');
    const forecastData = getForecastData(currentWeeklyTab, fridaySettings.quartersToProcess);
    
    if (!forecastData || forecastData.data.length === 0) {
      console.log('⚠️  No Friday forecast data found');
      return { rowsProcessed: 0, weeklyTab: currentWeeklyTab };
    }
    
    // Step 4: Process forecast data
    console.log('Step 4: Processing forecast data...');
    const dataSize = forecastData.data.length;
    
    if (dataSize > 25000) {
      processFridayLargeDataset(forecastData, currentWeeklyTab);
    } else {
      processFridayStandardDataset(forecastData, currentWeeklyTab);
    }
    
    return { rowsProcessed: dataSize, weeklyTab: currentWeeklyTab };
    
  } catch (error) {
    console.error('Error in Friday processing:', error);
    throw error;
  }
}

/**
 * Clean up existing Friday data (Open FC labels in Column Q)
 */
function cleanupFridayData(quartersToCleanup) {
  try {
    const parentFile = SpreadsheetApp.openById(PARENT_FILE_ID);
    const outputSheet = parentFile.getSheetByName(OUTPUT_TAB_NAME);
    
    const lastRow = outputSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('No Friday data to clean up');
      return;
    }
    
    console.log(`🧹 Cleaning up Friday data for quarters: ${quartersToCleanup.join(', ')}`);
    
    // Get Column Q (Open FC) and Column P (Season) data
    const lastCol = outputSheet.getLastColumn();
    let columnQData = [];
    let seasonData = [];
    
    if (lastCol >= 17) { // Column Q exists
      columnQData = outputSheet.getRange(1, 17, lastRow, 1).getValues().flat();
    }
    if (lastCol >= 16) { // Column P exists  
      seasonData = outputSheet.getRange(1, 16, lastRow, 1).getValues().flat();
    }
    
    // Find Friday rows to delete (have "Open FC" in Column Q)
    const rowsToDelete = [];
    
    for (let i = 0; i < Math.min(columnQData.length, lastRow); i++) {
      const openFCLabel = columnQData[i];
      const seasonValue = seasonData[i] || '';
      const rowNum = i + 1;
      
      // Delete if it has "Open FC" and matches cleanup quarters (or empty season)
      if (openFCLabel === 'Open FC') {
        if (quartersToCleanup.includes(seasonValue) || seasonValue === '') {
          rowsToDelete.push(rowNum);
        }
      }
    }
    
    console.log(`📋 Found ${rowsToDelete.length} Friday forecast rows to delete`);
    
    // Delete rows (from bottom to top)
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      outputSheet.deleteRow(rowsToDelete[i]);
    }
    
    console.log(`🗑️  Deleted ${rowsToDelete.length} old Friday forecast rows`);
    
  } catch (error) {
    console.error('Error cleaning up Friday data:', error);
    throw error;
  }
}

/**
 * Get forecast data with optional quarter filtering
 */
function getForecastData(weeklyTabName, quartersFilter) {
  try {
    const sourceFile = SpreadsheetApp.openById(SOURCE_FILE_ID);
    const sheet = sourceFile.getSheetByName(weeklyTabName);
    
    if (!sheet) {
      throw new Error(`Weekly tab "${weeklyTabName}" not found`);
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) {
      throw new Error('No forecast data found');
    }
    
    // Get headers and all data
    const headers = sheet.getRange(2, 1, 1, 18).getValues()[0];
    const allData = sheet.getRange(3, 1, lastRow - 2, 18).getValues();
    
    // If quartersFilter is specified and not "ALL", filter the data
    if (Array.isArray(quartersFilter) && quartersFilter.length > 0) {
      console.log(`📊 Filtering forecast data for quarters: ${quartersFilter.join(', ')}`);
      
      const filteredData = [];
      for (const row of allData) {
        const seasonValue = row[15]; // Column P (Season)
        if (quartersFilter.includes(seasonValue)) {
          filteredData.push(row);
        }
      }
      
      console.log(`✅ Filtered to ${filteredData.length} rows from ${allData.length} total`);
      return { headers, data: filteredData };
    } else {
      console.log(`✅ Using all ${allData.length} forecast rows (no filtering)`);
      return { headers, data: allData };
    }
    
  } catch (error) {
    console.error('Error getting forecast data:', error);
    throw error;
  }
}

/**
 * Process Friday standard dataset
 */
function processFridayStandardDataset(forecastData, weeklyTab) {
  try {
    const transformedData = transformData(forecastData.data);
    
    const parentFile = SpreadsheetApp.openById(PARENT_FILE_ID);
    const outputSheet = parentFile.getSheetByName(OUTPUT_TAB_NAME);
    const startRow = outputSheet.getLastRow() + 1;
    
    // Write data starting in column B
    if (transformedData.length > 0) {
      const range = outputSheet.getRange(startRow, 2, transformedData.length, 14);
      range.setValues(transformedData);
      console.log(`✅ Written ${transformedData.length} Friday rows starting at row ${startRow}`);
    }
    
    // Add Friday labels
    addFridayLabels(outputSheet, startRow, transformedData.length);
    
  } catch (error) {
    console.error('Error in Friday standard processing:', error);
    throw error;
  }
}

/**
 * Add Friday-specific labels
 */
function addFridayLabels(outputSheet, startRow, rowCount) {
  // Add "Current Wk" labels in column A
  const currentWkLabels = new Array(rowCount).fill(['Current Wk']);
  if (currentWkLabels.length > 0) {
    outputSheet.getRange(startRow, 1, rowCount, 1).setValues(currentWkLabels);
    console.log(`✅ Added ${rowCount} "Current Wk" labels for Friday data`);
  }
  
  // Add "Open FC" labels in column Q (column 17)
  const openFCLabels = new Array(rowCount).fill(['Open FC']);
  if (openFCLabels.length > 0) {
    outputSheet.getRange(startRow, 17, rowCount, 1).setValues(openFCLabels);
    console.log(`✅ Added ${rowCount} "Open FC" labels in column Q`);
  }
}

// ==========================================
// UTILITY FUNCTIONS
// ==========================================

/**
 * Get current quarter (dynamic)
 */
function getCurrentQuarter() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  
  let quarter;
  if (month <= 3) quarter = 1;
  else if (month <= 6) quarter = 2;
  else if (month <= 9) quarter = 3;
  else quarter = 4;
  
  return `${year}Q${quarter}`;
}

/**
 * Get previous quarter
 */
function getPreviousQuarter(currentQuarter) {
  const [year, quarter] = currentQuarter.split('Q');
  const yearNum = parseInt(year);
  const quarterNum = parseInt(quarter);
  
  if (quarterNum === 1) {
    return `${yearNum - 1}Q4`;
  } else {
    return `${yearNum}Q${quarterNum - 1}`;
  }
}

/**
 * Check if we're in transition period
 */
function isInTransitionPeriod(transitionDays) {
  const now = new Date();
  const month = now.getMonth() + 1; // 1-12
  const day = now.getDate();
  
  // Check if we're within transitionDays of quarter end
  const quarterEndDates = [
    { month: 3, day: 31 },  // Q1 end
    { month: 6, day: 30 },  // Q2 end  
    { month: 9, day: 30 },  // Q3 end
    { month: 12, day: 31 }  // Q4 end
  ];
  
  for (const endDate of quarterEndDates) {
    if (month === endDate.month) {
      const daysUntilEnd = endDate.day - day;
      if (daysUntilEnd >= 0 && daysUntilEnd <= transitionDays) {
        return true;
      }
    }
  }
  
  return false;
}

/**
 * Get all available quarters from source data
 */
function getAllAvailableQuarters() {
  try {
    const sourceFile = SpreadsheetApp.openById(SOURCE_FILE_ID);
    
    // Check current weekly tab for available quarters
    const currentWeeklyTab = findCurrentWeeklyTab();
    if (currentWeeklyTab) {
      const sheet = sourceFile.getSheetByName(currentWeeklyTab);
      if (sheet && sheet.getLastRow() >= 3) {
        const seasonData = sheet.getRange(3, 16, sheet.getLastRow() - 2, 1).getValues().flat();
        const uniqueQuarters = [...new Set(seasonData.filter(s => s && s !== ''))];
        return uniqueQuarters.sort();
      }
    }
    
    return [getCurrentQuarter()]; // Fallback
    
  } catch (error) {
    console.error('Error getting available quarters:', error);
    return [getCurrentQuarter()];
  }
}

/**
 * Find current weekly tab (same as before)
 */
function findCurrentWeeklyTab() {
  try {
    const sourceFile = SpreadsheetApp.openById(SOURCE_FILE_ID);
    const sheets = sourceFile.getSheets();
    
    let currentWeeklyTab = null;
    let highestWeekNumber = 0;
    
    const WEEKLY_TAB_PATTERN = /^WK\d{2} Sell in FCST$/;
    
    for (const sheet of sheets) {
      const name = sheet.getName();
      if (WEEKLY_TAB_PATTERN.test(name)) {
        const weekMatch = name.match(/^WK(\d{2}) Sell in FCST$/);
        if (weekMatch) {
          const weekNumber = parseInt(weekMatch[1]);
          if (weekNumber > highestWeekNumber) {
            highestWeekNumber = weekNumber;
            currentWeeklyTab = name;
          }
        }
      }
    }
    
    return currentWeeklyTab;
    
  } catch (error) {
    console.error('Error finding current weekly tab:', error);
    return null;
  }
}

// ==========================================
// BEAUTIFUL MANUAL OVERRIDE FUNCTIONS
// ==========================================

/**
 * Enable manual mode with beautiful interface
 */
function enableManualMode() {
  console.log('🎨 MANUAL OVERRIDE MODE ACTIVATED');
  console.log('==========================================');
  console.log('Edit the MANUAL_OVERRIDE configuration section:');
  console.log('');
  console.log('1. Set manualMode: true');
  console.log('2. Configure thursday.quartersToProcess: ["2025Q3"]');
  console.log('3. Configure friday.quartersToProcess: ["2025Q3", "2025Q4"]');
  console.log('4. Run runWeeklyAutomation() to execute');
  console.log('');
  console.log('For common scenarios:');
  console.log('• setupCurrentQuarterOnly() - Normal operations');
  console.log('• setupQuarterTransition() - Quarter transition period');
  console.log('• setupCustomScenario() - Custom quarter selection');
}

/**
 * Quick setup for current quarter only
 */
function setupCurrentQuarterOnly() {
  const currentQ = getCurrentQuarter();
  console.log('⚙️  SETUP: Current Quarter Only');
  console.log('==========================================');
  console.log('Copy this into MANUAL_OVERRIDE:');
  console.log('');
  console.log('manualMode: true,');
  console.log('thursday: {');
  console.log('  enabled: true,');
  console.log(`  quartersToProcess: ["${currentQ}"],`);
  console.log('  customDescription: "Normal current quarter processing"');
  console.log('},');
  console.log('friday: {');
  console.log('  enabled: true,');
  console.log(`  quartersToProcess: ["${currentQ}"],`);
  console.log(`  cleanupQuarters: ["${currentQ}"],`);
  console.log('  customDescription: "Current quarter forecast"');
  console.log('}');
}

// Include all previous functions (cleanup, transform, etc.) here...
// [Previous functions from the Thursday script would go here]
