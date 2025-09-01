/**
 * ANKER FORECAST AUTOMATION - FIXED VERSION
 * This script transforms wide forecast data to tall format and creates analysis
 * Updated to work with your actual sheet names and structure
 */

function transformForecastData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Log all available sheet names for debugging
  const allSheets = ss.getSheets();
  console.log('Available sheets:');
  allSheets.forEach(sheet => console.log(`  - "${sheet.getName()}"`));
  
  // Get source sheets with correct names
  const constrainedSheet = ss.getSheetByName('Constrained Wide');
  const unconstrainedSheet = ss.getSheetByName('Unconstrained Wide');
  
  // Check if sheets exist
  if (!constrainedSheet) {
    throw new Error('Could not find "Constrained Wide" sheet. Available sheets: ' + 
                   allSheets.map(s => s.getName()).join(', '));
  }
  if (!unconstrainedSheet) {
    throw new Error('Could not find "Unconstrained Wide" sheet. Available sheets: ' + 
                   allSheets.map(s => s.getName()).join(', '));
  }
  
  // Create or clear output sheet
  let outputSheet = ss.getSheetByName('Looker_Ready_View_New');
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet('Looker_Ready_View_New');
  }
  
  // Headers for output
  const headers = ['Customer', 'Anker SKU', 'PDT', 'Forecast Type', 'Quarter', 'Week', 
                   'Forecast - Units', 'Forecast Revenue', 'Delta Units', 'Delta - Revenue', 
                   'Gap Flag', 'IsCurrentQ', 'Helper', 'Sell-In Price'];
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get data ranges
  const constrainedData = constrainedSheet.getDataRange().getValues();
  const unconstrainedData = unconstrainedSheet.getDataRange().getValues();
  
  console.log(`Constrained data: ${constrainedData.length} rows x ${constrainedData[0].length} columns`);
  console.log(`Unconstrained data: ${unconstrainedData.length} rows x ${unconstrainedData[0].length} columns`);
  
  // Find week columns (they start with 2025xx format)
  const weekColumns = [];
  const headerRow = constrainedData[0];
  
  for (let i = 0; i < headerRow.length; i++) {
    const cellValue = headerRow[i];
    if (cellValue && cellValue.toString().match(/^202\d\d\d$/)) {
      weekColumns.push({index: i, week: cellValue.toString()});
    }
  }
  
  console.log(`Found ${weekColumns.length} week columns:`, weekColumns.map(w => w.week));
  
  if (weekColumns.length === 0) {
    throw new Error('No week columns found. Expected columns with format 202xxx');
  }
  
  // Create unconstrained lookup
  const unconstrainedLookup = {};
  for (let i = 1; i < unconstrainedData.length; i++) {
    const row = unconstrainedData[i];
    const helper = row[0]; // Helper column
    
    if (!helper) continue; // Skip empty rows
    
    weekColumns.forEach(weekCol => {
      const week = weekCol.week;
      const units = parseFloat(row[weekCol.index]) || 0;
      const sellInPrice = parseFloat(row[1]) || 0; // Sell-in Price column
      
      const key = `${helper}_${week}`;
      unconstrainedLookup[key] = {
        units: units,
        revenue: units * sellInPrice
      };
    });
  }
  
  console.log(`Created unconstrained lookup with ${Object.keys(unconstrainedLookup).length} entries`);
  
  const outputData = [];
  
  // Process constrained data
  for (let i = 1; i < constrainedData.length; i++) {
    const row = constrainedData[i];
    const helper = row[0];          // Helper
    const sellInPrice = parseFloat(row[1]) || 0;  // Sell-in price
    const pct = row[2];             // PCT
    const pdt = row[3];             // PDT
    const customerId = row[4];      // Customer ID
    const customer = row[5];        // Customer
    const sku = row[6];             // Anker SKU
    const skuDescription = row[7];  // SKU Description
    
    if (!helper || !customer || !sku) continue; // Skip incomplete rows
    
    weekColumns.forEach(weekCol => {
      const week = weekCol.week;
      const constrainedUnits = parseFloat(row[weekCol.index]) || 0;
      const constrainedRevenue = constrainedUnits * sellInPrice;
      
      // Look up unconstrained data
      const key = `${helper}_${week}`;
      const unconstrained = unconstrainedLookup[key] || { units: 0, revenue: 0 };
      
      const deltaUnits = constrainedUnits - unconstrained.units;
      const deltaRevenue = constrainedRevenue - unconstrained.revenue;
      
      const quarter = getQuarter(parseInt(week));
      const isCurrentQ = quarter === 'Q4 2025';
      const gapFlag = deltaUnits < 0 ? 'Supply Gap' : '';
      
      // Add constrained row
      outputData.push([
        customer,
        sku,
        pdt,
        'Constrained',
        quarter,
        parseInt(week),
        constrainedUnits,
        constrainedRevenue,
        deltaUnits,
        deltaRevenue,
        gapFlag,
        isCurrentQ,
        helper,
        sellInPrice
      ]);
      
      // Add unconstrained row
      outputData.push([
        customer,
        sku,
        pdt,
        'Unconstrained',
        quarter,
        parseInt(week),
        unconstrained.units,
        unconstrained.revenue,
        0, // Delta is 0 for unconstrained (baseline)
        0,
        '',
        isCurrentQ,
        helper,
        sellInPrice
      ]);
    });
  }
  
  // Write output data
  if (outputData.length > 0) {
    console.log(`Writing ${outputData.length} rows to output sheet`);
    outputSheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
    
    // Format the sheet
    formatOutputSheet(outputSheet, headers.length);
  }
  
  console.log(`Transformation complete! ${outputData.length} records created.`);
  
  // Create summary sheets
  createSummarySheets(ss);
  
  return outputData.length;
}

function getQuarter(week) {
  if (week >= 202529 && week <= 202539) return 'Q3 2025';
  if (week >= 202540 && week <= 202552) return 'Q4 2025';
  if (week >= 202553 && week <= 202605) return 'Q1 2026';
  return 'Unknown';
}

function formatOutputSheet(sheet, numColumns) {
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // Auto-resize columns
  for (let i = 1; i <= numColumns; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Format currency columns (Revenue columns)
  const revenueColumns = [8, 10]; // Forecast Revenue, Delta Revenue
  revenueColumns.forEach(col => {
    if (col <= numColumns) {
      const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
      range.setNumberFormat('$#,##0.00');
    }
  });
  
  // Format number columns (Units columns)
  const unitColumns = [7, 9]; // Forecast Units, Delta Units
  unitColumns.forEach(col => {
    if (col <= numColumns) {
      const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
      range.setNumberFormat('#,##0');
    }
  });
  
  // Conditional formatting for Gap Flag
  const gapFlagRange = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Supply Gap')
    .setBackground('#FFE6E6')
    .setFontColor('#CC0000')
    .setRanges([gapFlagRange])
    .build();
  sheet.setConditionalFormatRules([rule]);
}

function createSummarySheets(ss) {
  console.log('Creating summary sheets...');
  
  // Create or update SKU Summary
  createSKUSummary(ss);
  
  // Create or update Customer Summary
  createCustomerSummary(ss);
  
  // Create or update Weekly Trends
  createWeeklyTrends(ss);
  
  console.log('Summary sheets created successfully');
}

function createSKUSummary(ss) {
  let sheet = ss.getSheetByName('SKU_Summary_New');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('SKU_Summary_New');
  }
  
  // Headers
  const headers = ['Rank', 'SKU', 'PDT', 'Gap Units', 'Revenue Impact', 'Customers Affected'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // This would normally be a pivot table, but we'll create formulas
  // Note: In practice, you'd want to create actual pivot tables
  sheet.getRange('A2').setFormula('=ROW()-1');
  sheet.getRange('B2').setValue('Use pivot table from Looker_Ready_View_New');
  
  formatOutputSheet(sheet, headers.length);
}

function createCustomerSummary(ss) {
  let sheet = ss.getSheetByName('Customer_Summary_New');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Customer_Summary_New');
  }
  
  const headers = ['Rank', 'Customer', 'Gap Units', 'Revenue Impact', 'SKUs Affected'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  sheet.getRange('A2').setFormula('=ROW()-1');
  sheet.getRange('B2').setValue('Use pivot table from Looker_Ready_View_New');
  
  formatOutputSheet(sheet, headers.length);
}

function createWeeklyTrends(ss) {
  let sheet = ss.getSheetByName('Weekly_Trends_New');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Weekly_Trends_New');
  }
  
  const headers = ['Week', 'Quarter', 'Gap Units', 'Revenue Impact', 'Records Count'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  sheet.getRange('A2').setValue('Use pivot table from Looker_Ready_View_New');
  
  formatOutputSheet(sheet, headers.length);
}

/**
 * Test function to verify sheet existence
 */
function testSheetAccess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  
  console.log('Testing sheet access...');
  console.log(`Total sheets: ${allSheets.length}`);
  
  allSheets.forEach((sheet, index) => {
    const name = sheet.getName();
    const rows = sheet.getLastRow();
    const cols = sheet.getLastColumn();
    console.log(`${index + 1}. "${name}" - ${rows} rows x ${cols} columns`);
  });
  
  // Test specific sheets
  const constrainedSheet = ss.getSheetByName('Constrained Wide');
  const unconstrainedSheet = ss.getSheetByName('Unconstrained Wide');
  
  console.log(`Constrained Wide exists: ${!!constrainedSheet}`);
  console.log(`Unconstrained Wide exists: ${!!unconstrainedSheet}`);
  
  if (constrainedSheet) {
    console.log(`Constrained data: ${constrainedSheet.getLastRow()} rows x ${constrainedSheet.getLastColumn()} columns`);
  }
  
  if (unconstrainedSheet) {
    console.log(`Unconstrained data: ${unconstrainedSheet.getLastRow()} rows x ${unconstrainedSheet.getLastColumn()} columns`);
  }
}

/**
 * Clean up function to remove old sheets
 */
function cleanupOldSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ['Looker_Ready_View_New', 'SKU_Summary_New', 'Customer_Summary_New', 'Weekly_Trends_New'];
  
  sheetsToDelete.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      ss.deleteSheet(sheet);
      console.log(`Deleted sheet: ${sheetName}`);
    }
  });
}
