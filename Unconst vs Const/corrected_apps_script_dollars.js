/**
 * ANKER FORECAST AUTOMATION - CORRECTED FOR DOLLAR FORECASTS
 * IMPORTANT: Forecast data is in DOLLARS, not units
 * Updated to handle revenue forecasts and calculate units = revenue / sell-in price
 */

function transformForecastData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Log all available sheet names for debugging
  const allSheets = ss.getSheets();
  console.log('Available sheets:');
  allSheets.forEach(sheet => console.log(`  - "${sheet.getName()}"`));
  
  // Get source sheets with CORRECT FRESH DATA names
  const constrainedSheet = ss.getSheetByName('NEW WIDE CONSTRAINED FCST DATA');
  const unconstrainedSheet = ss.getSheetByName('NEW WIDE UNCONST. FCST DATA');
  
  // Check if sheets exist
  if (!constrainedSheet) {
    throw new Error('Could not find "NEW WIDE CONSTRAINED FCST DATA" sheet. Available sheets: ' + 
                   allSheets.map(s => s.getName()).join(', '));
  }
  if (!unconstrainedSheet) {
    throw new Error('Could not find "NEW WIDE UNCONST. FCST DATA" sheet. Available sheets: ' + 
                   allSheets.map(s => s.getName()).join(', '));
  }
  
  // Create or clear output sheet
  let outputSheet = ss.getSheetByName('Looker_Ready_View_Week33');
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet('Looker_Ready_View_Week33');
  }
  
  // Headers for output - CORRECTED for dollar forecasts
  const headers = ['Customer', 'Anker SKU', 'PDT', 'Forecast Type', 'Quarter', 'Week', 
                   'Forecast - Units', 'Forecast Revenue', 'Delta Units', 'Delta - Revenue', 
                   'Gap Flag', 'IsCurrentQ', 'Important Helper', 'Sell-In Price'];
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get data ranges
  const constrainedData = constrainedSheet.getDataRange().getValues();
  const unconstrainedData = unconstrainedSheet.getDataRange().getValues();
  
  console.log(`Constrained data: ${constrainedData.length} rows x ${constrainedData[0].length} columns`);
  console.log(`Unconstrained data: ${unconstrainedData.length} rows x ${unconstrainedData[0].length} columns`);
  
  // Find week columns (they start with 202534 and are integers in the header)
  const weekColumns = [];
  const headerRow = constrainedData[0];
  
  for (let i = 0; i < headerRow.length; i++) {
    const cellValue = headerRow[i];
    if (cellValue && typeof cellValue === 'number' && cellValue.toString().match(/^202\d\d\d$/)) {
      weekColumns.push({index: i, week: cellValue});
    }
  }
  
  console.log(`Found ${weekColumns.length} week columns:`, weekColumns.map(w => w.week));
  
  if (weekColumns.length === 0) {
    throw new Error('No week columns found. Expected integer columns with format 202xxx');
  }
  
  // Create unconstrained lookup using Important Helper
  // IMPORTANT: Forecast data is in DOLLARS
  const unconstrainedLookup = {};
  for (let i = 1; i < unconstrainedData.length; i++) {
    const row = unconstrainedData[i];
    const helper = row[0]; // Important Helper column (index 0)
    
    if (!helper) continue; // Skip empty rows
    
    weekColumns.forEach(weekCol => {
      const week = weekCol.week;
      const forecastRevenue = parseFloat(row[weekCol.index]) || 0; // This is in DOLLARS
      const sellInPrice = parseFloat(row[2]) || 0; // Index 2 for "Sell-in Price"
      
      // Calculate units from revenue
      const forecastUnits = sellInPrice > 0 ? forecastRevenue / sellInPrice : 0;
      
      const key = `${helper}_${week}`;
      unconstrainedLookup[key] = {
        units: forecastUnits,
        revenue: forecastRevenue
      };
    });
  }
  
  console.log(`Created unconstrained lookup with ${Object.keys(unconstrainedLookup).length} entries`);
  
  const outputData = [];
  
  // Process constrained data
  // IMPORTANT: Forecast data is in DOLLARS
  for (let i = 1; i < constrainedData.length; i++) {
    const row = constrainedData[i];
    const helper = row[0];           // Important Helper
    const shiftLogic = row[1];       // Shift logic
    const sellInPrice = parseFloat(row[2]) || 0;  // Sell-in price
    const category = row[3];         // Category
    const pct = row[4];              // PCT
    const bg = row[5];               // BG
    const pdt = row[6];              // PDT
    const customerId = row[7];       // Customer ID
    const customer = row[8];         // Customer
    const sku = row[9];              // Anker SKU
    
    if (!helper || !customer || !sku) continue; // Skip incomplete rows
    
    weekColumns.forEach(weekCol => {
      const week = weekCol.week;
      const constrainedRevenue = parseFloat(row[weekCol.index]) || 0; // This is in DOLLARS
      const constrainedUnits = sellInPrice > 0 ? constrainedRevenue / sellInPrice : 0; // Calculate units
      
      // Look up unconstrained data using Important Helper
      const key = `${helper}_${week}`;
      const unconstrained = unconstrainedLookup[key] || { units: 0, revenue: 0 };
      
      const deltaUnits = constrainedUnits - unconstrained.units;
      const deltaRevenue = constrainedRevenue - unconstrained.revenue;
      
      const quarter = getQuarter(week);
      const isCurrentQ = quarter === 'Q4 2025'; // Adjust current quarter as needed
      const gapFlag = deltaUnits < 0 ? 'Supply Gap' : '';
      
      // Add constrained row
      outputData.push([
        customer,
        sku,
        pdt,
        'Constrained',
        quarter,
        week,
        constrainedUnits,      // Calculated from revenue / price
        constrainedRevenue,    // Direct from forecast data
        deltaUnits,            // Difference in calculated units
        deltaRevenue,          // Difference in revenue
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
        week,
        unconstrained.units,   // Calculated from revenue / price
        unconstrained.revenue, // Direct from forecast data
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
  console.log(`NOTE: Forecast data was in DOLLARS and converted to units using sell-in price`);
  
  // Create summary sheets
  createSummarySheets(ss);
  
  return outputData.length;
}

function getQuarter(week) {
  // Updated for Week 33 data (starts at 202534)
  if (week >= 202534 && week <= 202539) return 'Q3 2025';
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
    if (col <= numColumns && sheet.getLastRow() > 1) {
      const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
      range.setNumberFormat('$#,##0.00');
    }
  });
  
  // Format number columns (Units columns) - now calculated from revenue
  const unitColumns = [7, 9]; // Forecast Units, Delta Units
  unitColumns.forEach(col => {
    if (col <= numColumns && sheet.getLastRow() > 1) {
      const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
      range.setNumberFormat('#,##0.0'); // Show decimal places for calculated units
    }
  });
  
  // Conditional formatting for Gap Flag
  if (sheet.getLastRow() > 1) {
    const gapFlagRange = sheet.getRange(2, 11, sheet.getLastRow() - 1, 1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Supply Gap')
      .setBackground('#FFE6E6')
      .setFontColor('#CC0000')
      .setRanges([gapFlagRange])
      .build();
    sheet.setConditionalFormatRules([rule]);
  }
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
  let sheet = ss.getSheetByName('SKU_Summary_Week33');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('SKU_Summary_Week33');
  }
  
  // Headers
  const headers = ['Instructions', 'Create Pivot Table from Looker_Ready_View_Week33'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const instructions = [
    ['IMPORTANT: Forecasts were in DOLLARS, units calculated as Revenue/Price', ''],
    ['Create Pivot Table:', ''],
    ['1. Data Source:', 'Looker_Ready_View_Week33'],
    ['2. Filter by:', 'Gap Flag = "Supply Gap"'],
    ['3. Rows:', 'Anker SKU, PDT'],
    ['4. Values:', 'Sum of Delta Units, Sum of Delta Revenue, Count of Customer'],
    ['5. Sort by:', 'Delta Revenue (descending)'],
    ['Note:', 'Units are calculated from Revenue ÷ Sell-in Price']
  ];
  
  sheet.getRange(2, 1, instructions.length, 2).setValues(instructions);
  formatOutputSheet(sheet, 2);
}

function createCustomerSummary(ss) {
  let sheet = ss.getSheetByName('Customer_Summary_Week33');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Customer_Summary_Week33');
  }
  
  const headers = ['Instructions', 'Create Pivot Table from Looker_Ready_View_Week33'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const instructions = [
    ['IMPORTANT: Forecasts were in DOLLARS, units calculated as Revenue/Price', ''],
    ['Create Pivot Table:', ''],
    ['1. Data Source:', 'Looker_Ready_View_Week33'],
    ['2. Filter by:', 'Gap Flag = "Supply Gap"'],
    ['3. Rows:', 'Customer'],
    ['4. Values:', 'Sum of Delta Units, Sum of Delta Revenue, Count of Anker SKU'],
    ['5. Sort by:', 'Delta Revenue (descending)'],
    ['Note:', 'Units are calculated from Revenue ÷ Sell-in Price']
  ];
  
  sheet.getRange(2, 1, instructions.length, 2).setValues(instructions);
  formatOutputSheet(sheet, 2);
}

function createWeeklyTrends(ss) {
  let sheet = ss.getSheetByName('Weekly_Trends_Week33');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Weekly_Trends_Week33');
  }
  
  const headers = ['Instructions', 'Create Pivot Table from Looker_Ready_View_Week33'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const instructions = [
    ['IMPORTANT: Forecasts were in DOLLARS, units calculated as Revenue/Price', ''],
    ['Create Pivot Table:', ''],
    ['1. Data Source:', 'Looker_Ready_View_Week33'],
    ['2. Filter by:', 'Gap Flag = "Supply Gap"'],
    ['3. Rows:', 'Week, Quarter'],
    ['4. Values:', 'Sum of Delta Units, Sum of Delta Revenue, Count of records'],
    ['5. Sort by:', 'Week (ascending)'],
    ['Note:', 'Units are calculated from Revenue ÷ Sell-in Price']
  ];
  
  sheet.getRange(2, 1, instructions.length, 2).setValues(instructions);
  formatOutputSheet(sheet, 2);
}

/**
 * Test function to verify sheet access and data format
 */
function testSheetAccessWeek33() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('Testing Week 33 sheet access and data format...');
  
  const constrainedSheet = ss.getSheetByName('NEW WIDE CONSTRAINED FCST DATA');
  
  if (constrainedSheet) {
    // Check sample data to confirm it's in dollars
    const sampleData = constrainedSheet.getRange(2, 1, 5, 15).getValues();
    console.log('Sample forecast data (confirming dollar format):');
    
    sampleData.forEach((row, index) => {
      const helper = row[0];
      const customer = row[8];
      const sku = row[9];
      const sellPrice = row[2];
      const forecast202537 = row[13]; // Assuming 202537 is around column 13
      
      if (forecast202537 && forecast202537 > 0) {
        const impliedUnits = sellPrice > 0 ? forecast202537 / sellPrice : 0;
        console.log(`Row ${index + 2}: ${customer} | ${sku} | Forecast: $${forecast202537} | Price: $${sellPrice} | Implied Units: ${impliedUnits.toFixed(1)}`);
      }
    });
  }
}

/**
 * Complete automation for Week 33 data with dollar forecasts
 */
function runCompleteAutomationWeek33() {
  console.log('🚀 Starting Week 33 forecast automation (DOLLAR FORECASTS)...');
  
  try {
    // Step 1: Transform data
    console.log('Step 1: Transforming Week 33 forecast data (converting dollars to units)...');
    const recordCount = transformForecastData();
    
    console.log(`✅ Week 33 automation finished successfully!`);
    console.log(`📊 Processed ${recordCount} records`);
    console.log(`💡 NOTE: Forecast data was in DOLLARS and converted to units using sell-in prices`);
    console.log('🎯 Navigate to Looker_Ready_View_Week33 sheet to view results');
    
    // Show completion message
    SpreadsheetApp.getUi().alert(
      'Week 33 Automation Complete!', 
      `Successfully processed ${recordCount} records from fresh Week 33 data.\n\nIMPORTANT: Forecast data was in DOLLARS and converted to units using:\nUnits = Revenue ÷ Sell-in Price\n\nNavigate to Looker_Ready_View_Week33 to view results.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ Week 33 automation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Week 33 Automation Error', 
      `Error: ${error.message}\n\nMake sure you're using the NEW EXCEL.xlsx file with the fresh Week 33 data sheets.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Menu setup function for Week 33
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Week 33 Automation (Dollar Forecasts)')
    .addItem('🚀 Run Week 33 Complete Automation', 'runCompleteAutomationWeek33')
    .addSeparator()
    .addItem('📊 Transform Week 33 Data Only', 'transformForecastData')
    .addSeparator()
    .addItem('🧪 Test Week 33 Sheet Access', 'testSheetAccessWeek33')
    .addToUi();
}
