/**
 * ENHANCED ANKER FORECAST AUTOMATION - DASHBOARD CREATOR
 * Creates the exact dashboard format your manager needs
 * Handles dollar forecasts and automatically populates all dashboard sheets
 */

function transformForecastDataAndCreateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('🚀 Starting enhanced automation with dashboard creation...');
  
  // Step 1: Transform the forecast data
  const recordCount = transformForecastData();
  
  // Step 2: Create the manager dashboard
  createManagerDashboard(ss);
  
  // Step 3: Update battery supply plan
  updateBatterySupplyPlan(ss);
  
  console.log('✅ Enhanced automation complete with manager dashboard!');
  return recordCount;
}

function transformForecastData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source sheets with CORRECT FRESH DATA names
  const constrainedSheet = ss.getSheetByName('NEW WIDE CONSTRAINED FCST DATA');
  const unconstrainedSheet = ss.getSheetByName('NEW WIDE UNCONST. FCST DATA');
  
  if (!constrainedSheet || !unconstrainedSheet) {
    throw new Error('Could not find fresh data sheets. Make sure you have "NEW WIDE CONSTRAINED FCST DATA" and "NEW WIDE UNCONST. FCST DATA" sheets.');
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
  
  console.log(`Processing constrained data: ${constrainedData.length} rows`);
  console.log(`Processing unconstrained data: ${unconstrainedData.length} rows`);
  
  // Find week columns (they start with 202534 and are integers in the header)
  const weekColumns = [];
  const headerRow = constrainedData[0];
  
  for (let i = 0; i < headerRow.length; i++) {
    const cellValue = headerRow[i];
    if (cellValue && typeof cellValue === 'number' && cellValue.toString().match(/^202\d\d\d$/)) {
      weekColumns.push({index: i, week: cellValue});
    }
  }
  
  console.log(`Found ${weekColumns.length} week columns starting with ${weekColumns[0]?.week}`);
  
  // Create unconstrained lookup using Important Helper
  const unconstrainedLookup = {};
  for (let i = 1; i < unconstrainedData.length; i++) {
    const row = unconstrainedData[i];
    const helper = row[0]; // Important Helper column
    
    if (!helper) continue;
    
    weekColumns.forEach(weekCol => {
      const week = weekCol.week;
      const forecastRevenue = parseFloat(row[weekCol.index]) || 0; // DOLLARS
      const sellInPrice = parseFloat(row[2]) || 0; // Sell-in price
      
      const forecastUnits = sellInPrice > 0 ? forecastRevenue / sellInPrice : 0;
      
      const key = `${helper}_${week}`;
      unconstrainedLookup[key] = {
        units: forecastUnits,
        revenue: forecastRevenue
      };
    });
  }
  
  console.log(`Created lookup with ${Object.keys(unconstrainedLookup).length} entries`);
  
  const outputData = [];
  
  // Process constrained data
  for (let i = 1; i < constrainedData.length; i++) {
    const row = constrainedData[i];
    const helper = row[0];           // Important Helper
    const sellInPrice = parseFloat(row[2]) || 0;  // Sell-in price
    const pdt = row[6];              // PDT
    const customer = row[8];         // Customer
    const sku = row[9];              // Anker SKU
    
    if (!helper || !customer || !sku) continue;
    
    weekColumns.forEach(weekCol => {
      const week = weekCol.week;
      const constrainedRevenue = parseFloat(row[weekCol.index]) || 0; // DOLLARS
      const constrainedUnits = sellInPrice > 0 ? constrainedRevenue / sellInPrice : 0;
      
      const key = `${helper}_${week}`;
      const unconstrained = unconstrainedLookup[key] || { units: 0, revenue: 0 };
      
      const deltaUnits = constrainedUnits - unconstrained.units;
      const deltaRevenue = constrainedRevenue - unconstrained.revenue;
      
      const quarter = getQuarter(week);
      const isCurrentQ = quarter === 'Q4 2025';
      const gapFlag = deltaUnits < 0 ? 'Supply Gap' : '';
      
      // Add constrained row
      outputData.push([
        customer, sku, pdt, 'Constrained', quarter, week,
        constrainedUnits, constrainedRevenue, deltaUnits, deltaRevenue,
        gapFlag, isCurrentQ, helper, sellInPrice
      ]);
      
      // Add unconstrained row
      outputData.push([
        customer, sku, pdt, 'Unconstrained', quarter, week,
        unconstrained.units, unconstrained.revenue, 0, 0,
        '', isCurrentQ, helper, sellInPrice
      ]);
    });
  }
  
  // Write output data
  if (outputData.length > 0) {
    outputSheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
    formatOutputSheet(outputSheet, headers.length);
  }
  
  console.log(`✅ Transformation complete! ${outputData.length} records created.`);
  return outputData.length;
}

function createManagerDashboard(ss) {
  console.log('📊 Creating manager dashboard...');
  
  // Create or update Dashboard sheet
  let dashboardSheet = ss.getSheetByName('Dashboard_Week33');
  if (dashboardSheet) {
    dashboardSheet.clear();
  } else {
    dashboardSheet = ss.insertSheet('Dashboard_Week33');
  }
  
  // Dashboard title
  dashboardSheet.getRange('A1').setValue('ANKER SUPPLY GAP ANALYSIS DASHBOARD - WEEK 33');
  dashboardSheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  dashboardSheet.getRange('A1:H1').merge();
  
  // Timestamp
  const timestamp = new Date().toLocaleString();
  dashboardSheet.getRange('A2').setValue(`Last Updated: ${timestamp} | Data: Fresh Week 33`);
  dashboardSheet.getRange('A2').setFontStyle('italic');
  dashboardSheet.getRange('A2:H2').merge();
  
  // Key metrics section
  createKeyMetrics(dashboardSheet);
  
  // Supply gap summary
  createSupplyGapSummary(ss, dashboardSheet);
  
  console.log('✅ Manager dashboard created');
}

function createKeyMetrics(sheet) {
  // Headers for metrics
  const metricsStart = 4;
  sheet.getRange(`A${metricsStart}`).setValue('KEY METRICS');
  sheet.getRange(`A${metricsStart}`).setFontSize(14).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  sheet.getRange(`A${metricsStart}:C${metricsStart}`).merge();
  
  const metrics = [
    ['Metric', 'Value', 'Description'],
    ['Total Supply Gaps', '=COUNTIFS(Looker_Ready_View_Week33!K:K,"Supply Gap")', 'Number of SKU-Customer-Week gaps'],
    ['Revenue at Risk', '=SUMIFS(Looker_Ready_View_Week33!J:J,Looker_Ready_View_Week33!K:K,"Supply Gap")*-1', 'Total revenue impact from gaps'],
    ['Units at Risk', '=SUMIFS(Looker_Ready_View_Week33!I:I,Looker_Ready_View_Week33!K:K,"Supply Gap")*-1', 'Total units affected'],
    ['Q3 2025 Impact', '=SUMIFS(Looker_Ready_View_Week33!J:J,Looker_Ready_View_Week33!K:K,"Supply Gap",Looker_Ready_View_Week33!E:E,"Q3 2025")*-1', 'Q3 revenue at risk'],
    ['Q4 2025 Impact', '=SUMIFS(Looker_Ready_View_Week33!J:J,Looker_Ready_View_Week33!K:K,"Supply Gap",Looker_Ready_View_Week33!E:E,"Q4 2025")*-1', 'Q4 revenue at risk'],
    ['Affected SKUs', '=SUMPRODUCT((Looker_Ready_View_Week33!K:K="Supply Gap")/(COUNTIFS(Looker_Ready_View_Week33!B:B,Looker_Ready_View_Week33!B:B,Looker_Ready_View_Week33!K:K,"Supply Gap")+1))', 'Unique SKUs with gaps'],
    ['Affected Customers', '=SUMPRODUCT((Looker_Ready_View_Week33!K:K="Supply Gap")/(COUNTIFS(Looker_Ready_View_Week33!A:A,Looker_Ready_View_Week33!A:A,Looker_Ready_View_Week33!K:K,"Supply Gap")+1))', 'Unique customers affected']
  ];
  
  const metricsRange = sheet.getRange(metricsStart + 1, 1, metrics.length, 3);
  metricsRange.setValues(metrics);
  
  // Format headers
  sheet.getRange(metricsStart + 1, 1, 1, 3).setBackground('#E8F0FE').setFontWeight('bold');
  
  // Format currency values
  sheet.getRange(metricsStart + 3, 2, 3, 1).setNumberFormat('$#,##0');
  
  // Format number values
  sheet.getRange(metricsStart + 2, 2, 1, 1).setNumberFormat('#,##0');
  sheet.getRange(metricsStart + 4, 2, 3, 1).setNumberFormat('#,##0');
}

function createSupplyGapSummary(ss, dashboardSheet) {
  console.log('📈 Creating supply gap summary...');
  
  // Create Top 10 SKUs by Revenue Impact
  createTopSKUsSummary(ss, dashboardSheet, 15);
  
  // Create Customer Impact Summary  
  createCustomerImpactSummary(ss, dashboardSheet, 15);
  
  // Create Weekly Trends Summary
  createWeeklyTrendsSummary(ss, dashboardSheet, 32);
}

function createTopSKUsSummary(ss, dashboardSheet, startRow) {
  // Title
  dashboardSheet.getRange(`E${startRow}`).setValue('TOP 10 SKUS BY REVENUE IMPACT');
  dashboardSheet.getRange(`E${startRow}`).setFontWeight('bold').setBackground('#FF9900').setFontColor('#FFFFFF');
  dashboardSheet.getRange(`E${startRow}:H${startRow}`).merge();
  
  // Headers
  const headers = ['SKU', 'PDT', 'Gap Units', 'Revenue Impact'];
  dashboardSheet.getRange(startRow + 1, 5, 1, 4).setValues([headers]);
  dashboardSheet.getRange(startRow + 1, 5, 1, 4).setBackground('#FFF2CC').setFontWeight('bold');
  
  // Instructions for manual pivot table creation
  const instructions = [
    ['Create Pivot Table:', '', '', ''],
    ['1. Source: Looker_Ready_View_Week33', '', '', ''],
    ['2. Filter: Gap Flag = "Supply Gap"', '', '', ''],
    ['3. Rows: Anker SKU, PDT', '', '', ''],
    ['4. Values: Sum Delta Units, Sum Delta Revenue', '', '', ''],
    ['5. Sort: Delta Revenue (descending)', '', '', ''],
    ['6. Show only top 10', '', '', '']
  ];
  
  dashboardSheet.getRange(startRow + 2, 5, instructions.length, 4).setValues(instructions);
}

function createCustomerImpactSummary(ss, dashboardSheet, startRow) {
  // Title
  dashboardSheet.getRange(`A${startRow}`).setValue('CUSTOMER IMPACT SUMMARY');
  dashboardSheet.getRange(`A${startRow}`).setFontWeight('bold').setBackground('#9900FF').setFontColor('#FFFFFF');
  dashboardSheet.getRange(`A${startRow}:D${startRow}`).merge();
  
  // Headers
  const headers = ['Customer', 'Gap Units', 'Revenue Impact', 'SKUs Affected'];
  dashboardSheet.getRange(startRow + 1, 1, 1, 4).setValues([headers]);
  dashboardSheet.getRange(startRow + 1, 1, 1, 4).setBackground('#F3E5F5').setFontWeight('bold');
  
  // Instructions
  const instructions = [
    ['Create Pivot Table:', '', '', ''],
    ['1. Source: Looker_Ready_View_Week33', '', '', ''],
    ['2. Filter: Gap Flag = "Supply Gap"', '', '', ''],
    ['3. Rows: Customer', '', '', ''],
    ['4. Values: Sum Delta Units, Sum Delta Revenue, Count SKUs', '', '', ''],
    ['5. Sort: Delta Revenue (descending)', '', '', '']
  ];
  
  dashboardSheet.getRange(startRow + 2, 1, instructions.length, 4).setValues(instructions);
}

function createWeeklyTrendsSummary(ss, dashboardSheet, startRow) {
  // Title
  dashboardSheet.getRange(`A${startRow}`).setValue('WEEKLY SUPPLY GAP TRENDS');
  dashboardSheet.getRange(`A${startRow}`).setFontWeight('bold').setBackground('#0D7377').setFontColor('#FFFFFF');
  dashboardSheet.getRange(`A${startRow}:H${startRow}`).merge();
  
  // Headers
  const headers = ['Week', 'Quarter', 'Gap Units', 'Revenue Impact', 'Records Count'];
  dashboardSheet.getRange(startRow + 1, 1, 1, 5).setValues([headers]);
  dashboardSheet.getRange(startRow + 1, 1, 1, 5).setBackground('#E0F2F1').setFontWeight('bold');
  
  // Instructions
  const instructions = [
    ['Create Pivot Table:', '', '', '', ''],
    ['1. Source: Looker_Ready_View_Week33', '', '', '', ''],
    ['2. Filter: Gap Flag = "Supply Gap"', '', '', '', ''],
    ['3. Rows: Week, Quarter', '', '', '', ''],
    ['4. Values: Sum Delta Units, Sum Delta Revenue, Count', '', '', '', ''],
    ['5. Sort: Week (ascending)', '', '', '', ''],
    ['6. Create line chart for trends', '', '', '', '']
  ];
  
  dashboardSheet.getRange(startRow + 2, 1, instructions.length, 5).setValues(instructions);
}

function updateBatterySupplyPlan(ss) {
  console.log('🔋 Updating battery supply plan...');
  
  let batterySheet = ss.getSheetByName('Battery_Supply_Plan_Week33');
  if (batterySheet) {
    batterySheet.clear();
  } else {
    batterySheet = ss.insertSheet('Battery_Supply_Plan_Week33');
  }
  
  // Title
  batterySheet.getRange('A1').setValue('NA Battery Supply Plan - Week 33');
  batterySheet.getRange('A1').setFontSize(14).setFontWeight('bold').setBackground('#FF6B6B').setFontColor('#FFFFFF');
  batterySheet.getRange('A1:K1').merge();
  
  // Headers matching your existing format
  const headers = [
    'PDT', 'PN', '供应情况打标', 'Supply Status', 'NPI or Existing', 
    'IOQ fully fulfilled', 'Total request (formula)', 'Total gap (formula)', 
    'Q3 supply gap qty', 'Q4 supply gap qty', 'Total supply gap qty', 'Total supply'
  ];
  
  batterySheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  batterySheet.getRange(2, 1, 1, headers.length).setBackground('#FFE066').setFontWeight('bold');
  
  // Instructions for battery data
  const instructions = [
    ['Filter Looker_Ready_View_Week33 for:', '', '', '', '', '', '', '', '', '', '', ''],
    ['1. PDT = "Battery"', '', '', '', '', '', '', '', '', '', '', ''],
    ['2. Gap Flag = "Supply Gap"', '', '', '', '', '', '', '', '', '', '', ''],
    ['3. Group by SKU and sum gaps', '', '', '', '', '', '', '', '', '', '', ''],
    ['4. Add supply status from planners', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', ''],
    ['Sample formulas:', '', '', '', '', '', '', '', '', '', '', ''],
    ['Q3 Gap: =SUMIFS(Looker_Ready_View_Week33!I:I,', '', '', '', '', '', '', '', '', '', '', ''],
    ['    Looker_Ready_View_Week33!C:C,"Battery",', '', '', '', '', '', '', '', '', '', '', ''],
    ['    Looker_Ready_View_Week33!K:K,"Supply Gap",', '', '', '', '', '', '', '', '', '', '', ''],
    ['    Looker_Ready_View_Week33!E:E,"Q3 2025")*-1', '', '', '', '', '', '', '', '', '', '', '']
  ];
  
  batterySheet.getRange(3, 1, instructions.length, headers.length).setValues(instructions);
  
  console.log('✅ Battery supply plan updated');
}

function getQuarter(week) {
  if (week >= 202534 && week <= 202539) return 'Q3 2025';
  if (week >= 202540 && week <= 202552) return 'Q4 2025'; 
  if (week >= 202553 && week <= 202605) return 'Q1 2026';
  return 'Unknown';
}

function formatOutputSheet(sheet, numColumns) {
  sheet.setFrozenRows(1);
  
  const headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  for (let i = 1; i <= numColumns; i++) {
    sheet.autoResizeColumn(i);
  }
  
  if (sheet.getLastRow() > 1) {
    // Format currency columns
    const revenueColumns = [8, 10];
    revenueColumns.forEach(col => {
      if (col <= numColumns) {
        const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
        range.setNumberFormat('$#,##0.00');
      }
    });
    
    // Format unit columns
    const unitColumns = [7, 9];
    unitColumns.forEach(col => {
      if (col <= numColumns) {
        const range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
        range.setNumberFormat('#,##0.0');
      }
    });
    
    // Conditional formatting for gaps
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

/**
 * MAIN FUNCTION - Run this for complete automation with manager dashboard
 */
function runCompleteManagerAutomation() {
  console.log('🚀 Starting complete manager automation...');
  
  try {
    const recordCount = transformForecastDataAndCreateDashboard();
    
    console.log(`✅ Complete manager automation finished!`);
    console.log(`📊 Processed ${recordCount} records`);
    console.log(`📈 Created manager dashboard sheets`);
    
    SpreadsheetApp.getUi().alert(
      'Manager Dashboard Complete!', 
      `✅ Successfully processed ${recordCount} records\n📊 Created Dashboard_Week33 with key metrics\n🔋 Updated Battery_Supply_Plan_Week33\n📈 All sheets ready for your manager\n\nNext: Create pivot tables from the instruction sheets`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ Manager automation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Automation Error', 
      `Error: ${error.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Menu setup
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 Manager Dashboard Automation')
    .addItem('🚀 Run Complete Manager Automation', 'runCompleteManagerAutomation')
    .addSeparator()
    .addItem('📊 Transform Data Only', 'transformForecastData')
    .addItem('📈 Create Dashboard Only', 'createManagerDashboard')
    .addItem('🔋 Update Battery Plan Only', 'updateBatterySupplyPlan')
    .addToUi();
}
