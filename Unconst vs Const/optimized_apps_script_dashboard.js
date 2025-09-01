/**
 * OPTIMIZED ANKER FORECAST AUTOMATION - BATCH PROCESSING
 * Handles large datasets without timeouts using batch processing
 * Creates manager dashboard with fresh Week 33 data
 */

function runCompleteManagerAutomation() {
  console.log('🚀 Starting optimized manager automation...');
  
  try {
    // Step 1: Transform data in batches
    const recordCount = transformForecastDataOptimized();
    
    // Step 2: Create dashboard (lightweight)
    createManagerDashboardOptimized();
    
    console.log(`✅ Optimized automation complete!`);
    console.log(`📊 Processed ${recordCount} records`);
    
    SpreadsheetApp.getUi().alert(
      'Manager Dashboard Complete!', 
      `✅ Successfully processed ${recordCount} records\n📊 Created Dashboard_Week33_Optimized\n🔋 All sheets ready for your manager\n\n⚠️ Large dataset processed in batches to avoid timeouts`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ Optimized automation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Automation Error', 
      `Error: ${error.message}\n\nTry running "Transform Data Only" first, then "Create Dashboard Only"`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

function transformForecastDataOptimized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source sheets
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
  
  // Headers
  const headers = ['Customer', 'Anker SKU', 'PDT', 'Forecast Type', 'Quarter', 'Week', 
                   'Forecast - Units', 'Forecast Revenue', 'Delta Units', 'Delta - Revenue', 
                   'Gap Flag', 'IsCurrentQ', 'Important Helper', 'Sell-In Price'];
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get data ranges
  const constrainedData = constrainedSheet.getDataRange().getValues();
  const unconstrainedData = unconstrainedSheet.getDataRange().getValues();
  
  console.log(`Processing ${constrainedData.length} constrained rows, ${unconstrainedData.length} unconstrained rows`);
  
  // Find week columns
  const weekColumns = [];
  const headerRow = constrainedData[0];
  
  for (let i = 0; i < headerRow.length; i++) {
    const cellValue = headerRow[i];
    if (cellValue && typeof cellValue === 'number' && cellValue.toString().match(/^202\d\d\d$/)) {
      weekColumns.push({index: i, week: cellValue});
    }
  }
  
  console.log(`Found ${weekColumns.length} week columns`);
  
  // Create unconstrained lookup - OPTIMIZED
  const unconstrainedLookup = {};
  console.log('Building unconstrained lookup...');
  
  for (let i = 1; i < unconstrainedData.length; i++) {
    if (i % 100 === 0) {
      console.log(`Processing unconstrained row ${i}/${unconstrainedData.length}`);
      Utilities.sleep(10); // Brief pause every 100 rows
    }
    
    const row = unconstrainedData[i];
    const helper = row[0];
    if (!helper) continue;
    
    const sellInPrice = parseFloat(row[2]) || 0;
    
    weekColumns.forEach(weekCol => {
      const week = weekCol.week;
      const forecastRevenue = parseFloat(row[weekCol.index]) || 0;
      const forecastUnits = sellInPrice > 0 ? forecastRevenue / sellInPrice : 0;
      
      const key = `${helper}_${week}`;
      unconstrainedLookup[key] = {
        units: forecastUnits,
        revenue: forecastRevenue
      };
    });
  }
  
  console.log(`Created lookup with ${Object.keys(unconstrainedLookup).length} entries`);
  
  // Process constrained data in BATCHES
  const BATCH_SIZE = 50; // Process 50 rows at a time
  const allOutputData = [];
  let totalRecords = 0;
  
  for (let startRow = 1; startRow < constrainedData.length; startRow += BATCH_SIZE) {
    const endRow = Math.min(startRow + BATCH_SIZE, constrainedData.length);
    console.log(`Processing constrained batch: rows ${startRow} to ${endRow}`);
    
    const batchData = [];
    
    for (let i = startRow; i < endRow; i++) {
      const row = constrainedData[i];
      const helper = row[0];
      const sellInPrice = parseFloat(row[2]) || 0;
      const pdt = row[6];
      const customer = row[8];
      const sku = row[9];
      
      if (!helper || !customer || !sku) continue;
      
      weekColumns.forEach(weekCol => {
        const week = weekCol.week;
        const constrainedRevenue = parseFloat(row[weekCol.index]) || 0;
        const constrainedUnits = sellInPrice > 0 ? constrainedRevenue / sellInPrice : 0;
        
        const key = `${helper}_${week}`;
        const unconstrained = unconstrainedLookup[key] || { units: 0, revenue: 0 };
        
        const deltaUnits = constrainedUnits - unconstrained.units;
        const deltaRevenue = constrainedRevenue - unconstrained.revenue;
        
        const quarter = getQuarter(week);
        const isCurrentQ = quarter === 'Q4 2025';
        const gapFlag = deltaUnits < 0 ? 'Supply Gap' : '';
        
        // Only add records with actual forecast data to reduce size
        if (constrainedRevenue > 0 || unconstrained.revenue > 0) {
          // Add constrained row
          batchData.push([
            customer, sku, pdt, 'Constrained', quarter, week,
            constrainedUnits, constrainedRevenue, deltaUnits, deltaRevenue,
            gapFlag, isCurrentQ, helper, sellInPrice
          ]);
          
          // Add unconstrained row
          batchData.push([
            customer, sku, pdt, 'Unconstrained', quarter, week,
            unconstrained.units, unconstrained.revenue, 0, 0,
            '', isCurrentQ, helper, sellInPrice
          ]);
        }
      });
    }
    
    // Write batch to sheet
    if (batchData.length > 0) {
      const startOutputRow = allOutputData.length + 2; // +2 for header
      outputSheet.getRange(startOutputRow, 1, batchData.length, headers.length).setValues(batchData);
      allOutputData.push(...batchData);
      totalRecords += batchData.length;
      
      console.log(`Wrote batch: ${batchData.length} records (total: ${totalRecords})`);
    }
    
    // Brief pause between batches
    Utilities.sleep(100);
  }
  
  // Format the sheet
  formatOutputSheetOptimized(outputSheet, headers.length);
  
  console.log(`✅ Optimized transformation complete! ${totalRecords} records created.`);
  return totalRecords;
}

function createManagerDashboardOptimized() {
  console.log('📊 Creating optimized manager dashboard...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create lightweight dashboard
  let dashboardSheet = ss.getSheetByName('Dashboard_Week33_Optimized');
  if (dashboardSheet) {
    dashboardSheet.clear();
  } else {
    dashboardSheet = ss.insertSheet('Dashboard_Week33_Optimized');
  }
  
  // Title
  dashboardSheet.getRange('A1').setValue('ANKER SUPPLY GAP ANALYSIS - WEEK 33 (OPTIMIZED)');
  dashboardSheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  dashboardSheet.getRange('A1:H1').merge();
  
  // Timestamp
  const timestamp = new Date().toLocaleString();
  dashboardSheet.getRange('A2').setValue(`Last Updated: ${timestamp} | Data: Fresh Week 33 (Optimized Processing)`);
  dashboardSheet.getRange('A2').setFontStyle('italic');
  dashboardSheet.getRange('A2:H2').merge();
  
  // Key metrics
  createKeyMetricsOptimized(dashboardSheet);
  
  // Instructions for analysis
  createAnalysisInstructions(dashboardSheet);
  
  console.log('✅ Optimized dashboard created');
}

function createKeyMetricsOptimized(sheet) {
  const metricsStart = 4;
  sheet.getRange(`A${metricsStart}`).setValue('KEY METRICS (Auto-Calculated)');
  sheet.getRange(`A${metricsStart}`).setFontSize(14).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  sheet.getRange(`A${metricsStart}:C${metricsStart}`).merge();
  
  const metrics = [
    ['Metric', 'Formula/Value', 'Description'],
    ['Total Supply Gaps', '=COUNTIFS(Looker_Ready_View_Week33!K:K,"Supply Gap")', 'Number of SKU-Customer-Week gaps'],
    ['Revenue at Risk', '=SUMIFS(Looker_Ready_View_Week33!J:J,Looker_Ready_View_Week33!K:K,"Supply Gap")*-1', 'Total revenue impact from gaps'],
    ['Units at Risk', '=SUMIFS(Looker_Ready_View_Week33!I:I,Looker_Ready_View_Week33!K:K,"Supply Gap")*-1', 'Total units affected'],
    ['Q3 2025 Impact', '=SUMIFS(Looker_Ready_View_Week33!J:J,Looker_Ready_View_Week33!K:K,"Supply Gap",Looker_Ready_View_Week33!E:E,"Q3 2025")*-1', 'Q3 revenue at risk'],
    ['Q4 2025 Impact', '=SUMIFS(Looker_Ready_View_Week33!J:J,Looker_Ready_View_Week33!K:K,"Supply Gap",Looker_Ready_View_Week33!E:E,"Q4 2025")*-1', 'Q4 revenue at risk'],
    ['', '', ''],
    ['Data Quality Check:', '', ''],
    ['Total Records', '=COUNTA(Looker_Ready_View_Week33!A:A)-1', 'Should be 20,000+ records'],
    ['Week Range', 'Week 34-53 (202534-202553)', 'Fresh Week 33 data'],
    ['Forecast Type', 'DOLLARS converted to units', 'Revenue ÷ Sell-in Price']
  ];
  
  const metricsRange = sheet.getRange(metricsStart + 1, 1, metrics.length, 3);
  metricsRange.setValues(metrics);
  
  // Format headers
  sheet.getRange(metricsStart + 1, 1, 1, 3).setBackground('#E8F0FE').setFontWeight('bold');
  
  // Format currency values (rows 3-6)
  sheet.getRange(metricsStart + 3, 2, 4, 1).setNumberFormat('$#,##0');
}

function createAnalysisInstructions(sheet) {
  const instructStart = 20;
  
  sheet.getRange(`A${instructStart}`).setValue('ANALYSIS INSTRUCTIONS FOR YOUR MANAGER');
  sheet.getRange(`A${instructStart}`).setFontWeight('bold').setBackground('#FF9900').setFontColor('#FFFFFF');
  sheet.getRange(`A${instructStart}:H${instructStart}`).merge();
  
  const instructions = [
    ['Analysis Type', 'Instructions', 'Purpose'],
    ['', '', ''],
    ['SKU Gap Analysis', '1. Go to Data → Pivot Table', 'Find top SKUs with supply gaps'],
    ['', '2. Source: Looker_Ready_View_Week33', ''],
    ['', '3. Filter: Gap Flag = "Supply Gap"', ''],
    ['', '4. Rows: Anker SKU, PDT', ''],
    ['', '5. Values: Sum of Delta Units, Sum of Delta Revenue', ''],
    ['', '6. Sort by Delta Revenue (descending)', ''],
    ['', '', ''],
    ['Customer Impact', '1. Create new Pivot Table', 'Revenue at risk by customer'],
    ['', '2. Filter: Gap Flag = "Supply Gap"', ''],
    ['', '3. Rows: Customer', ''],
    ['', '4. Values: Sum of Delta Revenue, Count of Anker SKU', ''],
    ['', '', ''],
    ['Weekly Trends', '1. Create Pivot Table', 'Timeline of supply gaps'],
    ['', '2. Rows: Week, Quarter', ''],
    ['', '3. Values: Sum of Delta Revenue', ''],
    ['', '4. Create line chart for visualization', ''],
    ['', '', ''],
    ['Battery Focus', '1. Filter data for PDT = "Battery"', 'Battery-specific analysis'],
    ['', '2. Review supply status by SKU', ''],
    ['', '3. Focus on highest revenue impact items', '']
  ];
  
  sheet.getRange(instructStart + 1, 1, instructions.length, 3).setValues(instructions);
  sheet.getRange(instructStart + 1, 1, 1, 3).setBackground('#FFF2CC').setFontWeight('bold');
}

function formatOutputSheetOptimized(sheet, numColumns) {
  console.log('Formatting output sheet...');
  
  // Basic formatting only to avoid timeouts
  sheet.setFrozenRows(1);
  
  const headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // Auto-resize key columns only
  for (let i = 1; i <= Math.min(6, numColumns); i++) {
    sheet.autoResizeColumn(i);
  }
  
  console.log('Basic formatting complete');
}

function getQuarter(week) {
  if (week >= 202534 && week <= 202539) return 'Q3 2025';
  if (week >= 202540 && week <= 202552) return 'Q4 2025';
  if (week >= 202553 && week <= 202605) return 'Q1 2026';
  return 'Unknown';
}

/**
 * Lighter weight functions for step-by-step execution
 */
function transformDataOnly() {
  console.log('📊 Running data transformation only...');
  const recordCount = transformForecastDataOptimized();
  
  SpreadsheetApp.getUi().alert(
    'Data Transformation Complete!', 
    `✅ Successfully processed ${recordCount} records\n📊 Data ready in Looker_Ready_View_Week33\n\nNext: Run "Create Dashboard Only" if needed`, 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createDashboardOnly() {
  console.log('📈 Creating dashboard only...');
  createManagerDashboardOptimized();
  
  SpreadsheetApp.getUi().alert(
    'Dashboard Created!', 
    `📊 Dashboard_Week33_Optimized ready\n📈 Key metrics and analysis instructions included\n\n✅ Ready for your manager!`, 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Menu setup
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 Optimized Manager Dashboard')
    .addItem('🚀 Run Complete Automation (Optimized)', 'runCompleteManagerAutomation')
    .addSeparator()
    .addItem('📊 Step 1: Transform Data Only', 'transformDataOnly')
    .addItem('📈 Step 2: Create Dashboard Only', 'createDashboardOnly')
    .addToUi();
}
