/**
 * CORRECTED ANKER FORECAST AUTOMATION - WEEK 33
 * Fixed to match your exact data structure based on diagnostic results
 */

function runCompleteManagerAutomationCorrected() {
  console.log('🚀 Starting CORRECTED manager automation...');
  
  try {
    // Step 1: Transform data in batches
    const recordCount = transformForecastDataCorrected();
    
    // Step 2: Create dashboard
    createManagerDashboardCorrected();
    
    console.log(`✅ CORRECTED automation complete!`);
    console.log(`📊 Processed ${recordCount} records`);
    
    SpreadsheetApp.getUi().alert(
      'CORRECTED Manager Dashboard Complete!', 
      `✅ Successfully processed ${recordCount} records\n📊 Created Looker_Ready_View_Week33_CORRECTED\n📊 Created Dashboard_Week33_CORRECTED\n\n🔧 Fixed to match your exact data structure!`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ CORRECTED automation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Automation Error', 
      `Error: ${error.message}\n\nCheck the execution transcript for details.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

function transformForecastDataCorrected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source sheets
  const constrainedSheet = ss.getSheetByName('NEW WIDE CONSTRAINED FCST DATA');
  const unconstrainedSheet = ss.getSheetByName('NEW WIDE UNCONST. FCST DATA');
  
  if (!constrainedSheet || !unconstrainedSheet) {
    throw new Error('Could not find source sheets');
  }
  
  // Create or clear output sheet
  let outputSheet = ss.getSheetByName('Looker_Ready_View_Week33_CORRECTED');
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet('Looker_Ready_View_Week33_CORRECTED');
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
  
  // Map columns based on your exact structure
  const COLUMN_MAP = {
    HELPER: 0,        // Col 1: Important Helper
    SHIFT_LOGIC: 1,   // Col 2: Shift logic  
    SELL_IN_PRICE: 2, // Col 3: Sell-in price
    CATEGORY: 3,      // Col 4: Category
    PCT: 4,           // Col 5: PCT
    BG: 5,            // Col 6: BG
    PDT: 6,           // Col 7: PDT
    CUSTOMER_ID: 7,   // Col 8: Customer ID
    CUSTOMER: 8,      // Col 9: Customer
    ANKER_SKU: 9,     // Col 10: Anker SKU
    WEEK_START: 10    // Col 11: First week (202534)
  };
  
  // Find week columns - using your exact structure
  const weekColumns = [];
  const headerRow = constrainedData[0];
  
  // Based on diagnostic: weeks are in columns 11-29 (indices 10-28)
  for (let i = COLUMN_MAP.WEEK_START; i < headerRow.length; i++) {
    const cellValue = headerRow[i];
    if (cellValue && typeof cellValue === 'number' && cellValue.toString().match(/^202\d\d\d$/)) {
      weekColumns.push({index: i, week: cellValue});
    }
  }
  
  console.log(`Found ${weekColumns.length} week columns: ${weekColumns.map(w => w.week).join(', ')}`);
  
  if (weekColumns.length === 0) {
    throw new Error('No week columns found!');
  }
  
  // Create unconstrained lookup
  const unconstrainedLookup = {};
  console.log('Building unconstrained lookup...');
  
  for (let i = 1; i < unconstrainedData.length; i++) {
    if (i % 100 === 0) {
      console.log(`Processing unconstrained row ${i}/${unconstrainedData.length}`);
      Utilities.sleep(10);
    }
    
    const row = unconstrainedData[i];
    const helper = row[COLUMN_MAP.HELPER];
    if (!helper) continue;
    
    const sellInPrice = parseFloat(row[COLUMN_MAP.SELL_IN_PRICE]) || 0;
    
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
  const BATCH_SIZE = 100;
  let totalRecords = 0;
  
  for (let startRow = 1; startRow < constrainedData.length; startRow += BATCH_SIZE) {
    const endRow = Math.min(startRow + BATCH_SIZE, constrainedData.length);
    console.log(`Processing constrained batch: rows ${startRow} to ${endRow}`);
    
    const batchData = [];
    
    for (let i = startRow; i < endRow; i++) {
      const row = constrainedData[i];
      const helper = row[COLUMN_MAP.HELPER];
      const sellInPrice = parseFloat(row[COLUMN_MAP.SELL_IN_PRICE]) || 0;
      const pdt = row[COLUMN_MAP.PDT];
      const customer = row[COLUMN_MAP.CUSTOMER];
      const sku = row[COLUMN_MAP.ANKER_SKU];
      
      if (!helper || !customer || !sku) continue;
      
      weekColumns.forEach(weekCol => {
        const week = weekCol.week;
        const constrainedRevenue = parseFloat(row[weekCol.index]) || 0;
        const constrainedUnits = sellInPrice > 0 ? constrainedRevenue / sellInPrice : 0;
        
        const key = `${helper}_${week}`;
        const unconstrained = unconstrainedLookup[key] || { units: 0, revenue: 0 };
        
        const deltaUnits = constrainedUnits - unconstrained.units;
        const deltaRevenue = constrainedRevenue - unconstrained.revenue;
        
        const quarter = getQuarterCorrected(week);
        const isCurrentQ = quarter === 'Q4 2025';
        const gapFlag = deltaUnits < 0 ? 'Supply Gap' : '';
        
        // Add records with any forecast data
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
      const startOutputRow = totalRecords + 2; // +2 for header and 1-based indexing
      outputSheet.getRange(startOutputRow, 1, batchData.length, headers.length).setValues(batchData);
      totalRecords += batchData.length;
      
      console.log(`Wrote batch: ${batchData.length} records (total: ${totalRecords})`);
    }
    
    // Brief pause between batches
    Utilities.sleep(100);
  }
  
  // Format the sheet
  formatOutputSheetCorrected(outputSheet, headers.length);
  
  console.log(`✅ CORRECTED transformation complete! ${totalRecords} records created.`);
  return totalRecords;
}

function createManagerDashboardCorrected() {
  console.log('📊 Creating CORRECTED manager dashboard...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create dashboard
  let dashboardSheet = ss.getSheetByName('Dashboard_Week33_CORRECTED');
  if (dashboardSheet) {
    dashboardSheet.clear();
  } else {
    dashboardSheet = ss.insertSheet('Dashboard_Week33_CORRECTED');
  }
  
  // Title
  dashboardSheet.getRange('A1').setValue('ANKER SUPPLY GAP ANALYSIS - FRESH WEEK 33 DATA (CORRECTED)');
  dashboardSheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  dashboardSheet.getRange('A1:H1').merge();
  
  // Timestamp and data info
  const timestamp = new Date().toLocaleString();
  dashboardSheet.getRange('A2').setValue(`Generated: ${timestamp} | Source: NEW WIDE CONSTRAINED/UNCONSTRAINED FCST DATA`);
  dashboardSheet.getRange('A2').setFontStyle('italic');
  dashboardSheet.getRange('A2:H2').merge();
  
  dashboardSheet.getRange('A3').setValue('Data Range: Week 34-52 (202534-202552) | Quarters: Q3 2025, Q4 2025');
  dashboardSheet.getRange('A3:H3').merge();
  
  // Key metrics
  createKeyMetricsCorrected(dashboardSheet);
  
  // Analysis instructions
  createAnalysisInstructionsCorrected(dashboardSheet);
  
  console.log('✅ CORRECTED dashboard created');
}

function createKeyMetricsCorrected(sheet) {
  const metricsStart = 5;
  sheet.getRange(`A${metricsStart}`).setValue('KEY METRICS - FRESH WEEK 33');
  sheet.getRange(`A${metricsStart}`).setFontSize(14).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  sheet.getRange(`A${metricsStart}:C${metricsStart}`).merge();
  
  const metrics = [
    ['Metric', 'Formula/Value', 'Description'],
    ['Total Supply Gaps', '=COUNTIFS(Looker_Ready_View_Week33_CORRECTED!K:K,"Supply Gap")', 'Number of supply gap scenarios'],
    ['Revenue at Risk', '=SUMIFS(Looker_Ready_View_Week33_CORRECTED!J:J,Looker_Ready_View_Week33_CORRECTED!K:K,"Supply Gap")*-1', 'Total revenue impact'],
    ['Units at Risk', '=SUMIFS(Looker_Ready_View_Week33_CORRECTED!I:I,Looker_Ready_View_Week33_CORRECTED!K:K,"Supply Gap")*-1', 'Total units affected'],
    ['Q3 2025 Impact', '=SUMIFS(Looker_Ready_View_Week33_CORRECTED!J:J,Looker_Ready_View_Week33_CORRECTED!K:K,"Supply Gap",Looker_Ready_View_Week33_CORRECTED!E:E,"Q3 2025")*-1', 'Q3 revenue at risk'],
    ['Q4 2025 Impact', '=SUMIFS(Looker_Ready_View_Week33_CORRECTED!J:J,Looker_Ready_View_Week33_CORRECTED!K:K,"Supply Gap",Looker_Ready_View_Week33_CORRECTED!E:E,"Q4 2025")*-1', 'Q4 revenue at risk'],
    ['', '', ''],
    ['Data Verification:', '', ''],
    ['Total Records', '=COUNTA(Looker_Ready_View_Week33_CORRECTED!A:A)-1', 'Should be 15,000+ records'],
    ['Week Range Check', 'Week 34-52 (202534-202552)', '19 weeks of forecast data'],
    ['Data Source Check', '662 rows x 19 weeks = ~25,000 potential records', 'From NEW WIDE data sheets']
  ];
  
  const metricsRange = sheet.getRange(metricsStart + 1, 1, metrics.length, 3);
  metricsRange.setValues(metrics);
  
  // Format headers
  sheet.getRange(metricsStart + 1, 1, 1, 3).setBackground('#E8F0FE').setFontWeight('bold');
  
  // Format currency values
  sheet.getRange(metricsStart + 3, 2, 3, 1).setNumberFormat('$#,##0');
}

function createAnalysisInstructionsCorrected(sheet) {
  const instructStart = 20;
  
  sheet.getRange(`A${instructStart}`).setValue('ANALYSIS INSTRUCTIONS FOR YOUR MANAGER');
  sheet.getRange(`A${instructStart}`).setFontWeight('bold').setBackground('#FF9900').setFontColor('#FFFFFF');
  sheet.getRange(`A${instructStart}:H${instructStart}`).merge();
  
  const instructions = [
    ['Analysis Type', 'Instructions', 'Purpose'],
    ['', '', ''],
    ['1. SKU Gap Analysis', '• Data → Pivot Table', 'Find top SKUs with supply gaps'],
    ['', '• Source: Looker_Ready_View_Week33_CORRECTED', ''],
    ['', '• Filter: Gap Flag = "Supply Gap"', ''],
    ['', '• Rows: Anker SKU, PDT', ''],
    ['', '• Values: Sum of Delta Revenue (descending)', ''],
    ['', '', ''],
    ['2. Customer Impact', '• Create new Pivot Table', 'Revenue at risk by customer'],
    ['', '• Filter: Gap Flag = "Supply Gap"', ''],
    ['', '• Rows: Customer', ''],
    ['', '• Values: Sum of Delta Revenue, Count of SKUs', ''],
    ['', '', ''],
    ['3. Weekly Timeline', '• Pivot: Rows = Week, Values = Sum Delta Revenue', 'Timeline of supply gaps'],
    ['', '• Create line chart for visualization', ''],
    ['', '', ''],
    ['4. Data Quality Check', '• Total records should be 15,000-25,000', 'Verify complete processing'],
    ['', '• All Gap Flags should show "Supply Gap" or blank', ''],
    ['', '• Week range: 202534-202552 (19 weeks)', '']
  ];
  
  sheet.getRange(instructStart + 1, 1, instructions.length, 3).setValues(instructions);
  sheet.getRange(instructStart + 1, 1, 1, 3).setBackground('#FFF2CC').setFontWeight('bold');
}

function formatOutputSheetCorrected(sheet, numColumns) {
  console.log('Formatting output sheet...');
  
  sheet.setFrozenRows(1);
  
  const headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // Auto-resize key columns
  for (let i = 1; i <= Math.min(8, numColumns); i++) {
    sheet.autoResizeColumn(i);
  }
  
  console.log('Formatting complete');
}

function getQuarterCorrected(week) {
  if (week >= 202534 && week <= 202539) return 'Q3 2025';
  if (week >= 202540 && week <= 202552) return 'Q4 2025';
  if (week >= 202553 && week <= 202605) return 'Q1 2026';
  return 'Unknown';
}

/**
 * Menu setup
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 CORRECTED Week 33 Automation')
    .addItem('🚀 Run CORRECTED Complete Automation', 'runCompleteManagerAutomationCorrected')
    .addToUi();
}


