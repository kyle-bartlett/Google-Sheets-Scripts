/**
 * FINAL CLEAN ANKER FORECAST AUTOMATION
 * ONLY uses fresh Week 33 data from specific sheets
 * Prevents any mixing with old data
 */

function runFreshDataAutomation() {
  console.log('🚀 Starting FRESH DATA ONLY automation...');
  
  try {
    // Step 1: Verify we're using the right sheets
    verifyFreshDataSheets();
    
    // Step 2: Transform ONLY fresh data
    const recordCount = transformFreshDataOnly();
    
    // Step 3: Create clean dashboard
    createCleanDashboard();
    
    console.log(`✅ Fresh data automation complete!`);
    console.log(`📊 Processed ${recordCount} records from Week 33 data ONLY`);
    
    SpreadsheetApp.getUi().alert(
      'Fresh Data Automation Complete!', 
      `✅ Success! Processed ${recordCount} records\n📅 Data: Week 34-53 (202534-202553)\n📊 Quarters: Q3 2025, Q4 2025, Q1 2026\n\n⚠️ Q1 2026 data IS correct (Week 202553 from your fresh data)`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ Fresh data automation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Automation Error', 
      `Error: ${error.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

function verifyFreshDataSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('🔍 Verifying fresh data sheets...');
  
  // Check for EXACT sheet names
  const constrainedSheet = ss.getSheetByName('NEW WIDE CONSTRAINED FCST DATA');
  const unconstrainedSheet = ss.getSheetByName('NEW WIDE UNCONST. FCST DATA');
  
  if (!constrainedSheet) {
    throw new Error('❌ Could not find "NEW WIDE CONSTRAINED FCST DATA" sheet. Make sure the sheet name is EXACTLY this.');
  }
  
  if (!unconstrainedSheet) {
    throw new Error('❌ Could not find "NEW WIDE UNCONST. FCST DATA" sheet. Make sure the sheet name is EXACTLY this.');
  }
  
  // Verify the data looks right
  const constrainedData = constrainedSheet.getDataRange().getValues();
  const headerRow = constrainedData[0];
  
  // Check for Week 34 (should be first week in fresh data)
  const hasWeek34 = headerRow.some(cell => cell === 202534);
  if (!hasWeek34) {
    throw new Error('❌ Week 202534 not found in constrained data. This might not be fresh Week 33 data.');
  }
  
  // Check that we don't have Week 29 (old data indicator)
  const hasWeek29 = headerRow.some(cell => cell === 202529);
  if (hasWeek29) {
    throw new Error('❌ Week 202529 found in data. This appears to be old Week 29 data, not fresh Week 33 data.');
  }
  
  console.log('✅ Fresh data sheets verified');
  console.log(`   Constrained: ${constrainedData.length} rows`);
  console.log(`   Unconstrained: ${unconstrainedSheet.getLastRow()} rows`);
}

function transformFreshDataOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get ONLY the fresh data sheets
  const constrainedSheet = ss.getSheetByName('NEW WIDE CONSTRAINED FCST DATA');
  const unconstrainedSheet = ss.getSheetByName('NEW WIDE UNCONST. FCST DATA');
  
  // Create or clear output sheet
  let outputSheet = ss.getSheetByName('Fresh_Data_Week33');
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet('Fresh_Data_Week33');
  }
  
  // Headers
  const headers = ['Customer', 'Anker SKU', 'PDT', 'Forecast Type', 'Quarter', 'Week', 
                   'Forecast - Units', 'Forecast Revenue', 'Delta Units', 'Delta - Revenue', 
                   'Gap Flag', 'IsCurrentQ', 'Important Helper', 'Sell-In Price', 'Data Source'];
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get data ranges
  const constrainedData = constrainedSheet.getDataRange().getValues();
  const unconstrainedData = unconstrainedSheet.getDataRange().getValues();
  
  console.log(`Processing constrained: ${constrainedData.length} rows`);
  console.log(`Processing unconstrained: ${unconstrainedData.length} rows`);
  
  // Find week columns - ONLY from fresh data
  const weekColumns = [];
  const headerRow = constrainedData[0];
  
  for (let i = 0; i < headerRow.length; i++) {
    const cellValue = headerRow[i];
    // STRICT validation: only weeks 202534-202553 (fresh Week 33 data)
    if (cellValue && typeof cellValue === 'number' && 
        cellValue >= 202534 && cellValue <= 202553) {
      weekColumns.push({index: i, week: cellValue});
    }
  }
  
  console.log(`Found ${weekColumns.length} fresh week columns: ${weekColumns[0]?.week} to ${weekColumns[weekColumns.length-1]?.week}`);
  
  if (weekColumns.length === 0) {
    throw new Error('❌ No fresh week columns found (202534-202553). Check your data.');
  }
  
  // Create unconstrained lookup - ONLY from fresh data
  const unconstrainedLookup = {};
  
  for (let i = 1; i < unconstrainedData.length; i++) {
    const row = unconstrainedData[i];
    const helper = row[0]; // Important Helper
    if (!helper) continue;
    
    const sellInPrice = parseFloat(row[2]) || 0; // Sell-in Price
    
    weekColumns.forEach(weekCol => {
      const week = weekCol.week;
      const forecastRevenue = parseFloat(row[weekCol.index]) || 0; // DOLLARS
      const forecastUnits = sellInPrice > 0 ? forecastRevenue / sellInPrice : 0;
      
      const key = `${helper}_${week}`;
      unconstrainedLookup[key] = {
        units: forecastUnits,
        revenue: forecastRevenue
      };
    });
  }
  
  console.log(`Created lookup with ${Object.keys(unconstrainedLookup).length} entries`);
  
  // Process constrained data in batches
  const BATCH_SIZE = 25; // Smaller batches for reliability
  const allOutputData = [];
  let totalRecords = 0;
  
  for (let startRow = 1; startRow < constrainedData.length; startRow += BATCH_SIZE) {
    const endRow = Math.min(startRow + BATCH_SIZE, constrainedData.length);
    console.log(`Processing batch: rows ${startRow} to ${endRow}`);
    
    const batchData = [];
    
    for (let i = startRow; i < endRow; i++) {
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
        
        const quarter = getQuarterFresh(week);
        const isCurrentQ = quarter === 'Q4 2025';
        const gapFlag = deltaUnits < 0 ? 'Supply Gap' : '';
        
        // Only include records with actual data
        if (constrainedRevenue > 0 || unconstrained.revenue > 0) {
          // Add constrained row
          batchData.push([
            customer, sku, pdt, 'Constrained', quarter, week,
            constrainedUnits, constrainedRevenue, deltaUnits, deltaRevenue,
            gapFlag, isCurrentQ, helper, sellInPrice, 'Fresh Week 33'
          ]);
          
          // Add unconstrained row
          batchData.push([
            customer, sku, pdt, 'Unconstrained', quarter, week,
            unconstrained.units, unconstrained.revenue, 0, 0,
            '', isCurrentQ, helper, sellInPrice, 'Fresh Week 33'
          ]);
        }
      });
    }
    
    // Write batch
    if (batchData.length > 0) {
      const startOutputRow = allOutputData.length + 2;
      outputSheet.getRange(startOutputRow, 1, batchData.length, headers.length).setValues(batchData);
      allOutputData.push(...batchData);
      totalRecords += batchData.length;
      
      console.log(`Batch complete: ${batchData.length} records (total: ${totalRecords})`);
    }
    
    Utilities.sleep(50); // Brief pause
  }
  
  // Format sheet
  formatFreshDataSheet(outputSheet, headers.length);
  
  console.log(`✅ Fresh data transformation complete! ${totalRecords} records created.`);
  return totalRecords;
}

function getQuarterFresh(week) {
  // EXACT mapping for fresh Week 33 data
  if (week >= 202534 && week <= 202539) return 'Q3 2025';
  if (week >= 202540 && week <= 202552) return 'Q4 2025';
  if (week === 202553) return 'Q1 2026'; // Only Week 53 is Q1 2026
  return 'Unknown';
}

function createCleanDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('📊 Creating clean dashboard...');
  
  let dashboardSheet = ss.getSheetByName('Dashboard_Fresh_Week33');
  if (dashboardSheet) {
    dashboardSheet.clear();
  } else {
    dashboardSheet = ss.insertSheet('Dashboard_Fresh_Week33');
  }
  
  // Title
  dashboardSheet.getRange('A1').setValue('ANKER SUPPLY GAP ANALYSIS - FRESH WEEK 33 DATA ONLY');
  dashboardSheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  dashboardSheet.getRange('A1:H1').merge();
  
  // Data verification
  const timestamp = new Date().toLocaleString();
  dashboardSheet.getRange('A2').setValue(`Generated: ${timestamp} | Source: NEW WIDE CONSTRAINED/UNCONSTRAINED FCST DATA ONLY`);
  dashboardSheet.getRange('A2').setFontStyle('italic');
  dashboardSheet.getRange('A2:H2').merge();
  
  // Week range verification
  dashboardSheet.getRange('A3').setValue('Data Range: Week 34-53 (202534-202553) | Quarters: Q3 2025, Q4 2025, Q1 2026');
  dashboardSheet.getRange('A3').setFontWeight('bold').setBackground('#E8F0FE');
  dashboardSheet.getRange('A3:H3').merge();
  
  // Key metrics
  createFreshMetrics(dashboardSheet);
  
  console.log('✅ Clean dashboard created');
}

function createFreshMetrics(sheet) {
  const metricsStart = 5;
  
  sheet.getRange(`A${metricsStart}`).setValue('KEY METRICS - FRESH WEEK 33 DATA ONLY');
  sheet.getRange(`A${metricsStart}`).setFontSize(14).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  sheet.getRange(`A${metricsStart}:C${metricsStart}`).merge();
  
  const metrics = [
    ['Metric', 'Formula/Value', 'Description'],
    ['Total Supply Gaps', '=COUNTIFS(Fresh_Data_Week33!K:K,"Supply Gap")', 'Number of supply gap scenarios'],
    ['Revenue at Risk', '=SUMIFS(Fresh_Data_Week33!J:J,Fresh_Data_Week33!K:K,"Supply Gap")*-1', 'Total revenue impact'],
    ['Units at Risk', '=SUMIFS(Fresh_Data_Week33!I:I,Fresh_Data_Week33!K:K,"Supply Gap")*-1', 'Total units affected'],
    ['Q3 2025 Impact', '=SUMIFS(Fresh_Data_Week33!J:J,Fresh_Data_Week33!K:K,"Supply Gap",Fresh_Data_Week33!E:E,"Q3 2025")*-1', 'Q3 revenue at risk'],
    ['Q4 2025 Impact', '=SUMIFS(Fresh_Data_Week33!J:J,Fresh_Data_Week33!K:K,"Supply Gap",Fresh_Data_Week33!E:E,"Q4 2025")*-1', 'Q4 revenue at risk'],
    ['Q1 2026 Impact', '=SUMIFS(Fresh_Data_Week33!J:J,Fresh_Data_Week33!K:K,"Supply Gap",Fresh_Data_Week33!E:E,"Q1 2026")*-1', 'Q1 2026 revenue at risk (Week 53 only)'],
    ['', '', ''],
    ['Data Verification:', '', ''],
    ['Total Records', '=COUNTA(Fresh_Data_Week33!A:A)-1', 'Should be 15,000-25,000 records'],
    ['Week Range Check', '=MIN(Fresh_Data_Week33!F:F)&" to "&MAX(Fresh_Data_Week33!F:F)', 'Should be 202534 to 202553'],
    ['Data Source Check', '=COUNTIF(Fresh_Data_Week33!O:O,"Fresh Week 33")', 'All records should be from fresh data']
  ];
  
  sheet.getRange(metricsStart + 1, 1, metrics.length, 3).setValues(metrics);
  
  // Format headers
  sheet.getRange(metricsStart + 1, 1, 1, 3).setBackground('#E8F0FE').setFontWeight('bold');
  
  // Format currency
  sheet.getRange(metricsStart + 3, 2, 4, 1).setNumberFormat('$#,##0');
}

function formatFreshDataSheet(sheet, numColumns) {
  sheet.setFrozenRows(1);
  
  const headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // Auto-resize key columns
  for (let i = 1; i <= Math.min(8, numColumns); i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Menu setup
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 Fresh Data ONLY Automation')
    .addItem('🚀 Run Fresh Data Automation', 'runFreshDataAutomation')
    .addItem('🔍 Verify Fresh Data Sheets', 'verifyFreshDataSheets')
    .addToUi();
}
