/**
 * DIAGNOSTIC SCRIPT - Debug the Google Sheets automation
 * This will help us understand what went wrong
 */

function diagnoseProblem() {
  console.log('🔍 Starting diagnostic...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Get all sheet names
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getName());
  
  console.log('📋 Available sheets:', sheetNames);
  
  let diagnosticReport = '🔍 DIAGNOSTIC REPORT\n\n';
  diagnosticReport += '📋 AVAILABLE SHEETS:\n';
  sheetNames.forEach((name, index) => {
    diagnosticReport += `${index + 1}. ${name}\n`;
  });
  
  // Check for expected source sheets
  const expectedSheets = [
    'NEW WIDE CONSTRAINED FCST DATA',
    'NEW WIDE UNCONST. FCST DATA'
  ];
  
  diagnosticReport += '\n🎯 LOOKING FOR EXPECTED SHEETS:\n';
  expectedSheets.forEach(expected => {
    const found = sheetNames.includes(expected);
    diagnosticReport += `${found ? '✅' : '❌'} ${expected}\n`;
  });
  
  // Look for sheets with similar names
  diagnosticReport += '\n🔍 SHEETS WITH "FCST" OR "DATA":\n';
  const relevantSheets = sheetNames.filter(name => 
    name.toUpperCase().includes('FCST') || 
    name.toUpperCase().includes('DATA') ||
    name.toUpperCase().includes('ANKER') ||
    name.toUpperCase().includes('WEEK')
  );
  
  relevantSheets.forEach(name => {
    diagnosticReport += `📊 ${name}\n`;
  });
  
  // Check if we have any sheets to analyze
  if (relevantSheets.length > 0) {
    const firstRelevantSheet = ss.getSheetByName(relevantSheets[0]);
    if (firstRelevantSheet) {
      diagnosticReport += `\n📊 ANALYZING FIRST RELEVANT SHEET: ${relevantSheets[0]}\n`;
      
      try {
        const dataRange = firstRelevantSheet.getDataRange();
        const data = dataRange.getValues();
        
        diagnosticReport += `Rows: ${data.length}\n`;
        diagnosticReport += `Columns: ${data[0] ? data[0].length : 0}\n`;
        
        if (data.length > 0) {
          diagnosticReport += '\n📊 FIRST ROW (HEADERS):\n';
          const headers = data[0];
          headers.slice(0, 10).forEach((header, index) => {
            diagnosticReport += `Col ${index + 1}: ${header}\n`;
          });
          
          // Look for week columns
          diagnosticReport += '\n📅 WEEK COLUMNS FOUND:\n';
          const weekColumns = [];
          for (let i = 0; i < headers.length; i++) {
            const cellValue = headers[i];
            if (cellValue && typeof cellValue === 'number' && cellValue.toString().match(/^202\d\d\d$/)) {
              weekColumns.push({index: i, week: cellValue});
              diagnosticReport += `Col ${i + 1}: ${cellValue}\n`;
            }
          }
          
          if (weekColumns.length === 0) {
            diagnosticReport += '❌ No week columns found with format 202XXX\n';
            diagnosticReport += '\n📊 CHECKING OTHER NUMERIC COLUMNS:\n';
            for (let i = 0; i < Math.min(headers.length, 20); i++) {
              const cellValue = headers[i];
              if (cellValue && typeof cellValue === 'number') {
                diagnosticReport += `Col ${i + 1}: ${cellValue} (${typeof cellValue})\n`;
              }
            }
          }
          
          // Check for key columns
          diagnosticReport += '\n🔍 LOOKING FOR KEY COLUMNS:\n';
          const keyTerms = ['customer', 'sku', 'pdt', 'helper', 'price'];
          keyTerms.forEach(term => {
            const found = headers.some(header => 
              header && header.toString().toLowerCase().includes(term)
            );
            diagnosticReport += `${found ? '✅' : '❌'} Column containing "${term}"\n`;
          });
        }
        
      } catch (error) {
        diagnosticReport += `❌ Error analyzing sheet: ${error.message}\n`;
      }
    }
  }
  
  // Show the diagnostic report
  console.log(diagnosticReport);
  
  // Split into chunks if too long for alert
  const maxLength = 1000;
  if (diagnosticReport.length > maxLength) {
    const chunks = [];
    for (let i = 0; i < diagnosticReport.length; i += maxLength) {
      chunks.push(diagnosticReport.slice(i, i + maxLength));
    }
    
    chunks.forEach((chunk, index) => {
      ui.alert(
        `Diagnostic Report (Part ${index + 1}/${chunks.length})`,
        chunk,
        ui.ButtonSet.OK
      );
    });
  } else {
    ui.alert('Diagnostic Report', diagnosticReport, ui.ButtonSet.OK);
  }
  
  return diagnosticReport;
}

function fixSheetNamesAndRerun() {
  console.log('🔧 Attempting to fix sheet names and rerun...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getName());
  
  // Try to find and rename sheets that look like our source data
  const constrainedCandidates = sheetNames.filter(name => 
    name.toUpperCase().includes('CONSTRAINED') || 
    name.toUpperCase().includes('CONST')
  );
  
  const unconstrainedCandidates = sheetNames.filter(name => 
    name.toUpperCase().includes('UNCONSTRAINED') || 
    name.toUpperCase().includes('UNCONST')
  );
  
  let fixReport = '🔧 FIXING SHEET NAMES:\n\n';
  
  if (constrainedCandidates.length > 0) {
    const sheet = ss.getSheetByName(constrainedCandidates[0]);
    sheet.setName('NEW WIDE CONSTRAINED FCST DATA');
    fixReport += `✅ Renamed "${constrainedCandidates[0]}" to "NEW WIDE CONSTRAINED FCST DATA"\n`;
  }
  
  if (unconstrainedCandidates.length > 0) {
    const sheet = ss.getSheetByName(unconstrainedCandidates[0]);
    sheet.setName('NEW WIDE UNCONST. FCST DATA');
    fixReport += `✅ Renamed "${unconstrainedCandidates[0]}" to "NEW WIDE UNCONST. FCST DATA"\n`;
  }
  
  if (constrainedCandidates.length === 0 || unconstrainedCandidates.length === 0) {
    fixReport += '❌ Could not find both required sheets to rename\n';
    fixReport += 'Please manually rename your sheets to:\n';
    fixReport += '- NEW WIDE CONSTRAINED FCST DATA\n';
    fixReport += '- NEW WIDE UNCONST. FCST DATA\n';
  }
  
  SpreadsheetApp.getUi().alert('Sheet Name Fix', fixReport, SpreadsheetApp.getUi().ButtonSet.OK);
  
  // If we successfully renamed both, try running the automation again
  if (constrainedCandidates.length > 0 && unconstrainedCandidates.length > 0) {
    try {
      runCompleteManagerAutomation();
    } catch (error) {
      SpreadsheetApp.getUi().alert('Error After Fix', `Fixed sheet names but automation still failed: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

// Add the original optimized functions here
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
      `Error: ${error.message}\n\nTry running "Diagnose Problem" first to see what's wrong`, 
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
  
  // Find week columns - ENHANCED DETECTION
  const weekColumns = [];
  const headerRow = constrainedData[0];
  
  for (let i = 0; i < headerRow.length; i++) {
    const cellValue = headerRow[i];
    
    // Try multiple formats for week detection
    if (cellValue && typeof cellValue === 'number') {
      const valueStr = cellValue.toString();
      
      // Original format: 202534 (6 digits)
      if (valueStr.match(/^202\d\d\d$/)) {
        weekColumns.push({index: i, week: cellValue});
      }
      // Alternative: 20253 (5 digits)
      else if (valueStr.match(/^2025\d$/)) {
        weekColumns.push({index: i, week: cellValue});
      }
    }
    // Also check for string formats
    else if (cellValue && typeof cellValue === 'string') {
      const cleanValue = cellValue.toString().replace(/[^\d]/g, '');
      if (cleanValue.match(/^202\d\d\d$/)) {
        weekColumns.push({index: i, week: parseInt(cleanValue)});
      }
    }
  }
  
  console.log(`Found ${weekColumns.length} week columns`);
  
  if (weekColumns.length === 0) {
    throw new Error(`No week columns found! Header row: ${headerRow.slice(0, 10).join(', ')}...`);
  }
  
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
  dashboardSheet.getRange('A1').setValue('ANKER SUPPLY GAP ANALYSIS - FRESH WEEK 33 DATA ONLY');
  dashboardSheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  dashboardSheet.getRange('A1:H1').merge();
  
  // Timestamp
  const timestamp = new Date().toLocaleString();
  dashboardSheet.getRange('A2').setValue(`Generated: ${timestamp} | Source: NEW WIDE CONSTRAINED/UNCONSTRAINED FCST DATA ONLY`);
  dashboardSheet.getRange('A2').setFontStyle('italic');
  dashboardSheet.getRange('A2:H2').merge();
  
  dashboardSheet.getRange('A3').setValue('Data Range: Week 34-53 (202534-202553) | Quarters: Q3 2025, Q4 2025, Q1 2026');
  dashboardSheet.getRange('A3:H3').merge();
  
  // Key metrics
  createKeyMetricsOptimized(dashboardSheet);
  
  console.log('✅ Optimized dashboard created');
}

function createKeyMetricsOptimized(sheet) {
  const metricsStart = 5;
  sheet.getRange(`A${metricsStart}`).setValue('KEY METRICS - FRESH WEEK 33');
  sheet.getRange(`A${metricsStart}`).setFontSize(14).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  sheet.getRange(`A${metricsStart}:C${metricsStart}`).merge();
  
  const metrics = [
    ['Metric', 'Formula/Value', 'Description'],
    ['Total Supply Gaps', '=COUNTIFS(Looker_Ready_View_Week33!K:K,"Supply Gap")', 'Number of supply gap scenarios'],
    ['Revenue at Risk', '=SUMIFS(Looker_Ready_View_Week33!J:J,Looker_Ready_View_Week33!K:K,"Supply Gap")*-1', 'Total revenue impact'],
    ['Units at Risk', '=SUMIFS(Looker_Ready_View_Week33!I:I,Looker_Ready_View_Week33!K:K,"Supply Gap")*-1', 'Total units affected'],
    ['Q3 2025 Impact', '=SUMIFS(Looker_Ready_View_Week33!J:J,Looker_Ready_View_Week33!K:K,"Supply Gap",Looker_Ready_View_Week33!E:E,"Q3 2025")*-1', 'Q3 revenue at risk'],
    ['Q4 2025 Impact', '=SUMIFS(Looker_Ready_View_Week33!J:J,Looker_Ready_View_Week33!K:K,"Supply Gap",Looker_Ready_View_Week33!E:E,"Q4 2025")*-1', 'Q4 revenue at risk'],
    ['Q1 2026 Impact', '=SUMIFS(Looker_Ready_View_Week33!J:J,Looker_Ready_View_Week33!K:K,"Supply Gap",Looker_Ready_View_Week33!E:E,"Q1 2026")*-1', 'Q1 revenue at risk (Week 53 only)'],
    ['', '', ''],
    ['Data Verification:', '', ''],
    ['Total Records', '=COUNTA(Looker_Ready_View_Week33!A:A)-1', 'Should be 15,000-25,000 records'],
    ['Week Range Check', 'Week 34-53 (202534-202553)', 'All records should be from fresh data'],
    ['Data Source Check', '14242 Should be 15,000-25,000 records', 'All records should be from fresh data']
  ];
  
  const metricsRange = sheet.getRange(metricsStart + 1, 1, metrics.length, 3);
  metricsRange.setValues(metrics);
  
  // Format headers
  sheet.getRange(metricsStart + 1, 1, 1, 3).setBackground('#E8F0FE').setFontWeight('bold');
  
  // Format currency values (rows 3-7)
  sheet.getRange(metricsStart + 3, 2, 5, 1).setNumberFormat('$#,##0');
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
 * Menu setup
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔍 Diagnostic & Fixed Automation')
    .addItem('🔍 Diagnose Problem', 'diagnoseProblem')
    .addItem('🔧 Fix Sheet Names & Rerun', 'fixSheetNamesAndRerun')
    .addSeparator()
    .addItem('🚀 Run Complete Automation', 'runCompleteManagerAutomation')
    .addToUi();
}
