/**
 * ANKER SUPPLY GAP ANALYSIS DASHBOARD
 * Professional automation script for forecast analysis and dashboard creation
 * 
 * Features:
 * - Automated data transformation from constrained/unconstrained forecasts
 * - Interactive dashboard with multiple analysis views
 * - Professional charts and visualizations
 * - Real-time data analysis with QUERY formulas
 * 
 * Created: 2025
 * Version: 2.0 - Refactored Professional Edition
 */

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================

/**
 * Main automation function - runs complete forecast analysis
 */
function runForecastAnalysisDashboard() {
  console.log('🚀 Starting Anker Forecast Analysis Dashboard automation...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if we have existing processed data
    const existingData = ss.getSheetByName('Looker_Ready_View');
    
    if (existingData && existingData.getLastRow() > 1) {
      console.log('📊 Found existing processed data - creating dashboard from current data');
      
      // Create dashboard using existing data
      createExecutiveDashboard();
      createAnalysisDashboards();
      
      console.log(`✅ Dashboard automation complete!`);
      console.log(`📊 Used existing data (${existingData.getLastRow() - 1} records)`);
      
      showCompletionMessage('existing');
      
    } else {
      // Process new data and create dashboard
      console.log('📊 Processing new forecast data...');
      const recordCount = transformForecastData();
      
      createExecutiveDashboard();
      createAnalysisDashboards();
      
      console.log(`✅ Complete automation finished!`);
      console.log(`📊 Processed ${recordCount} records`);
      
      showCompletionMessage('new', recordCount);
    }
    
  } catch (error) {
    console.error('❌ Automation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Automation Error', 
      `Error: ${error.message}\n\nPlease check the execution log for details.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Shows completion message to user
 */
function showCompletionMessage(type, recordCount = 0) {
  const ui = SpreadsheetApp.getUi();
  
  if (type === 'existing') {
    ui.alert(
      'Anker Dashboard Complete!', 
      `✅ Successfully used existing Looker_Ready_View data\n📊 Created Executive Dashboard\n📈 Created Customer Analysis\n📈 Created SKU Analysis\n📈 Created Weekly Trends\n📈 Created Product Analysis\n📊 Created Supply Gap Summary Chart\n\n🎯 All dashboards updated with current data!`, 
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      'Anker Dashboard Complete!', 
      `✅ Successfully processed ${recordCount} records\n📊 Created Executive Dashboard\n📈 Created Customer Analysis\n📈 Created SKU Analysis\n📈 Created Weekly Trends\n📈 Created Product Analysis\n📊 Created Supply Gap Summary Chart\n\n🎯 All dashboards created successfully!`, 
      ui.ButtonSet.OK
    );
  }
}

// ============================================================================
// DATA TRANSFORMATION
// ============================================================================

/**
 * Transforms raw forecast data into analysis-ready format
 */
function transformForecastData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Log all available sheet names for debugging
  const allSheets = ss.getSheets();
  console.log('Available sheets:');
  allSheets.forEach(sheet => console.log(`  - "${sheet.getName()}"`));
  
  // Find source sheets with flexible naming
  const constrainedSheet = findConstrainedSheet(allSheets);
  const unconstrainedSheet = findUnconstrainedSheet(allSheets);
  
  if (!constrainedSheet || !unconstrainedSheet) {
    const availableNames = allSheets.map(s => s.getName()).join('", "');
    throw new Error(`Could not find source sheets. Available sheets: "${availableNames}". Please ensure your constrained and unconstrained sheets are properly named.`);
  }
  
  console.log(`Found sheets: "${constrainedSheet.getName()}" and "${unconstrainedSheet.getName()}"`);
  
  // Create or clear output sheet
  let outputSheet = ss.getSheetByName('Looker_Ready_View');
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet('Looker_Ready_View');
  }
  
  // Set up headers
  const outputHeaders = ['Customer', 'ID', 'Anker SKU', 'PDT', 'Forecast Type', 'Quarter', 'Week', 
                        'Forecast - Units', 'Forecast Revenue', 'Delta Units', 'Delta Revenue', 
                        'Gap Flag', 'IsCurrentQ', 'Important Helper', 'Sell-In Price'];
  outputSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  
  // Process the data
  const recordCount = processData(constrainedSheet, unconstrainedSheet, outputSheet, outputHeaders);
  
  // Format the output sheet
  formatOutputSheet(outputSheet, outputHeaders.length);
  
  console.log(`✅ Data transformation complete! ${recordCount} records created.`);
  return recordCount;
}

/**
 * Finds constrained forecast sheet with flexible naming
 */
function findConstrainedSheet(allSheets) {
  return ss.getSheetByName('NEW WIDE CONSTRAINED FCST DATA') ||
         ss.getSheetByName('Constrained Wide') ||
         ss.getSheetByName('Constrained Forecast - "Confirm') ||
         ss.getSheetByName('CONSTRAINED') ||
         allSheets.find(sheet => sheet.getName().toUpperCase().includes('CONSTRAINED'));
}

/**
 * Finds unconstrained forecast sheet with flexible naming
 */
function findUnconstrainedSheet(allSheets) {
  return ss.getSheetByName('NEW WIDE UNCONST. FCST DATA') ||
         ss.getSheetByName('Unconstrained Wide') ||
         ss.getSheetByName('Unconstrained Forecast - "Optim') ||
         ss.getSheetByName('UNCONSTRAINED') ||
         allSheets.find(sheet => sheet.getName().toUpperCase().includes('UNCONST'));
}

/**
 * Processes constrained and unconstrained data
 */
function processData(constrainedSheet, unconstrainedSheet, outputSheet, outputHeaders) {
  const constrainedData = constrainedSheet.getDataRange().getValues();
  const unconstrainedData = unconstrainedSheet.getDataRange().getValues();
  
  console.log(`Processing ${constrainedData.length} constrained rows, ${unconstrainedData.length} unconstrained rows`);
  
  // Column mapping based on standard layout
  const COLUMN_MAP = {
    HELPER: 0,        // Important Helper
    SHIFT_LOGIC: 1,   // Shift logic  
    SELL_IN_PRICE: 2, // Sell-in price
    CATEGORY: 3,      // Category
    PCT: 4,           // PCT
    BG: 5,            // BG
    PDT: 6,           // PDT
    CUSTOMER_ID: 7,   // Customer ID
    CUSTOMER: 8,      // Customer
    ANKER_SKU: 9,     // Anker SKU
    WEEK_START: 10    // First week column
  };
  
  // Find week columns
  const weekColumns = findWeekColumns(constrainedData[0], COLUMN_MAP.WEEK_START);
  console.log(`Found ${weekColumns.length} week columns: ${weekColumns.map(w => w.week).join(', ')}`);
  
  if (weekColumns.length === 0) {
    throw new Error('No week columns found in data!');
  }
  
  // Create unconstrained lookup
  const unconstrainedLookup = createUnconstrainedLookup(unconstrainedData, COLUMN_MAP, weekColumns);
  console.log(`Created lookup with ${Object.keys(unconstrainedLookup).length} entries`);
  
  // Process constrained data in batches
  const totalRecords = processConstrainedData(constrainedData, unconstrainedLookup, outputSheet, outputHeaders, COLUMN_MAP, weekColumns);
  
  return totalRecords;
}

/**
 * Finds week columns in the header row
 */
function findWeekColumns(headerRow, startColumn) {
  const weekColumns = [];
  
  for (let i = startColumn; i < headerRow.length; i++) {
    const cellValue = headerRow[i];
    if (cellValue && typeof cellValue === 'number' && cellValue.toString().match(/^202\d\d\d$/)) {
      weekColumns.push({index: i, week: cellValue});
    }
  }
  
  return weekColumns;
}

/**
 * Creates lookup table from unconstrained data
 */
function createUnconstrainedLookup(unconstrainedData, COLUMN_MAP, weekColumns) {
  const lookup = {};
  
  for (let i = 1; i < unconstrainedData.length; i++) {
    if (i % 100 === 0) {
      console.log(`Processing unconstrained row ${i}/${unconstrainedData.length}`);
      Utilities.sleep(10);
    }
    
    const row = unconstrainedData[i];
    const helper = row[COLUMN_MAP.HELPER];
    if (!helper) continue;
    
    const sellInPrice = parseCurrencyValue(row[COLUMN_MAP.SELL_IN_PRICE]);
    
    weekColumns.forEach(weekCol => {
      const week = weekCol.week;
      const forecastRevenue = parseCurrencyValue(row[weekCol.index]);
      const forecastUnits = sellInPrice > 0 ? forecastRevenue / sellInPrice : 0;
      
      const key = `${helper}_${week}`;
      lookup[key] = {
        units: forecastUnits,
        revenue: forecastRevenue,
        sellInPrice: sellInPrice
      };
    });
  }
  
  return lookup;
}

/**
 * Processes constrained data and writes to output sheet
 */
function processConstrainedData(constrainedData, unconstrainedLookup, outputSheet, outputHeaders, COLUMN_MAP, weekColumns) {
  const BATCH_SIZE = 100;
  let totalRecords = 0;
  
  for (let startRow = 1; startRow < constrainedData.length; startRow += BATCH_SIZE) {
    const endRow = Math.min(startRow + BATCH_SIZE, constrainedData.length);
    console.log(`Processing constrained batch: rows ${startRow} to ${endRow}`);
    
    const batchData = [];
    
    for (let i = startRow; i < endRow; i++) {
      const row = constrainedData[i];
      const helper = row[COLUMN_MAP.HELPER];
      const sellInPrice = parseCurrencyValue(row[COLUMN_MAP.SELL_IN_PRICE]);
      const pdt = row[COLUMN_MAP.PDT];
      const customer = row[COLUMN_MAP.CUSTOMER];
      const customerId = row[COLUMN_MAP.CUSTOMER_ID];
      const sku = row[COLUMN_MAP.ANKER_SKU];
      
      if (!helper || !customer || !sku) continue;
      
      weekColumns.forEach(weekCol => {
        const week = weekCol.week;
        const constrainedRevenue = parseCurrencyValue(row[weekCol.index]);
        const constrainedUnits = sellInPrice > 0 ? constrainedRevenue / sellInPrice : 0;
        
        const key = `${helper}_${week}`;
        const unconstrained = unconstrainedLookup[key] || { units: 0, revenue: 0, sellInPrice: 0 };
        
        const deltaUnits = constrainedUnits - unconstrained.units;
        const deltaRevenue = constrainedRevenue - unconstrained.revenue;
        
        const quarter = getQuarter(week);
        const isCurrentQ = quarter === 'Q4 2025';
        const gapFlag = deltaUnits < 0 ? 'Supply Gap' : '';
        
        // Add records with any forecast data or meaningful delta
        if (constrainedRevenue > 0 || unconstrained.revenue > 0 || Math.abs(deltaRevenue) > 0.01) {
          // Add constrained row
          batchData.push([
            customer, customerId, sku, pdt, 'Constrained', quarter, week,
            constrainedUnits, constrainedRevenue, deltaUnits, deltaRevenue,
            gapFlag, isCurrentQ, helper, sellInPrice
          ]);
          
          // Add unconstrained row
          batchData.push([
            customer, customerId, sku, pdt, 'Unconstrained', quarter, week,
            unconstrained.units, unconstrained.revenue, 0, 0,
            '', isCurrentQ, helper, unconstrained.sellInPrice || sellInPrice
          ]);
        }
      });
    }
    
    // Write batch to sheet
    if (batchData.length > 0) {
      const startOutputRow = totalRecords + 2;
      outputSheet.getRange(startOutputRow, 1, batchData.length, outputHeaders.length).setValues(batchData);
      totalRecords += batchData.length;
      
      console.log(`Wrote batch: ${batchData.length} records (total: ${totalRecords})`);
    }
    
    Utilities.sleep(100);
  }
  
  return totalRecords;
}

/**
 * Enhanced currency parsing function
 */
function parseCurrencyValue(value) {
  if (value === null || value === undefined || value === '' || value === '#N/A' || value === 'N/A') {
    return 0;
  }
  
  // Handle string values with dollar signs and commas
  if (typeof value === 'string') {
    const cleanValue = value.replace(/[\$",\s]/g, '');
    
    if (cleanValue === '#N/A' || cleanValue === 'N/A' || cleanValue === '' || cleanValue === '#DIV/0!') {
      return 0;
    }
    
    const parsed = parseFloat(cleanValue);
    return isNaN(parsed) ? 0 : parsed;
  }
  
  // Handle numeric values
  if (typeof value === 'number') {
    return isNaN(value) ? 0 : value;
  }
  
  return 0;
}

/**
 * Determines quarter from week number
 */
function getQuarter(week) {
  if (week >= 202534 && week <= 202539) return 'Q3 2025';
  if (week >= 202540 && week <= 202552) return 'Q4 2025';
  if (week >= 202553 && week <= 202605) return 'Q1 2026';
  return 'Unknown';
}

/**
 * Formats the output sheet
 */
function formatOutputSheet(sheet, numColumns) {
  console.log('Formatting output sheet...');
  
  sheet.setFrozenRows(1);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // Auto-resize columns
  for (let i = 1; i <= Math.min(10, numColumns); i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Format currency columns
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 9, sheet.getLastRow() - 1, 1).setNumberFormat('$#,##0.00'); // Forecast Revenue
    sheet.getRange(2, 11, sheet.getLastRow() - 1, 1).setNumberFormat('$#,##0.00'); // Delta Revenue
    sheet.getRange(2, 15, sheet.getLastRow() - 1, 1).setNumberFormat('$#,##0.00'); // Sell-In Price
  }
  
  console.log('Output sheet formatting complete');
}

// ============================================================================
// DASHBOARD CREATION
// ============================================================================

/**
 * Creates the executive dashboard
 */
function createExecutiveDashboard() {
  console.log('📊 Creating executive dashboard...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create dashboard
  let dashboardSheet = ss.getSheetByName('Executive Dashboard');
  if (dashboardSheet) {
    dashboardSheet.clear();
  } else {
    dashboardSheet = ss.insertSheet('Executive Dashboard');
  }
  
  // Title
  dashboardSheet.getRange('A1').setValue('ANKER SUPPLY GAP ANALYSIS - EXECUTIVE DASHBOARD');
  dashboardSheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  dashboardSheet.getRange('A1:H1').merge();
  
  // Timestamp
  const timestamp = new Date().toLocaleString();
  dashboardSheet.getRange('A2').setValue(`Generated: ${timestamp} | Source: Looker_Ready_View data`);
  dashboardSheet.getRange('A2').setFontStyle('italic');
  dashboardSheet.getRange('A2:H2').merge();
  
  // Key metrics
  createKeyMetrics(dashboardSheet);
  
  // Navigation links
  createNavigationLinks(dashboardSheet);
  
  console.log('✅ Executive dashboard created');
}

/**
 * Creates key metrics section
 */
function createKeyMetrics(sheet) {
  const metricsStart = 5;
  sheet.getRange(`A${metricsStart}`).setValue('KEY PERFORMANCE INDICATORS');
  sheet.getRange(`A${metricsStart}`).setFontSize(14).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  sheet.getRange(`A${metricsStart}:C${metricsStart}`).merge();
  
  const metrics = [
    ['Metric', 'Value', 'Description'],
    ['Total Supply Gaps', '=COUNTIFS(Looker_Ready_View!L:L,"Supply Gap")', 'Number of supply gap scenarios'],
    ['Revenue at Risk', '=SUMIFS(Looker_Ready_View!K:K,Looker_Ready_View!L:L,"Supply Gap")*-1', 'Total revenue impact from gaps'],
    ['Units at Risk', '=SUMIFS(Looker_Ready_View!J:J,Looker_Ready_View!L:L,"Supply Gap")*-1', 'Total units affected'],
    ['Q3 2025 Impact', '=SUMIFS(Looker_Ready_View!K:K,Looker_Ready_View!L:L,"Supply Gap",Looker_Ready_View!F:F,"Q3 2025")*-1', 'Q3 revenue at risk'],
    ['Q4 2025 Impact', '=SUMIFS(Looker_Ready_View!K:K,Looker_Ready_View!L:L,"Supply Gap",Looker_Ready_View!F:F,"Q4 2025")*-1', 'Q4 revenue at risk'],
    ['', '', ''],
    ['Data Quality:', '', ''],
    ['Total Records', '=COUNTA(Looker_Ready_View!A:A)-1', 'All processed records']
  ];
  
  const metricsRange = sheet.getRange(metricsStart + 1, 1, metrics.length, 3);
  metricsRange.setValues(metrics);
  
  // Format headers
  sheet.getRange(metricsStart + 1, 1, 1, 3).setBackground('#E8F0FE').setFontWeight('bold');
  
  // Format currency values
  sheet.getRange(metricsStart + 3, 2, 3, 1).setNumberFormat('$#,##0');
}

/**
 * Creates navigation links section
 */
function createNavigationLinks(sheet) {
  const linksStart = 18;
  sheet.getRange(`A${linksStart}`).setValue('DETAILED ANALYSIS DASHBOARDS');
  sheet.getRange(`A${linksStart}`).setFontSize(14).setFontWeight('bold').setBackground('#FF6D01').setFontColor('#FFFFFF');
  sheet.getRange(`A${linksStart}:D${linksStart}`).merge();
  
  const links = [
    ['Dashboard', 'Description', 'Click to Navigate', ''],
    ['Customer Analysis', 'Customer impact ranking with charts', '=HYPERLINK("#gid=" & RIGHT(CELL("address",\'Customer Analysis\'!A1),FIND("\'",CELL("address",\'Customer Analysis\'!A1)&"\'")-1), "Go to Customer Analysis")', ''],
    ['SKU Analysis', 'SKU impact ranking with charts', '=HYPERLINK("#gid=" & RIGHT(CELL("address",\'SKU Analysis\'!A1),FIND("\'",CELL("address",\'SKU Analysis\'!A1)&"\'")-1), "Go to SKU Analysis")', ''],
    ['Weekly Trends', 'Weekly gap trends with line charts', '=HYPERLINK("#gid=" & RIGHT(CELL("address",\'Weekly Trends\'!A1),FIND("\'",CELL("address",\'Weekly Trends\'!A1)&"\'")-1), "Go to Weekly Trends")', ''],
    ['Product Analysis', 'Product type analysis with pie charts', '=HYPERLINK("#gid=" & RIGHT(CELL("address",\'Product Analysis\'!A1),FIND("\'",CELL("address",\'Product Analysis\'!A1)&"\'")-1), "Go to Product Analysis")', ''],
    ['Supply Gap Summary', 'Comprehensive week-by-week summary chart', '=HYPERLINK("#gid=" & RIGHT(CELL("address",\'Supply Gap Summary\'!A1),FIND("\'",CELL("address",\'Supply Gap Summary\'!A1)&"\'")-1), "Go to Supply Gap Summary")', '']
  ];
  
  const linksRange = sheet.getRange(linksStart + 1, 1, links.length, 4);
  linksRange.setValues(links);
  
  // Format headers
  sheet.getRange(linksStart + 1, 1, 1, 4).setBackground('#FFF2CC').setFontWeight('bold');
}

/**
 * Creates all analysis dashboards
 */
function createAnalysisDashboards() {
  console.log('📈 Creating analysis dashboards...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source data
  const sourceSheet = ss.getSheetByName('Looker_Ready_View');
  const sourceData = sourceSheet ? sourceSheet.getDataRange().getValues() : null;
  
  console.log(`📊 Processing ${sourceData ? sourceData.length : 0} rows for analysis dashboards...`);
  
  // Create all analysis dashboards
  createCustomerAnalysisDashboard(ss, sourceData);
  createSKUAnalysisDashboard(ss, sourceData);
  createWeeklyTrendsDashboard(ss, sourceData);
  createProductAnalysisDashboard(ss, sourceData);
  createSupplyGapSummaryDashboard(ss, sourceData);
  
  console.log('✅ All analysis dashboards created!');
}

// ============================================================================
// CUSTOMER ANALYSIS DASHBOARD
// ============================================================================

/**
 * Creates customer analysis dashboard
 */
function createCustomerAnalysisDashboard(ss, sourceData) {
  console.log('📊 Creating Customer Analysis dashboard...');
  
  if (!sourceData || sourceData.length === 0) {
    console.log('No source data provided');
    return;
  }
  
  // Create or clear sheet
  let customerSheet = ss.getSheetByName('Customer Analysis');
  if (customerSheet) {
    ss.deleteSheet(customerSheet);
  }
  customerSheet = ss.insertSheet('Customer Analysis');
  
  // Add title and headers
  customerSheet.getRange('A1').setValue('CUSTOMER IMPACT ANALYSIS');
  customerSheet.getRange('A1').setFontSize(14).setFontWeight('bold').setBackground('#FF6D01').setFontColor('#FFFFFF');
  customerSheet.getRange('A1:E1').merge();
  
  const customerHeaders = ['Rank', 'Customer', 'Gap Units', 'Revenue Impact', 'SKUs Affected'];
  customerSheet.getRange('A3:E3').setValues([customerHeaders]);
  customerSheet.getRange('A3:E3').setBackground('#FFF2CC').setFontWeight('bold');
  
  // Use QUERY formulas for real-time data
  const queryFormula = `=QUERY(Looker_Ready_View!A:O, "SELECT A, SUM(J), SUM(K), COUNT(C) WHERE L = 'Supply Gap' AND A <> 'Customer' GROUP BY A ORDER BY SUM(K) DESC LIMIT 15", 1)`;
  
  customerSheet.getRange('B4').setFormula(queryFormula);
  
  // Add ranking formulas
  for (let i = 4; i <= 18; i++) {
    customerSheet.getRange(`A${i}`).setFormula(`=IF(B${i}<>"", ${i-3}, "")`);
  }
  
  // Format data
  customerSheet.getRange('C4:C18').setNumberFormat('#,##0');
  customerSheet.getRange('D4:D18').setNumberFormat('$#,##0');
  customerSheet.getRange('E4:E18').setNumberFormat('#,##0');
  
  // Create chart
  Utilities.sleep(2000);
  
  const chartRange = customerSheet.getRange('B4:D13');
  const chart = customerSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartRange)
    .setPosition(4, 7, 0, 0)
    .setOption('title', 'Top 10 Customers by Revenue Impact')
    .setOption('titleTextStyle', {fontSize: 14, bold: true})
    .setOption('height', 400)
    .setOption('width', 600)
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('vAxis.format', '$#,##0')
    .setOption('series.1.targetAxisIndex', 1)
    .setOption('vAxes.1.format', '$#,##0')
    .build();
  
  customerSheet.insertChart(chart);
  
  // Add summary statistics
  addCustomerSummaryStats(customerSheet);
  
  console.log('✓ Customer Analysis dashboard created');
}

/**
 * Adds summary statistics to customer dashboard
 */
function addCustomerSummaryStats(sheet) {
  sheet.getRange('A20').setValue('SUMMARY STATISTICS');
  sheet.getRange('A20').setFontWeight('bold').setBackground('#E8F0FE');
  sheet.getRange('A20:E20').merge();
  
  const summaryData = [
    ['Total Customers with Gaps:', '=COUNTA(B4:B18)-COUNTBLANK(B4:B18)', '', '', ''],
    ['Total Revenue at Risk:', '=SUM(D4:D18)', '', '', ''],
    ['Total Units at Risk:', '=SUM(C4:C18)', '', '', ''],
    ['Average Impact per Customer:', '=AVERAGE(D4:D18)', '', '', '']
  ];
  
  sheet.getRange('A21:E24').setValues(summaryData);
  sheet.getRange('B21:B24').setNumberFormat('$#,##0');
}

// ============================================================================
// SKU ANALYSIS DASHBOARD
// ============================================================================

/**
 * Creates SKU analysis dashboard
 */
function createSKUAnalysisDashboard(ss, sourceData) {
  console.log('📊 Creating SKU Analysis dashboard...');
  
  if (!sourceData || sourceData.length === 0) {
    console.log('No source data provided');
    return;
  }
  
  // Create or clear sheet
  let skuSheet = ss.getSheetByName('SKU Analysis');
  if (skuSheet) {
    ss.deleteSheet(skuSheet);
  }
  skuSheet = ss.insertSheet('SKU Analysis');
  
  // Add title and headers
  skuSheet.getRange('A1').setValue('SKU IMPACT ANALYSIS');
  skuSheet.getRange('A1').setFontSize(14).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  skuSheet.getRange('A1:F1').merge();
  
  const skuHeaders = ['Rank', 'SKU', 'PDT', 'Gap Units', 'Revenue Impact', 'Customers Affected'];
  skuSheet.getRange('A3:F3').setValues([skuHeaders]);
  skuSheet.getRange('A3:F3').setBackground('#E8F0FE').setFontWeight('bold');
  
  // Use QUERY formulas for real-time data
  const queryFormula = `=QUERY(Looker_Ready_View!A:O, "SELECT C, D, SUM(J), SUM(K), COUNT(A) WHERE L = 'Supply Gap' AND C <> 'Anker SKU' GROUP BY C, D ORDER BY SUM(K) DESC LIMIT 20", 1)`;
  
  skuSheet.getRange('B4').setFormula(queryFormula);
  
  // Add ranking formulas
  for (let i = 4; i <= 23; i++) {
    skuSheet.getRange(`A${i}`).setFormula(`=IF(B${i}<>"", ${i-3}, "")`);
  }
  
  // Format data
  skuSheet.getRange('D4:D23').setNumberFormat('#,##0');
  skuSheet.getRange('E4:E23').setNumberFormat('$#,##0');
  skuSheet.getRange('F4:F23').setNumberFormat('#,##0');
  
  // Create horizontal bar chart
  Utilities.sleep(2000);
  
  const chartRange = skuSheet.getRange('B4:C13');
  const revenueRange = skuSheet.getRange('E4:E13');
  
  const chart = skuSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(chartRange)
    .addRange(revenueRange)
    .setPosition(4, 8, 0, 0)
    .setOption('title', 'Top 10 SKUs by Revenue Impact')
    .setOption('titleTextStyle', {fontSize: 14, bold: true})
    .setOption('height', 500)
    .setOption('width', 700)
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('hAxis.format', '$#,##0')
    .setOption('chartArea', {left: 200, top: 50, width: '60%', height: '75%'})
    .build();
  
  skuSheet.insertChart(chart);
  
  // Add summary statistics
  addSKUSummaryStats(skuSheet);
  
  console.log('✓ SKU Analysis dashboard created');
}

/**
 * Adds summary statistics to SKU dashboard
 */
function addSKUSummaryStats(sheet) {
  sheet.getRange('A25').setValue('SKU SUMMARY STATISTICS');
  sheet.getRange('A25').setFontWeight('bold').setBackground('#E8F0FE');
  sheet.getRange('A25:F25').merge();
  
  const summaryData = [
    ['Total SKUs with Gaps:', '=COUNTA(B4:B23)-COUNTBLANK(B4:B23)', '', '', '', ''],
    ['Total Revenue at Risk:', '=SUM(E4:E23)', '', '', '', ''],
    ['Total Units at Risk:', '=SUM(D4:D23)', '', '', '', ''],
    ['Avg Revenue per SKU:', '=AVERAGE(E4:E23)', '', '', '', ''],
    ['Top SKU Impact:', '=MAX(E4:E23)', '', '', '', '']
  ];
  
  sheet.getRange('A26:F30').setValues(summaryData);
  sheet.getRange('B26:B30').setNumberFormat('$#,##0');
}

// ============================================================================
// WEEKLY TRENDS DASHBOARD
// ============================================================================

/**
 * Creates weekly trends dashboard
 */
function createWeeklyTrendsDashboard(ss, sourceData) {
  console.log('📊 Creating Weekly Trends dashboard...');
  
  if (!sourceData || sourceData.length === 0) {
    console.log('No source data provided');
    return;
  }
  
  // Create or clear sheet
  let weeklySheet = ss.getSheetByName('Weekly Trends');
  if (weeklySheet) {
    ss.deleteSheet(weeklySheet);
  }
  weeklySheet = ss.insertSheet('Weekly Trends');
  
  // Add title and headers
  weeklySheet.getRange('A1').setValue('WEEKLY TRENDS ANALYSIS');
  weeklySheet.getRange('A1').setFontSize(14).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  weeklySheet.getRange('A1:E1').merge();
  
  const weeklyHeaders = ['Week', 'Quarter', 'Gap Units', 'Revenue Impact', 'Gap Count'];
  weeklySheet.getRange('A3:E3').setValues([weeklyHeaders]);
  weeklySheet.getRange('A3:E3').setBackground('#E0F2F1').setFontWeight('bold');
  
  // Use QUERY formulas for real-time data
  const queryFormula = `=QUERY(Looker_Ready_View!A:O, "SELECT G, F, SUM(J), SUM(K), COUNT(L) WHERE L = 'Supply Gap' AND G <> 'Week' GROUP BY G, F ORDER BY G", 1)`;
  
  weeklySheet.getRange('A4').setFormula(queryFormula);
  
  // Format data
  weeklySheet.getRange('C4:C25').setNumberFormat('#,##0');
  weeklySheet.getRange('D4:D25').setNumberFormat('$#,##0');
  weeklySheet.getRange('E4:E25').setNumberFormat('#,##0');
  
  // Create line chart
  Utilities.sleep(2000);
  
  const chartRange = weeklySheet.getRange('A4:D25');
  const chart = weeklySheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartRange)
    .setPosition(4, 7, 0, 0)
    .setOption('title', 'Supply Gap Trends by Week')
    .setOption('titleTextStyle', {fontSize: 14, bold: true})
    .setOption('height', 400)
    .setOption('width', 800)
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('vAxis.format', '$#,##0')
    .setOption('curveType', 'function')
    .setOption('pointSize', 5)
    .setOption('series.0.color', '#FF6B6B')
    .setOption('series.1.color', '#4ECDC4')
    .setOption('legend.position', 'bottom')
    .build();
  
  weeklySheet.insertChart(chart);
  
  // Add summary statistics
  addWeeklySummaryStats(weeklySheet);
  
  console.log('✓ Weekly Trends dashboard created');
}

/**
 * Adds summary statistics to weekly trends dashboard
 */
function addWeeklySummaryStats(sheet) {
  sheet.getRange('A27').setValue('WEEKLY TREND STATISTICS');
  sheet.getRange('A27').setFontWeight('bold').setBackground('#E0F2F1');
  sheet.getRange('A27:E27').merge();
  
  const summaryData = [
    ['Total Weeks with Gaps:', '=COUNTA(A4:A25)-COUNTBLANK(A4:A25)', '', '', ''],
    ['Peak Week Revenue:', '=MAX(D4:D25)', '', '', ''],
    ['Peak Week Units:', '=MAX(C4:C25)', '', '', ''],
    ['Avg Weekly Impact:', '=AVERAGE(D4:D25)', '', '', ''],
    ['Total Impact (All Weeks):', '=SUM(D4:D25)', '', '', '']
  ];
  
  sheet.getRange('A28:E32').setValues(summaryData);
  sheet.getRange('B28:B32').setNumberFormat('$#,##0');
}

// ============================================================================
// PRODUCT ANALYSIS DASHBOARD
// ============================================================================

/**
 * Creates product analysis dashboard
 */
function createProductAnalysisDashboard(ss, sourceData) {
  console.log('📊 Creating Product Analysis dashboard...');
  
  if (!sourceData || sourceData.length === 0) {
    console.log('No source data provided');
    return;
  }
  
  // Create or clear sheet
  let pdtSheet = ss.getSheetByName('Product Analysis');
  if (pdtSheet) {
    ss.deleteSheet(pdtSheet);
  }
  pdtSheet = ss.insertSheet('Product Analysis');
  
  // Add title and headers
  pdtSheet.getRange('A1').setValue('PRODUCT TYPE ANALYSIS');
  pdtSheet.getRange('A1').setFontSize(14).setFontWeight('bold').setBackground('#9900FF').setFontColor('#FFFFFF');
  pdtSheet.getRange('A1:D1').merge();
  
  const pdtHeaders = ['Product Type', 'Gap Units', 'Revenue Impact', '% of Total Revenue'];
  pdtSheet.getRange('A3:D3').setValues([pdtHeaders]);
  pdtSheet.getRange('A3:D3').setBackground('#F3E5F5').setFontWeight('bold');
  
  // Use QUERY formulas for real-time data
  const queryFormula = `=QUERY(Looker_Ready_View!A:O, "SELECT D, SUM(J), SUM(K) WHERE L = 'Supply Gap' AND D <> 'PDT' GROUP BY D ORDER BY SUM(K) DESC", 1)`;
  
  pdtSheet.getRange('A4').setFormula(queryFormula);
  
  // Add percentage calculation formulas
  for (let i = 4; i <= 10; i++) {
    pdtSheet.getRange(`D${i}`).setFormula(`=IF(C${i}<>"", C${i}/SUM($C$4:$C$10), "")`);
  }
  
  // Format data
  pdtSheet.getRange('B4:B10').setNumberFormat('#,##0');
  pdtSheet.getRange('C4:C10').setNumberFormat('$#,##0');
  pdtSheet.getRange('D4:D10').setNumberFormat('0.0%');
  
  // Create pie chart
  Utilities.sleep(2000);
  
  const chartRange = pdtSheet.getRange('A4:C10');
  const chart = pdtSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(chartRange)
    .setPosition(4, 6, 0, 0)
    .setOption('title', 'Revenue Impact by Product Type')
    .setOption('titleTextStyle', {fontSize: 14, bold: true})
    .setOption('height', 400)
    .setOption('width', 600)
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('pieSliceText', 'percentage')
    .setOption('pieSliceTextStyle', {fontSize: 10})
    .setOption('chartArea', {left: 20, top: 50, width: '90%', height: '80%'})
    .setOption('legend.position', 'right')
    .setOption('legend.textStyle', {fontSize: 10})
    .build();
  
  pdtSheet.insertChart(chart);
  
  // Add summary statistics
  addProductSummaryStats(pdtSheet);
  
  console.log('✓ Product Analysis dashboard created');
}

/**
 * Adds summary statistics to product dashboard
 */
function addProductSummaryStats(sheet) {
  sheet.getRange('A12').setValue('PRODUCT SUMMARY STATISTICS');
  sheet.getRange('A12').setFontWeight('bold').setBackground('#F3E5F5');
  sheet.getRange('A12:D12').merge();
  
  const summaryData = [
    ['Total Product Types with Gaps:', '=COUNTA(A4:A10)-COUNTBLANK(A4:A10)', '', ''],
    ['Total Revenue at Risk:', '=SUM(C4:C10)', '', ''],
    ['Total Units at Risk:', '=SUM(B4:B10)', '', ''],
    ['Highest Impact Product:', '=INDEX(A4:A10, MATCH(MAX(C4:C10), C4:C10, 0))', '', ''],
    ['Highest Impact Amount:', '=MAX(C4:C10)', '', '']
  ];
  
  sheet.getRange('A13:D17').setValues(summaryData);
  sheet.getRange('B13:B17').setNumberFormat('$#,##0');
}

// ============================================================================
// SUPPLY GAP SUMMARY DASHBOARD
// ============================================================================

/**
 * Creates supply gap summary dashboard with stacked chart
 */
function createSupplyGapSummaryDashboard(ss, sourceData) {
  console.log('📊 Creating Supply Gap Summary dashboard...');
  
  if (!sourceData || sourceData.length === 0) {
    console.log('No source data provided');
    return;
  }
  
  // Create or clear sheet
  let sumSheet = ss.getSheetByName('Supply Gap Summary');
  if (sumSheet) {
    ss.deleteSheet(sumSheet);
  }
  sumSheet = ss.insertSheet('Supply Gap Summary');
  
  // Add title and filters
  sumSheet.getRange('A1').setValue('SUPPLY GAP SUMMARY - UNITS & REVENUE BY WEEK');
  sumSheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  sumSheet.getRange('A1:X1').merge();
  
  // Add filter controls
  addFilterControls(sumSheet);
  
  // Create weekly summary table
  createWeeklySummaryTable(sumSheet);
  
  // Create stacked chart
  createStackedChart(sumSheet);
  
  // Add summary statistics
  addSummaryStatistics(sumSheet);
  
  console.log('✓ Supply Gap Summary dashboard created');
}

/**
 * Adds filter controls to summary dashboard
 */
function addFilterControls(sheet) {
  sheet.getRange('A3').setValue('Filters:');
  sheet.getRange('A3').setFontWeight('bold').setBackground('#E8F0FE');
  
  sheet.getRange('B3').setValue('Customer:');
  sheet.getRange('C3').setValue('ALL');
  sheet.getRange('C3').setBackground('#FFF2CC');
  
  sheet.getRange('E3').setValue('SKU:');
  sheet.getRange('F3').setValue('ALL');
  sheet.getRange('F3').setBackground('#FFF2CC');
  
  sheet.getRange('A4').setValue('(Change filter values above to focus on specific Customer or SKU)');
  sheet.getRange('A4').setFontStyle('italic').setFontColor('#666666');
  sheet.getRange('A4:F4').merge();
}

/**
 * Creates weekly summary table
 */
function createWeeklySummaryTable(sheet) {
  // Headers
  const headers = ['Metric', 'Week →'];
  sheet.getRange('A6:B6').setValues([headers]);
  sheet.getRange('A6:B6').setBackground('#34A853').setFontColor('#FFFFFF').setFontWeight('bold');
  
  // Week headers
  const weeks = ['202529', '202530', '202531', '202532', '202533', '202534', '202535', '202536', '202537', '202538', '202539', '202540', '202541', '202542', '202543', '202544', '202545'];
  const weekHeaders = [...weeks, 'Total'];
  
  sheet.getRange(6, 3, 1, weekHeaders.length).setValues([weekHeaders]);
  sheet.getRange(6, 3, 1, weekHeaders.length).setBackground('#34A853').setFontColor('#FFFFFF').setFontWeight('bold');
  
  // Row labels
  const rowLabels = [
    ['Sum of Units'],
    ['Sum of Revenue'],
    ['Total Sum of Units'],
    ['Total Sum of Revenue']
  ];
  sheet.getRange('A7:A10').setValues(rowLabels);
  sheet.getRange('A7:A10').setBackground('#E8F0FE').setFontWeight('bold');
  
  // Create formulas for each week
  let col = 3;
  weeks.forEach(week => {
    // Units formula
    const unitsFormula = `=SUMIFS(Looker_Ready_View!J:J,Looker_Ready_View!L:L,"Supply Gap",Looker_Ready_View!G:G,${week})*-1`;
    sheet.getRange(7, col).setFormula(unitsFormula);
    
    // Revenue formula
    const revenueFormula = `=SUMIFS(Looker_Ready_View!K:K,Looker_Ready_View!L:L,"Supply Gap",Looker_Ready_View!G:G,${week})*-1`;
    sheet.getRange(8, col).setFormula(revenueFormula);
    
    col++;
  });
  
  // Total column
  const totalCol = col;
  sheet.getRange(7, totalCol).setFormula(`=SUM(C7:${String.fromCharCode(64 + col - 1)}7)`);
  sheet.getRange(8, totalCol).setFormula(`=SUM(C8:${String.fromCharCode(64 + col - 1)}8)`);
  
  // Total sum rows
  for (let i = 3; i <= totalCol; i++) {
    const columnLetter = String.fromCharCode(64 + i);
    sheet.getRange(9, i).setFormula(`=${columnLetter}7`);
    sheet.getRange(10, i).setFormula(`=${columnLetter}8`);
  }
  
  // Format numbers
  sheet.getRange(`C7:${String.fromCharCode(64 + totalCol)}7`).setNumberFormat('#,##0');
  sheet.getRange(`C8:${String.fromCharCode(64 + totalCol)}8`).setNumberFormat('$#,##0');
  sheet.getRange(`C9:${String.fromCharCode(64 + totalCol)}9`).setNumberFormat('#,##0');
  sheet.getRange(`C10:${String.fromCharCode(64 + totalCol)}10`).setNumberFormat('$#,##0');
}

/**
 * Creates stacked chart
 */
function createStackedChart(sheet) {
  Utilities.sleep(2000);
  
  const chartDataRange = sheet.getRange('A7:R8'); // Exclude total column
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartDataRange)
    .setPosition(12, 1, 0, 0)
    .setOption('title', 'Supply Gap Impact: Units & Revenue by Week')
    .setOption('titleTextStyle', {fontSize: 16, bold: true})
    .setOption('height', 400)
    .setOption('width', 1000)
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('isStacked', true)
    .setOption('vAxes.0.format', '#,##0')
    .setOption('vAxes.1.format', '$#,##0')
    .setOption('series.0.color', '#FF6B6B')
    .setOption('series.0.targetAxisIndex', 0)
    .setOption('series.1.color', '#4ECDC4')
    .setOption('series.1.targetAxisIndex', 1)
    .setOption('legend.position', 'top')
    .setOption('legend.textStyle', {fontSize: 12})
    .setOption('hAxis.title', 'Week')
    .setOption('vAxes.0.title', 'Units at Risk')
    .setOption('vAxes.1.title', 'Revenue at Risk ($)')
    .build();
  
  sheet.insertChart(chart);
}

/**
 * Adds summary statistics to summary dashboard
 */
function addSummaryStatistics(sheet) {
  sheet.getRange('A30').setValue('SUMMARY STATISTICS');
  sheet.getRange('A30').setFontWeight('bold').setBackground('#FF6D01').setFontColor('#FFFFFF');
  sheet.getRange('A30:D30').merge();
  
  const summaryStats = [
    ['Total Units at Risk:', '=S7', '', ''],
    ['Total Revenue at Risk:', '=S8', '', ''],
    ['Peak Week (Units):', '=INDEX(C6:R6,MATCH(MAX(C7:R7),C7:R7,0))', '', ''],
    ['Peak Week (Revenue):', '=INDEX(C6:R6,MATCH(MAX(C8:R8),C8:R8,0))', '', ''],
    ['Weeks with Gaps:', '=COUNTIF(C7:R7,">0")', '', '']
  ];
  
  sheet.getRange('A31:D35').setValues(summaryStats);
  sheet.getRange('B31:B35').setNumberFormat('#,##0');
}

// ============================================================================
// MENU SETUP
// ============================================================================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Anker Forecast Analysis')
    .addItem('🚀 Run Complete Analysis', 'runForecastAnalysisDashboard')
    .addSeparator()
    .addItem('📊 Transform Data Only', 'transformForecastData')
    .addItem('📈 Create Dashboards Only', 'createAnalysisDashboards')
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

/**
 * Shows information about the automation
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Anker Forecast Analysis Dashboard',
    'Professional automation tool for supply gap analysis.\n\nFeatures:\n• Automated data transformation\n• Executive dashboard\n• Customer impact analysis\n• SKU analysis\n• Weekly trends\n• Product analysis\n• Supply gap summary\n\nVersion 2.0 - Professional Edition',
    ui.ButtonSet.OK
  );
}
