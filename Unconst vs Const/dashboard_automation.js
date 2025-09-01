/**
 * DASHBOARD AUTOMATION FOR GOOGLE SHEETS
 * Creates automated charts, pivot tables, and dashboard layouts
 * Run after transformForecastData() completes
 */

function createFullDashboard() {
  console.log('Creating full automated dashboard...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ensure we have the data
  const dataSheet = ss.getSheetByName('Looker_Ready_View_New');
  if (!dataSheet || dataSheet.getLastRow() <= 1) {
    throw new Error('No data found in Looker_Ready_View_New. Run transformForecastData() first.');
  }
  
  // Create all components
  createAutomatedPivotTables(ss);
  createAutomatedCharts(ss);
  createDashboardLayout(ss);
  
  console.log('✅ Dashboard automation complete!');
}

function createAutomatedPivotTables(ss) {
  console.log('Creating automated pivot tables...');
  
  const sourceSheet = ss.getSheetByName('Looker_Ready_View_New');
  const sourceRange = sourceSheet.getDataRange();
  
  // SKU Summary Pivot Table
  createSKUPivotTable(ss, sourceRange);
  
  // Customer Summary Pivot Table
  createCustomerPivotTable(ss, sourceRange);
  
  // Weekly Trends Pivot Table
  createWeeklyPivotTable(ss, sourceRange);
  
  // PDT Summary Pivot Table
  createPDTPivotTable(ss, sourceRange);
}

function createSKUPivotTable(ss, sourceRange) {
  let sheet = ss.getSheetByName('SKU_Analysis_Auto');
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet('SKU_Analysis_Auto');
  
  const pivotTable = sheet.getRange('A1').createPivotTable(sourceRange);
  
  // Configure pivot table
  pivotTable.addRowGroup('Anker SKU');
  pivotTable.addRowGroup('PDT');
  
  // Add filters
  const gapFlagGroup = pivotTable.addFilter('Gap Flag');
  gapFlagGroup.filterCriteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues(['Supply Gap'])
    .build();
  
  // Add values
  pivotTable.addPivotValue('Delta Units', SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotTable.addPivotValue('Delta - Revenue', SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotTable.addPivotValue('Customer', SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  
  // Format the sheet
  formatPivotSheet(sheet, 'SKU Analysis - Supply Gaps');
  
  console.log('✓ Created SKU pivot table');
}

function createCustomerPivotTable(ss, sourceRange) {
  let sheet = ss.getSheetByName('Customer_Analysis_Auto');
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet('Customer_Analysis_Auto');
  
  const pivotTable = sheet.getRange('A1').createPivotTable(sourceRange);
  
  pivotTable.addRowGroup('Customer');
  
  const gapFlagGroup = pivotTable.addFilter('Gap Flag');
  gapFlagGroup.filterCriteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues(['Supply Gap'])
    .build();
  
  pivotTable.addPivotValue('Delta Units', SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotTable.addPivotValue('Delta - Revenue', SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotTable.addPivotValue('Anker SKU', SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  
  formatPivotSheet(sheet, 'Customer Analysis - Supply Gaps');
  
  console.log('✓ Created Customer pivot table');
}

function createWeeklyPivotTable(ss, sourceRange) {
  let sheet = ss.getSheetByName('Weekly_Analysis_Auto');
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet('Weekly_Analysis_Auto');
  
  const pivotTable = sheet.getRange('A1').createPivotTable(sourceRange);
  
  pivotTable.addRowGroup('Week');
  pivotTable.addRowGroup('Quarter');
  
  const gapFlagGroup = pivotTable.addFilter('Gap Flag');
  gapFlagGroup.filterCriteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues(['Supply Gap'])
    .build();
  
  pivotTable.addPivotValue('Delta Units', SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotTable.addPivotValue('Delta - Revenue', SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotTable.addPivotValue('Customer', SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  
  formatPivotSheet(sheet, 'Weekly Trends - Supply Gaps');
  
  console.log('✓ Created Weekly pivot table');
}

function createPDTPivotTable(ss, sourceRange) {
  let sheet = ss.getSheetByName('PDT_Analysis_Auto');
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet('PDT_Analysis_Auto');
  
  const pivotTable = sheet.getRange('A1').createPivotTable(sourceRange);
  
  pivotTable.addRowGroup('PDT');
  
  const gapFlagGroup = pivotTable.addFilter('Gap Flag');
  gapFlagGroup.filterCriteria = SpreadsheetApp.newFilterCriteria()
    .setVisibleValues(['Supply Gap'])
    .build();
  
  pivotTable.addPivotValue('Delta Units', SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotTable.addPivotValue('Delta - Revenue', SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotTable.addPivotValue('Customer', SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  
  formatPivotSheet(sheet, 'PDT Analysis - Supply Gaps');
  
  console.log('✓ Created PDT pivot table');
}

function formatPivotSheet(sheet, title) {
  // Add title
  sheet.insertRowBefore(1);
  sheet.getRange('A1').setValue(title);
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  
  // Format headers
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow > 2) {
    const headerRange = sheet.getRange(2, 1, 1, lastCol);
    headerRange.setBackground('#4285F4');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
  }
  
  // Auto-resize columns
  for (let i = 1; i <= lastCol; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Format currency columns
  if (lastCol >= 3) {
    const revenueRange = sheet.getRange(3, 3, lastRow - 2, 1);
    revenueRange.setNumberFormat('$#,##0.00');
  }
}

function createAutomatedCharts(ss) {
  console.log('Creating automated charts...');
  
  // Create charts sheet
  let chartsSheet = ss.getSheetByName('Charts_Dashboard');
  if (chartsSheet) {
    ss.deleteSheet(chartsSheet);
  }
  chartsSheet = ss.insertSheet('Charts_Dashboard');
  
  // Create individual charts
  createSKUChart(ss, chartsSheet);
  createCustomerChart(ss, chartsSheet);
  createWeeklyTrendChart(ss, chartsSheet);
  createPDTChart(ss, chartsSheet);
  
  console.log('✓ Created automated charts');
}

function createSKUChart(ss, targetSheet) {
  const skuSheet = ss.getSheetByName('SKU_Analysis_Auto');
  if (!skuSheet) return;
  
  const dataRange = skuSheet.getDataRange();
  
  const chart = targetSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setPosition(2, 1, 0, 0)
    .setOption('title', 'Top SKUs by Revenue Impact')
    .setOption('titleTextStyle', {fontSize: 14, bold: true})
    .setOption('height', 300)
    .setOption('width', 600)
    .setOption('backgroundColor', '#FFFFFF')
    .build();
  
  targetSheet.insertChart(chart);
}

function createCustomerChart(ss, targetSheet) {
  const customerSheet = ss.getSheetByName('Customer_Analysis_Auto');
  if (!customerSheet) return;
  
  const dataRange = customerSheet.getDataRange();
  
  const chart = targetSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setPosition(2, 8, 0, 0)
    .setOption('title', 'Top Customers by Revenue Impact')
    .setOption('titleTextStyle', {fontSize: 14, bold: true})
    .setOption('height', 300)
    .setOption('width', 600)
    .build();
  
  targetSheet.insertChart(chart);
}

function createWeeklyTrendChart(ss, targetSheet) {
  const weeklySheet = ss.getSheetByName('Weekly_Analysis_Auto');
  if (!weeklySheet) return;
  
  const dataRange = weeklySheet.getDataRange();
  
  const chart = targetSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(dataRange)
    .setPosition(18, 1, 0, 0)
    .setOption('title', 'Weekly Revenue Impact Trend')
    .setOption('titleTextStyle', {fontSize: 14, bold: true})
    .setOption('height', 300)
    .setOption('width', 600)
    .build();
  
  targetSheet.insertChart(chart);
}

function createPDTChart(ss, targetSheet) {
  const pdtSheet = ss.getSheetByName('PDT_Analysis_Auto');
  if (!pdtSheet) return;
  
  const dataRange = pdtSheet.getDataRange();
  
  const chart = targetSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange)
    .setPosition(18, 8, 0, 0)
    .setOption('title', 'Revenue Impact by Product Type')
    .setOption('titleTextStyle', {fontSize: 14, bold: true})
    .setOption('height', 300)
    .setOption('width', 600)
    .build();
  
  targetSheet.insertChart(chart);
}

function createDashboardLayout(ss) {
  console.log('Creating dashboard layout...');
  
  let dashboard = ss.getSheetByName('Executive_Dashboard');
  if (dashboard) {
    ss.deleteSheet(dashboard);
  }
  dashboard = ss.insertSheet('Executive_Dashboard');
  
  // Create dashboard structure
  createDashboardHeader(dashboard);
  createKPISection(dashboard);
  addDashboardInstructions(dashboard);
  
  console.log('✓ Created executive dashboard');
}

function createDashboardHeader(sheet) {
  // Title
  sheet.getRange('A1').setValue('ANKER SUPPLY GAP ANALYSIS DASHBOARD');
  sheet.getRange('A1').setFontSize(18).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  sheet.getRange('A1:H1').merge();
  
  // Subtitle
  const timestamp = new Date().toLocaleString();
  sheet.getRange('A2').setValue(`Last Updated: ${timestamp}`);
  sheet.getRange('A2').setFontStyle('italic');
  sheet.getRange('A2:H2').merge();
}

function createKPISection(sheet) {
  const sourceSheet = sheet.getParent().getSheetByName('Looker_Ready_View_New');
  if (!sourceSheet) return;
  
  // KPI Headers
  const kpiHeaders = ['METRIC', 'VALUE', 'DESCRIPTION'];
  sheet.getRange('A4:C4').setValues([kpiHeaders]);
  sheet.getRange('A4:C4').setBackground('#34A853').setFontColor('#FFFFFF').setFontWeight('bold');
  
  // KPI Formulas
  const kpis = [
    ['Total Supply Gaps', '=COUNTIFS(Looker_Ready_View_New!K:K,"Supply Gap")', 'Number of SKU-Customer-Week combinations with supply gaps'],
    ['Revenue at Risk', '=SUMIFS(Looker_Ready_View_New!J:J,Looker_Ready_View_New!K:K,"Supply Gap")*-1', 'Total revenue impact from supply gaps'],
    ['Units at Risk', '=SUMIFS(Looker_Ready_View_New!I:I,Looker_Ready_View_New!K:K,"Supply Gap")*-1', 'Total units affected by supply gaps'],
    ['Q3 2025 Impact', '=SUMIFS(Looker_Ready_View_New!J:J,Looker_Ready_View_New!K:K,"Supply Gap",Looker_Ready_View_New!E:E,"Q3 2025")*-1', 'Revenue impact in Q3 2025'],
    ['Q4 2025 Impact', '=SUMIFS(Looker_Ready_View_New!J:J,Looker_Ready_View_New!K:K,"Supply Gap",Looker_Ready_View_New!E:E,"Q4 2025")*-1', 'Revenue impact in Q4 2025'],
    ['Affected SKUs', '=SUMPRODUCT((Looker_Ready_View_New!K:K="Supply Gap")*(COUNTIFS(Looker_Ready_View_New!B:B,Looker_Ready_View_New!B:B,Looker_Ready_View_New!K:K,"Supply Gap")>0))', 'Unique SKUs with supply gaps'],
    ['Affected Customers', '=SUMPRODUCT((Looker_Ready_View_New!K:K="Supply Gap")*(COUNTIFS(Looker_Ready_View_New!A:A,Looker_Ready_View_New!A:A,Looker_Ready_View_New!K:K,"Supply Gap")>0))', 'Unique customers affected']
  ];
  
  for (let i = 0; i < kpis.length; i++) {
    const row = 5 + i;
    sheet.getRange(row, 1, 1, 3).setValues([kpis[i]]);
  }
  
  // Format KPI values
  sheet.getRange('B5:B11').setNumberFormat('#,##0');
  sheet.getRange('B6:B8').setNumberFormat('$#,##0');
  
  // Conditional formatting for high impact values
  const revenueRange = sheet.getRange('B6:B8');
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(1000000)
    .setBackground('#FFE6E6')
    .setFontColor('#CC0000')
    .setRanges([revenueRange])
    .build();
  sheet.setConditionalFormatRules([rule]);
}

function addDashboardInstructions(sheet) {
  const instructions = [
    ['NAVIGATION GUIDE', '', ''],
    ['Sheet', 'Purpose', 'Description'],
    ['Looker_Ready_View_New', 'Raw Data', 'Complete transformed dataset'],
    ['SKU_Analysis_Auto', 'SKU Analysis', 'Supply gaps by SKU with pivot table'],
    ['Customer_Analysis_Auto', 'Customer Analysis', 'Supply gaps by customer'],
    ['Weekly_Analysis_Auto', 'Trend Analysis', 'Weekly trends and patterns'],
    ['PDT_Analysis_Auto', 'Product Analysis', 'Analysis by product type'],
    ['Charts_Dashboard', 'Visualizations', 'All automated charts'],
    ['Executive_Dashboard', 'Summary', 'This summary dashboard']
  ];
  
  const startRow = 14;
  sheet.getRange(startRow, 1, instructions.length, 3).setValues(instructions);
  
  // Format instructions
  sheet.getRange(startRow, 1, 1, 3).setBackground('#FF9900').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(startRow + 1, 1, 1, 3).setBackground('#FFF2CC').setFontWeight('bold');
  
  // Auto-resize columns
  for (let i = 1; i <= 8; i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Complete automation workflow
 */
function runCompleteAutomation() {
  console.log('🚀 Starting complete forecast automation...');
  
  try {
    // Step 1: Transform data
    console.log('Step 1: Transforming forecast data...');
    const recordCount = transformForecastData();
    
    // Step 2: Create dashboard
    console.log('Step 2: Creating automated dashboard...');
    createFullDashboard();
    
    console.log(`✅ Complete automation finished successfully!`);
    console.log(`📊 Processed ${recordCount} records`);
    console.log('🎯 Navigate to Executive_Dashboard sheet to view results');
    
    // Show completion message
    SpreadsheetApp.getUi().alert(
      'Automation Complete!', 
      `Successfully processed ${recordCount} records.\n\nNavigate to the Executive_Dashboard sheet to view your results.\n\nAll pivot tables and charts have been automatically created.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ Automation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Automation Error', 
      `Error: ${error.message}\n\nPlease check the console logs for details.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Menu setup function
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Anker Automation')
    .addItem('🚀 Run Complete Automation', 'runCompleteAutomation')
    .addSeparator()
    .addItem('📊 Transform Data Only', 'transformForecastData')
    .addItem('📈 Create Dashboard Only', 'createFullDashboard')
    .addSeparator()
    .addItem('🧪 Test Sheet Access', 'testSheetAccess')
    .addItem('🧹 Cleanup Old Sheets', 'cleanupOldSheets')
    .addToUi();
}
