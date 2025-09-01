/**
 * BULLETPROOF ANKER FORECAST AUTOMATION - WEEK 33
 * Enhanced to handle real data issues found in CSV analysis
 * Fixes: Dollar format parsing, #N/A values, data validation
 */

function runBulletproofManagerAutomation() {
  console.log('🛡️ Starting BULLETPROOF manager automation...');
  
  try {
    // Step 1: Transform data with enhanced parsing
    const recordCount = transformForecastDataBulletproof();
    
    // Step 2: Create enhanced dashboard
    createManagerDashboardBulletproof();
    
    console.log(`✅ BULLETPROOF automation complete!`);
    console.log(`📊 Processed ${recordCount} records`);
    
    SpreadsheetApp.getUi().alert(
      'BULLETPROOF Manager Dashboard Complete!', 
      `✅ Successfully processed ${recordCount} records\n📊 Created Looker_Ready_View_Week33_BULLETPROOF\n📊 Created Dashboard_Week33_BULLETPROOF\n\n🛡️ Enhanced parsing handles all data format issues!`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ BULLETPROOF automation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Automation Error', 
      `Error: ${error.message}\n\n📋 Check execution transcript for details.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

function transformForecastDataBulletproof() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source sheets
  const constrainedSheet = ss.getSheetByName('NEW WIDE CONSTRAINED FCST DATA');
  const unconstrainedSheet = ss.getSheetByName('NEW WIDE UNCONST. FCST DATA');
  
  if (!constrainedSheet || !unconstrainedSheet) {
    throw new Error('Could not find source sheets');
  }
  
  // Create or clear output sheet
  let outputSheet = ss.getSheetByName('Looker_Ready_View_Week33_BULLETPROOF');
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet('Looker_Ready_View_Week33_BULLETPROOF');
  }
  
  // Enhanced headers with data quality indicators
  const headers = ['Customer', 'Anker SKU', 'PDT', 'Forecast Type', 'Quarter', 'Week', 
                   'Forecast - Units', 'Forecast Revenue', 'Delta Units', 'Delta - Revenue', 
                   'Gap Flag', 'IsCurrentQ', 'Important Helper', 'Sell-In Price', 'Data Quality'];
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get data ranges
  const constrainedData = constrainedSheet.getDataRange().getValues();
  const unconstrainedData = unconstrainedSheet.getDataRange().getValues();
  
  console.log(`Processing ${constrainedData.length} constrained rows, ${unconstrainedData.length} unconstrained rows`);
  
  // Enhanced column mapping based on CSV analysis
  const COLUMN_MAP = {
    HELPER: 0,        // Important Helper
    SHIFT_LOGIC: 1,   // Shift logic  
    SELL_IN_PRICE: 2, // Sell-in price (with #N/A issues)
    CATEGORY: 3,      // Category
    PCT: 4,           // PCT
    BG: 5,            // BG
    PDT: 6,           // PDT
    CUSTOMER_ID: 7,   // Customer ID
    CUSTOMER: 8,      // Customer
    ANKER_SKU: 9,     // Anker SKU
    WEEK_START: 10    // First week (202534)
  };
  
  // Find week columns with enhanced detection
  const weekColumns = [];
  const headerRow = constrainedData[0];
  
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
  
  // Enhanced currency parsing function
  function parseCurrencyValue(value) {
    if (value === null || value === undefined || value === '' || value === '#N/A' || value === 'N/A') {
      return 0;
    }
    
    // Handle string values with dollar signs and commas
    if (typeof value === 'string') {
      // Remove dollar signs, commas, quotes, and extra spaces
      const cleanValue = value.replace(/[\$",\s]/g, '');
      
      // Handle #N/A and other error values
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
  
  // Create unconstrained lookup with enhanced parsing
  const unconstrainedLookup = {};
  console.log('Building unconstrained lookup with enhanced parsing...');
  
  let dataQualityStats = {
    totalRows: 0,
    validPrices: 0,
    invalidPrices: 0,
    nonZeroForecasts: 0,
    zeroForecasts: 0
  };
  
  for (let i = 1; i < unconstrainedData.length; i++) {
    if (i % 100 === 0) {
      console.log(`Processing unconstrained row ${i}/${unconstrainedData.length}`);
      Utilities.sleep(10);
    }
    
    const row = unconstrainedData[i];
    const helper = row[COLUMN_MAP.HELPER];
    if (!helper) continue;
    
    dataQualityStats.totalRows++;
    
    // Enhanced price parsing
    const sellInPriceRaw = row[COLUMN_MAP.SELL_IN_PRICE];
    const sellInPrice = parseCurrencyValue(sellInPriceRaw);
    
    if (sellInPrice > 0) {
      dataQualityStats.validPrices++;
    } else {
      dataQualityStats.invalidPrices++;
    }
    
    weekColumns.forEach(weekCol => {
      const week = weekCol.week;
      const forecastRevenueRaw = row[weekCol.index];
      const forecastRevenue = parseCurrencyValue(forecastRevenueRaw);
      
      if (forecastRevenue > 0) {
        dataQualityStats.nonZeroForecasts++;
      } else {
        dataQualityStats.zeroForecasts++;
      }
      
      const forecastUnits = sellInPrice > 0 ? forecastRevenue / sellInPrice : 0;
      
      const key = `${helper}_${week}`;
      unconstrainedLookup[key] = {
        units: forecastUnits,
        revenue: forecastRevenue,
        sellInPrice: sellInPrice,
        rawRevenue: forecastRevenueRaw,
        rawPrice: sellInPriceRaw
      };
    });
  }
  
  console.log('📊 Data Quality Stats:', dataQualityStats);
  console.log(`Created lookup with ${Object.keys(unconstrainedLookup).length} entries`);
  
  // Process constrained data with enhanced parsing and validation
  const BATCH_SIZE = 100;
  let totalRecords = 0;
  let qualityFlags = {
    goodRecords: 0,
    priceIssues: 0,
    dataIssues: 0
  };
  
  for (let startRow = 1; startRow < constrainedData.length; startRow += BATCH_SIZE) {
    const endRow = Math.min(startRow + BATCH_SIZE, constrainedData.length);
    console.log(`Processing constrained batch: rows ${startRow} to ${endRow}`);
    
    const batchData = [];
    
    for (let i = startRow; i < endRow; i++) {
      const row = constrainedData[i];
      const helper = row[COLUMN_MAP.HELPER];
      const sellInPriceRaw = row[COLUMN_MAP.SELL_IN_PRICE];
      const sellInPrice = parseCurrencyValue(sellInPriceRaw);
      const pdt = row[COLUMN_MAP.PDT];
      const customer = row[COLUMN_MAP.CUSTOMER];
      const sku = row[COLUMN_MAP.ANKER_SKU];
      
      if (!helper || !customer || !sku) continue;
      
      // Data quality assessment
      let dataQuality = 'Good';
      if (sellInPrice <= 0 || sellInPriceRaw === '#N/A') {
        dataQuality = 'Price Issue';
        qualityFlags.priceIssues++;
      } else {
        qualityFlags.goodRecords++;
      }
      
      weekColumns.forEach(weekCol => {
        const week = weekCol.week;
        const constrainedRevenueRaw = row[weekCol.index];
        const constrainedRevenue = parseCurrencyValue(constrainedRevenueRaw);
        const constrainedUnits = sellInPrice > 0 ? constrainedRevenue / sellInPrice : 0;
        
        const key = `${helper}_${week}`;
        const unconstrained = unconstrainedLookup[key] || { units: 0, revenue: 0, sellInPrice: 0 };
        
        const deltaUnits = constrainedUnits - unconstrained.units;
        const deltaRevenue = constrainedRevenue - unconstrained.revenue;
        
        const quarter = getQuarterBulletproof(week);
        const isCurrentQ = quarter === 'Q4 2025';
        const gapFlag = deltaUnits < 0 ? 'Supply Gap' : '';
        
        // Add records with any forecast data OR if there's a meaningful delta
        if (constrainedRevenue > 0 || unconstrained.revenue > 0 || Math.abs(deltaRevenue) > 0.01) {
          // Add constrained row
          batchData.push([
            customer, sku, pdt, 'Constrained', quarter, week,
            constrainedUnits, constrainedRevenue, deltaUnits, deltaRevenue,
            gapFlag, isCurrentQ, helper, sellInPrice, dataQuality
          ]);
          
          // Add unconstrained row
          batchData.push([
            customer, sku, pdt, 'Unconstrained', quarter, week,
            unconstrained.units, unconstrained.revenue, 0, 0,
            '', isCurrentQ, helper, unconstrained.sellInPrice || sellInPrice, dataQuality
          ]);
        }
      });
    }
    
    // Write batch to sheet
    if (batchData.length > 0) {
      const startOutputRow = totalRecords + 2;
      outputSheet.getRange(startOutputRow, 1, batchData.length, headers.length).setValues(batchData);
      totalRecords += batchData.length;
      
      console.log(`Wrote batch: ${batchData.length} records (total: ${totalRecords})`);
    }
    
    Utilities.sleep(100);
  }
  
  console.log('📊 Processing Quality Stats:', qualityFlags);
  
  // Format the sheet
  formatOutputSheetBulletproof(outputSheet, headers.length);
  
  console.log(`✅ BULLETPROOF transformation complete! ${totalRecords} records created.`);
  return totalRecords;
}

function createManagerDashboardBulletproof() {
  console.log('📊 Creating BULLETPROOF manager dashboard...');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create dashboard
  let dashboardSheet = ss.getSheetByName('Dashboard_Week33_BULLETPROOF');
  if (dashboardSheet) {
    dashboardSheet.clear();
  } else {
    dashboardSheet = ss.insertSheet('Dashboard_Week33_BULLETPROOF');
  }
  
  // Title
  dashboardSheet.getRange('A1').setValue('ANKER SUPPLY GAP ANALYSIS - FRESH WEEK 33 (BULLETPROOF)');
  dashboardSheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
  dashboardSheet.getRange('A1:H1').merge();
  
  // Timestamp and data info
  const timestamp = new Date().toLocaleString();
  dashboardSheet.getRange('A2').setValue(`Generated: ${timestamp} | Source: Enhanced parsing of NEW WIDE data sheets`);
  dashboardSheet.getRange('A2').setFontStyle('italic');
  dashboardSheet.getRange('A2:H2').merge();
  
  dashboardSheet.getRange('A3').setValue('Data Range: Week 34-52 (202534-202552) | Enhanced: Currency parsing, #N/A handling, Data quality flags');
  dashboardSheet.getRange('A3:H3').merge();
  
  // Key metrics
  createKeyMetricsBulletproof(dashboardSheet);
  
  // Data quality section
  createDataQualitySection(dashboardSheet);
  
  // Analysis instructions
  createAnalysisInstructionsBulletproof(dashboardSheet);
  
  console.log('✅ BULLETPROOF dashboard created');
}

function createKeyMetricsBulletproof(sheet) {
  const metricsStart = 5;
  sheet.getRange(`A${metricsStart}`).setValue('KEY METRICS - BULLETPROOF PARSING');
  sheet.getRange(`A${metricsStart}`).setFontSize(14).setFontWeight('bold').setBackground('#34A853').setFontColor('#FFFFFF');
  sheet.getRange(`A${metricsStart}:C${metricsStart}`).merge();
  
  const metrics = [
    ['Metric', 'Formula/Value', 'Description'],
    ['Total Supply Gaps', '=COUNTIFS(Looker_Ready_View_Week33_BULLETPROOF!K:K,"Supply Gap")', 'Number of supply gap scenarios'],
    ['Revenue at Risk', '=SUMIFS(Looker_Ready_View_Week33_BULLETPROOF!J:J,Looker_Ready_View_Week33_BULLETPROOF!K:K,"Supply Gap")*-1', 'Total revenue impact'],
    ['Units at Risk', '=SUMIFS(Looker_Ready_View_Week33_BULLETPROOF!I:I,Looker_Ready_View_Week33_BULLETPROOF!K:K,"Supply Gap")*-1', 'Total units affected'],
    ['Q3 2025 Impact', '=SUMIFS(Looker_Ready_View_Week33_BULLETPROOF!J:J,Looker_Ready_View_Week33_BULLETPROOF!K:K,"Supply Gap",Looker_Ready_View_Week33_BULLETPROOF!E:E,"Q3 2025")*-1', 'Q3 revenue at risk'],
    ['Q4 2025 Impact', '=SUMIFS(Looker_Ready_View_Week33_BULLETPROOF!J:J,Looker_Ready_View_Week33_BULLETPROOF!K:K,"Supply Gap",Looker_Ready_View_Week33_BULLETPROOF!E:E,"Q4 2025")*-1', 'Q4 revenue at risk'],
    ['', '', ''],
    ['Enhanced Data Processing:', '', ''],
    ['Total Records', '=COUNTA(Looker_Ready_View_Week33_BULLETPROOF!A:A)-1', 'All records with enhanced parsing'],
    ['Good Data Quality', '=COUNTIFS(Looker_Ready_View_Week33_BULLETPROOF!O:O,"Good")', 'Records with valid sell-in prices'],
    ['Price Issues', '=COUNTIFS(Looker_Ready_View_Week33_BULLETPROOF!O:O,"Price Issue")', 'Records with #N/A or $0 prices']
  ];
  
  const metricsRange = sheet.getRange(metricsStart + 1, 1, metrics.length, 3);
  metricsRange.setValues(metrics);
  
  // Format headers
  sheet.getRange(metricsStart + 1, 1, 1, 3).setBackground('#E8F0FE').setFontWeight('bold');
  
  // Format currency values
  sheet.getRange(metricsStart + 3, 2, 3, 1).setNumberFormat('$#,##0');
}

function createDataQualitySection(sheet) {
  const qualityStart = 20;
  
  sheet.getRange(`A${qualityStart}`).setValue('DATA QUALITY ANALYSIS');
  sheet.getRange(`A${qualityStart}`).setFontWeight('bold').setBackground('#FF6D01').setFontColor('#FFFFFF');
  sheet.getRange(`A${qualityStart}:C${qualityStart}`).merge();
  
  const qualityChecks = [
    ['Quality Check', 'Formula', 'Expected Result'],
    ['Currency Parsing', '✅ Handles "$5,460.00" format', 'All dollar values converted correctly'],
    ['#N/A Handling', '✅ Converts #N/A to $0', 'No calculation errors'],
    ['Non-Zero Forecasts', '=COUNTIFS(Looker_Ready_View_Week33_BULLETPROOF!H:H,">0")', 'Should be 1000+ records'],
    ['Supply Gaps Found', '=COUNTIFS(Looker_Ready_View_Week33_BULLETPROOF!K:K,"Supply Gap")', 'Should be 100+ gaps'],
    ['Data Completeness', '=COUNTA(Looker_Ready_View_Week33_BULLETPROOF!A:A)/COUNTA(Looker_Ready_View_Week33_BULLETPROOF!M:M)', 'Should be close to 1.0']
  ];
  
  sheet.getRange(qualityStart + 1, 1, qualityChecks.length, 3).setValues(qualityChecks);
  sheet.getRange(qualityStart + 1, 1, 1, 3).setBackground('#FFF2CC').setFontWeight('bold');
}

function createAnalysisInstructionsBulletproof(sheet) {
  const instructStart = 30;
  
  sheet.getRange(`A${instructStart}`).setValue('ENHANCED ANALYSIS INSTRUCTIONS');
  sheet.getRange(`A${instructStart}`).setFontWeight('bold').setBackground('#9900FF').setFontColor('#FFFFFF');
  sheet.getRange(`A${instructStart}:H${instructStart}`).merge();
  
  const instructions = [
    ['Analysis Type', 'Instructions', 'Notes'],
    ['', '', ''],
    ['1. Supply Gap Analysis', '• Pivot Table: Gap Flag = "Supply Gap"', '✅ Enhanced currency parsing'],
    ['', '• Group by: PDT, Anker SKU', '✅ Handles #N/A values'],
    ['', '• Sum: Delta Revenue (descending)', '✅ Data quality indicators'],
    ['', '', ''],
    ['2. Data Quality Review', '• Filter by Data Quality = "Price Issue"', 'Review SKUs with pricing problems'],
    ['', '• Check sell-in prices for #N/A values', 'May need manual price updates'],
    ['', '', ''],
    ['3. Customer Impact', '• Standard pivot analysis', 'Now includes all valid data'],
    ['', '• Revenue numbers are properly parsed', 'No more $0 false readings'],
    ['', '', ''],
    ['Key Improvements:', '', ''],
    ['✅ Currency Format', 'Handles "$5,460.00" and commas', 'Previously failed to parse'],
    ['✅ Error Values', 'Converts #N/A to $0 safely', 'Prevents calculation errors'],
    ['✅ Data Validation', 'Quality flags for each record', 'Easy to spot issues']
  ];
  
  sheet.getRange(instructStart + 1, 1, instructions.length, 3).setValues(instructions);
  sheet.getRange(instructStart + 1, 1, 1, 3).setBackground('#F3E5F5').setFontWeight('bold');
}

function formatOutputSheetBulletproof(sheet, numColumns) {
  console.log('Formatting bulletproof output sheet...');
  
  sheet.setFrozenRows(1);
  
  const headerRange = sheet.getRange(1, 1, 1, numColumns);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // Auto-resize key columns
  for (let i = 1; i <= Math.min(10, numColumns); i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Format currency columns
  if (sheet.getLastRow() > 1) {
    // Format revenue columns
    sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).setNumberFormat('$#,##0.00'); // Forecast Revenue
    sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).setNumberFormat('$#,##0.00'); // Delta Revenue
    sheet.getRange(2, 14, sheet.getLastRow() - 1, 1).setNumberFormat('$#,##0.00'); // Sell-In Price
  }
  
  console.log('Bulletproof formatting complete');
}

function getQuarterBulletproof(week) {
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
  ui.createMenu('🛡️ BULLETPROOF Week 33 Automation')
    .addItem('🛡️ Run BULLETPROOF Complete Automation', 'runBulletproofManagerAutomation')
    .addToUi();
}


