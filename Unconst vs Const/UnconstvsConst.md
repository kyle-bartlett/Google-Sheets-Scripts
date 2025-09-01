function transformForecastData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get source sheets
  const constrainedSheet = ss.getSheetByName('Constrained Forecast - "Confirm');
  const unconstrainedSheet = ss.getSheetByName('Unconstrained Forecast - "Optim');
  
  // Create or clear output sheet
  let outputSheet = ss.getSheetByName('Looker_Ready_View');
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = ss.insertSheet('Looker_Ready_View');
  }
  
  // Headers for output
  const headers = ['Customer', 'Anker SKU', 'PDT', 'Forecast Type', 'Quarter', 'Week', 
                   'Forecast - Units', 'Forecast Revenue', 'Delta Units', 'Delta - Revenue', 
                   'Gap Flag', 'IsCurrentQ'];
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Get data ranges
  const constrainedData = constrainedSheet.getDataRange().getValues();
  const unconstrainedData = unconstrainedSheet.getDataRange().getValues();
  
  // Find week columns (start from column K = index 10)
  const weekStartIndex = 10;
  const weekEndIndex = constrainedData[0].length - 1;
  
  // Create unconstrained lookup
  const unconstrainedLookup = {};
  for (let i = 1; i < unconstrainedData.length; i++) {
    const row = unconstrainedData[i];
    const helper = row[0]; // Important Helper
    for (let j = weekStartIndex; j <= weekEndIndex; j++) {
      const week = constrainedData[0][j]; // Get week from header
      if (week && week.toString().match(/202\d{3}/)) {
        const key = `${helper}_${week}`;
        unconstrainedLookup[key] = {
          units: parseFloat(row[j]) || 0,
          revenue: (parseFloat(row[j]) || 0) * (parseFloat(row[2]) || 0) // units * sell-in price
        };
      }
    }
  }
  
  const outputData = [];
  
  // Process constrained data
  for (let i = 1; i < constrainedData.length; i++) {
    const row = constrainedData[i];
    const helper = row[0];
    const customer = row[8];
    const sku = row[9];
    const pdt = row[6];
    const sellInPrice = parseFloat(row[2]) || 0;
    
    for (let j = weekStartIndex; j <= weekEndIndex; j++) {
      const week = constrainedData[0][j];
      if (week && week.toString().match(/202\d{3}/)) {
        const constrainedUnits = parseFloat(row[j]) || 0;
        const constrainedRevenue = constrainedUnits * sellInPrice;
        
        const key = `${helper}_${week}`;
        const unconstrained = unconstrainedLookup[key] || { units: 0, revenue: 0 };
        
        const deltaUnits = constrainedUnits - unconstrained.units;
        const deltaRevenue = constrainedRevenue - unconstrained.revenue;
        
        const quarter = getQuarter(parseInt(week));
        
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
          deltaUnits < 0 ? 'Supply Gap' : '',
          quarter === 'Q4 2025'
        ]);
      }
    }
  }
  
  // Write output data
  if (outputData.length > 0) {
    outputSheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
  }
  
  console.log(`Transformation complete! ${outputData.length} records created.`);
}

function getQuarter(week) {
  if (week >= 202534 && week <= 202539) return 'Q3 2025';
  if (week >= 202540 && week <= 202552) return 'Q4 2025';
  if (week >= 202553) return 'Q1 2026';
  return 'Unknown';
}
