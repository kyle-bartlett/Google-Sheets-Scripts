# ANKER FORECAST AUTOMATION - FINAL SYSTEM

## Step 1: Upload Excel File to Google Drive
1. Upload "Charging Team Q3Q4 Constrained vs Unconstrained Forecast Data 1.xlsx" to Google Drive
2. Open in Google Sheets
3. This will create separate sheets for each tab

## Step 2: Transform Wide to Tall Data (Google Sheets Method)

### Create "Looker_Ready_View" Sheet

**Method A: Use Apps Script (Recommended - 5 minutes)**
1. In Google Sheets, go to Extensions → Apps Script
2. Paste this code:

```javascript
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
```

3. Save and run `transformForecastData()`
4. This creates your Looker_Ready_View sheet automatically

**Method B: Manual Formula Method (15 minutes)**
If Apps Script doesn't work, use this formula approach:
1. Create new sheet "Looker_Ready_View"
2. Set up headers in row 1
3. Use complex formulas to unpivot the data (more time consuming)

## Step 3: Create Dashboard Sheets

### Summary_SKU Sheet
1. Create pivot table from Looker_Ready_View
2. Filters: Gap Flag = "Supply Gap"
3. Rows: Anker SKU
4. Values: Sum of Delta Units, Sum of Delta Revenue
5. Sort by Delta Revenue descending

### Summary_Customer Sheet  
1. Create pivot table from Looker_Ready_View
2. Filters: Gap Flag = "Supply Gap" 
3. Rows: Customer
4. Values: Sum of Delta Units, Sum of Delta Revenue
5. Sort by Delta Revenue descending

### Weekly_Trends Sheet
1. Create pivot table from Looker_Ready_View
2. Filters: Gap Flag = "Supply Gap"
3. Rows: Week
4. Values: Sum of Delta Units, Sum of Delta Revenue
5. Sort by Week ascending

## Step 4: Create Charts
1. **SKU Impact Chart**: Column chart from Summary_SKU (top 10)
2. **Customer Impact Chart**: Column chart from Summary_Customer (top 10)  
3. **Weekly Trend Chart**: Line chart from Weekly_Trends
4. **Customer-Week Heatmap**: Pivot table with Customer (rows) × Week (columns)

## Step 5: Dashboard Layout
### Dashboard Sheet Structure:
- **Rows 1-8**: Key metrics and title
- **Rows 10-25**: Embedded charts
- **Rows 27+**: Links to detailed pivot tables

### Key Metrics Formulas:
```
Total Gaps: =COUNTIF(Looker_Ready_View!K:K,"Supply Gap")
Revenue at Risk: =SUMIF(Looker_Ready_View!K:K,"Supply Gap",Looker_Ready_View!J:J)*-1
Q3 Impact: =SUMIFS(Looker_Ready_View!I:I,Looker_Ready_View!K:K,"Supply Gap",Looker_Ready_View!E:E,"Q3 2025")*-1
Q4 Impact: =SUMIFS(Looker_Ready_View!I:I,Looker_Ready_View!K:K,"Supply Gap",Looker_Ready_View!E:E,"Q4 2025")*-1
```

## Step 6: Conditional Formatting
1. Select revenue impact columns
2. Format → Conditional formatting
3. **Red**: Values > $50,000
4. **Orange**: Values $20,000-$50,000  
5. **Yellow**: Values $5,000-$20,000

## Step 7: Final Polish
1. Bold headers and key metrics
2. Add borders around sections
3. Format numbers (currency, thousands separators)
4. Set up data validation dropdowns for filtering
5. Protect formulas but allow data entry where needed

## Time Estimates:
- **Apps Script method**: 30 minutes total
- **Manual method**: 2 hours total
- **Charts and formatting**: 45 minutes
- **Dashboard polish**: 30 minutes

## For Next Time - Full Automation:
This creates a reusable template. For future updates:
1. Upload new Excel file
2. Run the Apps Script transformation
3. All pivot tables and charts update automatically
4. Dashboard refreshes with new data

## Troubleshooting:
- If Apps Script fails: Check sheet names match exactly
- If pivot tables break: Refresh data source ranges
- If formulas error: Check column references match your data structure

## CURSOR AI INTEGRATION:
To build this automatically with Cursor, provide this specification:
1. Read Excel file with two sheets (Constrained/Unconstrained forecasts)
2. Transform wide format to tall format (unpivot week columns)
3. Calculate deltas (Constrained - Unconstrained)
4. Create Google Sheets workbook with pivot tables and charts
5. Export as Google Sheets format or generate Apps Script code

The key insight: Your Excel has 20 weeks of data per SKU-Customer combination. The transformation creates 20 rows per combination, making it analyzable in Google Sheets pivot tables.