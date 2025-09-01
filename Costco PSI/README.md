# Costco PSI Weekly Sales Update Automation

This Google Apps Script automates your weekly sales update process by pulling data from multiple source sheets and pasting it into specific cells in a target summary sheet.

## 🚀 Features

- **Automated Data Transfer**: Copy data from multiple source sheets to a single target sheet
- **Flexible Range Mapping**: Define custom source and target ranges for each data section
- **Automatic Formatting**: Apply consistent formatting to the target sheet
- **Timestamp Tracking**: Add date stamps to track when updates were performed
- **Error Handling**: Comprehensive error checking and user feedback
- **Custom Menu**: Easy-to-use interface integrated into Google Sheets

## 📋 Prerequisites

- Google Sheets with your sales data
- Google Apps Script access
- Basic understanding of spreadsheet ranges (e.g., A1:D10)

## 🛠️ Setup Instructions

### Step 1: Open Your Google Sheet
1. Open the Google Sheet containing your sales data
2. Go to **Extensions** → **Apps Script**

### Step 2: Copy the Code
1. In the Apps Script editor, replace the default `Code.gs` content with the code from `Code.gs`
2. Save the project (Ctrl+S or Cmd+S)

### Step 3: Configure Your Data Sources
1. Modify the `CONFIG` object in the code to match your sheet structure
2. Update sheet names, ranges, and target locations
3. Save the project again

### Step 4: Test the Setup
1. Return to your Google Sheet
2. Refresh the page
3. You should see a new menu: **Costco PSI Automation**
4. Click **Test Configuration** to verify your setup

## ⚙️ Configuration Guide

### Basic Configuration Structure

```javascript
const CONFIG = {
  TARGET_SHEET_NAME: 'Weekly Summary',  // Your target sheet name
  
  SOURCE_SHEETS: [
    {
      name: 'Sheet Name',               // Source sheet name
      ranges: [
        {
          source: 'A1:D10',            // Source range to copy FROM
          target: 'B2:E11',            // Target range to paste TO
          description: 'What this data represents'
        }
      ]
    }
  ]
};
```

### Range Format Examples

- **Single Cell**: `'A1'`
- **Row Range**: `'A1:D1'`
- **Column Range**: `'A1:A10'`
- **Block Range**: `'A1:D10'`

### Common Use Cases

#### Weekly Sales Figures
```javascript
{
  name: 'Sales Data',
  ranges: [
    {
      source: 'B2:F8',        // Weekly sales data
      target: 'C3:G9',        // Pasted into summary
      description: 'Weekly sales by day'
    }
  ]
}
```

#### Product Performance
```javascript
{
  name: 'Product Metrics',
  ranges: [
    {
      source: 'A1:E20',       // Product performance data
      target: 'I2:M21',       // Pasted into summary
      description: 'Product sales metrics'
    }
  ]
}
```

## 🎯 How to Use

### Running the Weekly Update
1. Open your Google Sheet
2. Click **Costco PSI Automation** → **Run Weekly Update**
3. The script will automatically:
   - Copy data from all configured source sheets
   - Paste it into the target sheet at specified locations
   - Add a timestamp header
   - Apply formatting

### Testing Your Configuration
1. Click **Costco PSI Automation** → **Test Configuration**
2. The script will verify all source sheets exist
3. You'll get feedback on any missing sheets or configuration issues

### Viewing Current Settings
1. Click **Costco PSI Automation** → **View Configuration**
2. See your current configuration in a popup

## 🔧 Customization Options

### Date Format
```javascript
DATE_FORMAT: 'MM/dd/yyyy'  // Change to your preferred format
```

### Logging
```javascript
ENABLE_LOGGING: true       // Set to false to disable console logs
```

### Target Sheet Name
```javascript
TARGET_SHEET_NAME: 'Your Custom Sheet Name'
```

## 📊 Example Workflow

1. **Monday Morning**: Run the weekly update
2. **Data Transfer**: Script copies data from 3 source sheets
3. **Summary Creation**: All data appears in the "Weekly Summary" sheet
4. **Formatting Applied**: Borders, headers, and column sizing applied
5. **Timestamp Added**: Date and time of the update recorded

## 🚨 Troubleshooting

### Common Issues

#### "Sheet not found" Error
- Check that source sheet names exactly match your actual sheet names
- Ensure sheet names are spelled correctly (case-sensitive)

#### "Range not found" Error
- Verify that the specified ranges exist in your source sheets
- Check that ranges are formatted correctly (e.g., 'A1:D10')

#### Menu Not Appearing
- Refresh your Google Sheet page
- Check that the script saved successfully
- Ensure you're in the correct Google Sheet

### Debug Mode
- Open **View** → **Execution log** in Apps Script
- Run the script and check for error messages
- Use **Test Configuration** to validate your setup

## 📈 Advanced Features

### Conditional Data Copying
You can modify the script to only copy data that meets certain criteria:

```javascript
// Example: Only copy rows with sales > 0
const data = sourceRange.getValues().filter(row => row[2] > 0);
```

### Data Transformation
Add calculations or formatting during the copy process:

```javascript
// Example: Calculate totals during copy
const data = sourceRange.getValues();
const totals = data.map(row => [...row, row.reduce((a, b) => a + b, 0)]);
```

### Multiple Target Sheets
Extend the script to copy data to multiple summary sheets:

```javascript
const TARGET_SHEETS = ['Weekly Summary', 'Monthly Summary', 'Executive Summary'];
```

## 🔒 Security Notes

- The script only accesses the Google Sheet it's attached to
- No data is sent to external servers
- All processing happens within Google's secure environment
- You can review all code before running

## 📞 Support

If you encounter issues:
1. Check the execution logs in Apps Script
2. Verify your configuration matches your sheet structure
3. Test with a small range first
4. Ensure all source sheets exist and are accessible

## 🎉 Success Tips

- Start with a simple configuration and build up
- Test with small data ranges first
- Use descriptive names for your ranges
- Keep a backup of your original data
- Document any custom modifications you make

---

**Happy Automating! 🚀**

This tool should save you significant time on your weekly sales updates while ensuring consistency and reducing manual errors.
