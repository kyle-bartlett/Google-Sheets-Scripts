# 🚀 Quick Setup Checklist

## ✅ Before You Start
- [ ] You have a Google Sheet with your sales data
- [ ] You can access Google Apps Script (Extensions → Apps Script)
- [ ] You know the names of your source sheets
- [ ] You know which cell ranges contain the data you want to copy

## 🎯 Step-by-Step Setup

### 1. Open Apps Script
- [ ] Open your Google Sheet
- [ ] Go to **Extensions** → **Apps Script**
- [ ] Delete the default code in `Code.gs`

### 2. Copy the Automation Code
- [ ] Copy the entire contents of `Code.gs` from this folder
- [ ] Paste it into the Apps Script editor
- [ ] Click **Save** (Ctrl+S or Cmd+S)

### 3. Configure Your Data Sources
- [ ] Update `TARGET_SHEET_NAME` to match your target sheet
- [ ] Modify `SOURCE_SHEETS` array with your actual sheet names
- [ ] Update the `source` and `target` ranges for each data section
- [ ] Save the project again

### 4. Test Your Setup
- [ ] Return to your Google Sheet
- [ ] Refresh the page
- [ ] Look for **Costco PSI Automation** menu
- [ ] Click **Test Configuration**
- [ ] Fix any errors that appear

### 5. Run Your First Update
- [ ] Click **Costco PSI Automation** → **Run Weekly Update**
- [ ] Check that data appears in the correct locations
- [ ] Verify formatting is applied correctly

## 🔧 Configuration Examples

### Example 1: Simple Copy
```javascript
{
  name: 'Sales Sheet',
  ranges: [
    {
      source: 'A1:D10',      // Copy from A1 to D10
      target: 'B2:E11',      // Paste to B2 to E11
      description: 'Sales data'
    }
  ]
}
```

### Example 2: Multiple Ranges
```javascript
{
  name: 'Product Data',
  ranges: [
    {
      source: 'A1:C20',
      target: 'F2:H21',
      description: 'Product list'
    },
    {
      source: 'E1:G15',
      target: 'J2:L16',
      description: 'Performance metrics'
    }
  ]
}
```

## 🚨 Common Issues & Solutions

| Issue | Solution |
|-------|----------|
| Menu not appearing | Refresh the page, check script saved |
| "Sheet not found" | Verify sheet names match exactly |
| "Range not found" | Check that ranges exist in source sheets |
| Script errors | Use "Test Configuration" to debug |

## 📞 Need Help?
1. Check the execution logs in Apps Script
2. Verify your configuration matches your sheet structure
3. Start with a simple range (like A1:A5) to test
4. Ensure all source sheets exist and are accessible

## 🎉 You're Ready!
Once you complete this checklist, you'll have a fully automated weekly sales update system that will save you hours of manual copy-paste work!
