# Google Apps Script: Two-Way Live Editing Between Tabs

This script enables real-time synchronization between two tabs in your Google Sheets report, allowing you to edit data in either tab and have it automatically update in the other tab.

## Overview

- **Report ID**: `1I-O6PF3LyK0CQ3SrU2ygmd1uxOVNE920v_TpmLUvG6k`
- **Tab 1**: "All SKU Rollup WoW - FY'2025" (Column T for editing, Column W for helper lookup)
- **Tab 2**: "TOP SKU Level Summary" (Column T for editing, Column A for helper lookup)

## How It Works

1. **Helper Columns**: Each tab has a helper column that contains unique identifiers
   - "All SKU Rollup" tab: Column W (W7:W) contains helper values
   - "TOP SKU Level" tab: Column A (A7:A) contains helper values

2. **Editing Columns**: Both tabs use Column T (T7:T) for data entry

3. **Synchronization**: When you edit a cell in Column T on either tab, the script:
   - Reads the helper value from the same row
   - Finds the matching helper value in the other tab
   - Updates the corresponding Column T cell in the other tab

## Setup Instructions

### Step 1: Open Google Apps Script
1. Go to [script.google.com](https://script.google.com)
2. Click "New Project"
3. Delete the default code and paste the contents of `Code.gs`

### Step 2: Configure the Script
1. Update the `SPREADSHEET_ID` if needed (currently set to your report ID)
2. Verify the tab names match exactly:
   - `ALL_SKU_TAB_NAME = "All SKU Rollup WoW - FY'2025"`
   - `TOP_SKU_TAB_NAME = "TOP SKU Level Summary"`

### Step 3: Save and Authorize
1. Click "Save" (Ctrl+S or Cmd+S)
2. Give your project a name (e.g., "Two-Way Editing Script")
3. Click "Authorize" when prompted
4. Grant the necessary permissions to access your Google Sheets

### Step 4: Run the Setup
1. Select the `setupTwoWayEditing` function from the dropdown
2. Click the "Run" button (▶️)
3. Check the execution log for success messages

### Step 5: Verify Installation
1. Go back to your Google Sheet
2. You should see a new menu item: "Two-Way Editing"
3. Use "Test Connection" to verify both tabs are accessible

## Usage

### Automatic Synchronization
Once set up, the script automatically syncs changes:
- Edit any cell in Column T (starting from row 7) on either tab
- The change will automatically appear in the corresponding row on the other tab
- Synchronization happens in real-time as you type

### Manual Functions
Use the "Two-Way Editing" menu for additional options:

- **Setup System**: Re-run the initial setup
- **Test Connection**: Verify both tabs are accessible
- **Manual Sync**: Force a complete synchronization between tabs
- **Cleanup**: Remove all triggers (use if you want to disable the system)

## Troubleshooting

### Common Issues

1. **"Script not running"**
   - Check that triggers are properly set up
   - Run `setupTwoWayEditing` again
   - Verify the script has edit permissions on your sheet

2. **"Changes not syncing"**
   - Ensure helper columns (W and A) contain matching values
   - Check that you're editing in Column T starting from row 7
   - Verify tab names match exactly (including spaces and special characters)

3. **"Permission denied"**
   - Re-authorize the script
   - Ensure you're the owner or have edit access to the sheet

### Debug Mode
- Check the execution logs in Google Apps Script
- Use `Logger.log()` statements to track script execution
- Run individual functions to test specific functionality

## Data Structure Requirements

### Helper Columns Must Contain:
- **Unique identifiers** that match between both tabs
- **No empty cells** in rows with data (starting from row 7)
- **Consistent formatting** (avoid mixed data types)

### Example Helper Values:
```
Tab 1 (Column W): SKU001, SKU002, SKU003
Tab 2 (Column A): SKU001, SKU002, SKU003
```

## Performance Notes

- The script is optimized for your data size (~100 rows in TOP SKU, ~1000 rows in ALL SKU)
- Changes sync in real-time with minimal delay
- The script includes safeguards to prevent infinite loops
- Manual sync can handle bulk updates efficiently

## Security

- The script only accesses the specified spreadsheet
- No data is sent to external servers
- All operations are logged for audit purposes
- You can revoke permissions at any time

## Support

If you encounter issues:
1. Check the execution logs in Google Apps Script
2. Verify all configuration settings match your sheet
3. Test with a simple edit to ensure basic functionality works
4. Use the "Test Connection" function to verify access

## Customization

To modify the script for different columns or tabs:
1. Update the column constants at the top of the script
2. Modify the tab names if needed
3. Adjust the starting row number (currently set to 7)
4. Test thoroughly before deploying to production
