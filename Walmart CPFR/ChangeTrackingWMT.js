/**
 * This script tracks changes made to specific tabs within the current Google Sheet
 * and logs them to a designated Master Change Log sheet.
 * It now handles both single-cell and multi-cell edits.
 *
 * IMPORTANT:
 * 1. Replace 'YOUR_MASTER_CHANGE_LOG_SHEET_ID_HERE' with the actual ID of your "Anker Master Change Log" sheet.
 * 2. Customize the 'tabsToMonitor' array to include only the exact names of the tabs you want to track in THIS spreadsheet.
 * 3. This script should be deployed in EACH of the 6 critical sheets you want to monitor.
 */

function onEdit(e) {
  const masterLogSheetId = '1DAz-UqHZIcx26rnw8GEcpeLv5ceY5a8K91YCLGIfEcA'; // <<< REPLACE THIS with your Master Log Sheet ID
  const masterLogSheetName = 'Change_Log_Raw_Data'; // Default tab name in your new Master Log sheet. If you rename it, update this.

  // Define the exact names of the tabs to monitor in THIS specific spreadsheet
  // <<< CUSTOMIZE THIS ARRAY FOR EACH OF YOUR 6 SHEETS >>>
  // Example for WMT Charging CPFR 25WMWK10 ANKERWK22:
  // const tabsToMonitor = ['CPFR', 'CPFR SKUs'];
  // Example for WM Mod reset NPI/IOQ & Tracker (all tabs):
  // const tabsToMonitor = []; // An empty array or removing the if condition below will make it track all tabs
  const tabsToMonitor = []; // Current example from your WMT sheet. ADJUST FOR EACH SCRIPT.

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // Only proceed if the edited tab is one of the tabs we want to monitor
  // This check is skipped if tabsToMonitor is empty (meaning all tabs are monitored)
  if (tabsToMonitor.length > 0 && !tabsToMonitor.includes(sheetName)) {
    return; // Exit if the edited tab is not in our monitored list
  }

  const userEmail = e.user ? e.user.getEmail() : 'Unknown';
  const timestamp = new Date();
  const cellAddress = range.getA1Notation(); // This will be the address of the edited cell or the range (e.g., A1 or A1:B10)

  let oldValueLog;
  let newValueLog;

  // Check if this is a multi-cell edit (paste, clear range, delete rows/cols)
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
    // Multi-cell edit detected
    oldValueLog = "Multiple Cells Changed (Old: N/A)"; // Cannot reliably get all old values for a range

    // Capture new values of the range. e.values is a 2D array if a paste occurred.
    // If cells were cleared/deleted, e.values might be undefined or an array of empty strings/nulls.
    const newValues = e.values;

    if (newValues) {
      // If e.values exists, convert the 2D array to a string for logging
      // Example: [[1,2],[3,4]] -> "[[1,2],[3,4]]"
      // If it's a delete/clear, newValues might be an array of empty strings
      const isCompletelyCleared = newValues.every(row => row.every(cell => cell === null || cell === ''));
      if (isCompletelyCleared) {
        newValueLog = "Range Cleared/Deleted";
      } else {
        newValueLog = JSON.stringify(newValues);
      }
    } else {
      // Fallback for cases where e.values might not be available for multi-cell (e.g., specific delete operations)
      newValueLog = "Bulk Operation (New values not directly captured)";
    }

  } else {
    // Single-cell edit (original behavior)
    oldValueLog = e.oldValue !== undefined ? e.oldValue : ''; // Capture old value, handle undefined for new cells
    newValueLog = e.value !== undefined ? e.value : '';     // Capture new value, handle undefined for cleared cells

    // If old value and new value are the same, don't log (e.g., re-entering same data)
    if (oldValueLog === newValueLog) {
      return;
    }
  }

  // Get the Master Change Log sheet
  let masterLogSheet;
  try {
    const masterSpreadsheet = SpreadsheetApp.openById(masterLogSheetId);
    masterLogSheet = masterSpreadsheet.getSheetByName(masterLogSheetName);
    if (!masterLogSheet) {
      throw new Error('Master log sheet not found: ' + masterLogSheetName);
    }
  } catch (error) {
    Logger.log('Error opening master log sheet: ' + error.message);
    // You might want to send an email notification here if the log sheet can't be accessed
    return;
  }

  // Append the change data to the Master Change Log
  masterLogSheet.appendRow([
    timestamp,
    userEmail,
    SpreadsheetApp.getActiveSpreadsheet().getName(), // Name of the *source* sheet
    sheetName, // Name of the tab within the source sheet
    cellAddress, // Will be A1 notation of cell or range (e.g., A1:B10)
    oldValueLog,
    newValueLog
  ]);
}