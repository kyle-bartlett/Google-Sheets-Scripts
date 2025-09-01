function updateDashboard() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("Dashboard");
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "Dashboard" not found');
    return;
  }
  
  // Copy C3:I9 to C11:I17 as values only
  const sourceRange = sheet.getRange("C3:I9");
  const destRange = sheet.getRange("C11:I17");
  const values = sourceRange.getValues();
  destRange.setValues(values);
  
  // Set E1 to today's date
  const today = new Date();
  sheet.getRange("E1").setValue(today);
  
  SpreadsheetApp.getUi().alert('Dashboard updated successfully');
}

// Combined menu for all functions
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Scripts')
    .addItem('Update Dashboard', 'updateDashboard')
    .addItem('Duplicate Columns I:AH', 'duplicateColumns')
    .addSeparator()
    .addSubMenu(ui.createMenu('Run All Updates')
      .addItem('Run All Weekly Updates', 'runAllWeeklyUpdates'))
    .addToUi();
}

// Master function to run all updates
function runAllWeeklyUpdates() {
  try {
    // Add your existing function names here
    updateDashboard();
    duplicateColumns(); // if this applies to weekly updates
    
    // Add calls to your existing functions like:
    // existingFunction1();
    // existingFunction2();
    // etc.
    
    SpreadsheetApp.getUi().alert('All weekly updates completed successfully');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error during updates: ' + error.toString());
  }
}

// Set up weekly schedule (run this once to establish the trigger)
function setupWeeklyTrigger() {
  // Delete existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runAllWeeklyUpdates') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new weekly trigger (runs every Monday at 9 AM)
  ScriptApp.newTrigger('runAllWeeklyUpdates')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
    
  SpreadsheetApp.getUi().alert('Weekly trigger set for Mondays at 9 AM');
}
