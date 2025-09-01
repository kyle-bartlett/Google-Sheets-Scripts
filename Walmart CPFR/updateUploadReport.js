function processWeeklyCsvEmail() {
  // Configuration - UPDATE THESE VALUES
  const SHEET_ID = '1SYIFq0_AN9ziuPe4N59V1tDdHylhtbpSLUwVZMvyCsU'; // Get from your Google Sheet URL
  const SHEET_NAME = 'ReportUpload'; // Name of the tab where you want data pasted
  const START_CELL = 'M2'; // Where to start pasting (A1 = top-left corner)
  const EMAIL_SUBJECT_CONTAINS = 'FANTASIA TRADING LLC Vendor# 54205587'; // Part of subject line to search for
  const SENDER_EMAIL = 'rso-am-dp.groups@anker.com'; // Email address of sender (optional filter)
  
  try {
    // Search for unread emails with CSV attachments
    let searchQuery = `has:attachment filename:csv is:unread`;
    if (EMAIL_SUBJECT_CONTAINS) {
      searchQuery += ` subject:"${EMAIL_SUBJECT_CONTAINS}"`;
    }
    if (SENDER_EMAIL) {
      searchQuery += ` from:${SENDER_EMAIL}`;
    }
    
    const threads = GmailApp.search(searchQuery, 0, 10);
    
    if (threads.length === 0) {
      console.log('No new emails with CSV attachments found');
      return;
    }
    
    // Process the most recent email
    const messages = threads[0].getMessages();
    const latestMessage = messages[messages.length - 1];
    
    console.log(`Processing email: ${latestMessage.getSubject()}`);
    
    // Find CSV attachment
    const attachments = latestMessage.getAttachments();
    let csvAttachment = null;
    
    for (let attachment of attachments) {
      if (attachment.getName() === 'data_dump.csv') {
        csvAttachment = attachment;
        break;
      }
    }
    
    if (!csvAttachment) {
      console.log('No CSV attachment found in email');
      return;
    }
    
    // Parse CSV data
    const csvContent = csvAttachment.getDataAsString();
    const csvData = parseCSV(csvContent);
    
    if (csvData.length === 0) {
      console.log('CSV file appears to be empty');
      return;
    }
    
    // Open target Google Sheet
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`Sheet named "${SHEET_NAME}" not found`);
    }
    
    // Clear only the data area where we'll paste new CSV data
    // This calculates the range based on START_CELL and previous CSV dimensions
    const startRange = sheet.getRange(START_CELL);
    const startRow = startRange.getRow();
    const startCol = startRange.getColumn();
    
    // Determine how much area to clear (50 rows x 25 columns based on your data size)
    // Adjust these numbers if your CSV dimensions change
    const maxRowsToClear = 75; // A few extra rows for safety
    const maxColsToClear = 300; // A few extra columns for safety
    
    // Clear only the specific range where CSV data goes
    const clearRange = sheet.getRange(startRow, startCol, maxRowsToClear, maxColsToClear);
    clearRange.clear();
    
    // Paste new data
    const range = sheet.getRange(START_CELL);
    const targetRange = sheet.getRange(range.getRow(), range.getColumn(), csvData.length, csvData[0].length);
    targetRange.setValues(csvData);
    
    // Sort by column AH (LSTWKPOS) (Z to A - descending order)
    // CRITICAL: Sort only the DATA rows, NOT the header row
    const sortStartRow = startRow + 1; // Start from row 3 (skip header row 2)
    const sortStartCol = startCol; // Start from column M (our pasted data)
    const sortEndRow = startRow + csvData.length - 1; // Last row of data
    const sortEndCol = startCol + csvData[0].length - 1; // End of pasted data
    
    const sortRange = sheet.getRange(sortStartRow, sortStartCol, 
                                    sortEndRow - sortStartRow + 1, 
                                    sortEndCol - sortStartCol + 1);
    
    // Column LSTWKPOS is position 22 in our range (relative to column M)
    sortRange.sort({column: 22, ascending: false}); // Z to A = descending
    
    // Mark email as read
    latestMessage.markRead();
    
    console.log(`Successfully imported ${csvData.length} rows and ${csvData[0].length} columns`);
    console.log(`Data pasted starting at ${START_CELL} in sheet "${SHEET_NAME}"`);
    console.log('Data sorted by column AH (Z to A)');
    
  } catch (error) {
    console.error('Error processing email:', error.toString());
    
    // Optional: Send yourself an error notification
    // GmailApp.sendEmail('your@email.com', 'CSV Import Error', error.toString());
  }
}

function parseCSV(csvText) {
  const lines = csvText.split('\n');
  const result = [];
  
  for (let line of lines) {
    line = line.trim();
    if (line.length === 0) continue;
    
    // Simple CSV parsing (handles basic cases)
    const row = [];
    let currentField = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === ',' && !inQuotes) {
        row.push(currentField.trim());
        currentField = '';
      } else {
        currentField += char;
      }
    }
    
    // Add the last field
    row.push(currentField.trim());
    result.push(row);
  }
  
  return result;
}

function createTrigger() {
  // Delete existing triggers
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'processWeeklyCsvEmail')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new trigger
  const trigger = ScriptApp.newTrigger('processWeeklyCsvEmail')
    .timeBased()
    .everyHours(1)
    .create();
    
  console.log('Trigger ID:', trigger.getUniqueId());
  console.log('Handler function:', trigger.getHandlerFunction());
}