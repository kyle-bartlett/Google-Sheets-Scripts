// Make sure to update folderId, importrangesheetID, and include/exclude sheetname. 
// Run the GetimportrangeID first to get importrangesheetID and include sheetname.

var SCRIPT_PROP = PropertiesService.getScriptProperties();

// Function to initiate the file copying and processing.
function copyAndProcessFile() {
  Logger.log('Initiating copyAndProcessFile function.');

  // Obtain the current spreadsheet file.
  const sourceFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  Logger.log('Obtained current spreadsheet file.');

  // Define the destination folder ID.
  const folderId = '1hQHqQZGsejwINUlDTb4JICN3XdIES07B';
  const folder = DriveApp.getFolderById(folderId);
  Logger.log('Folder Name: ' + folder.getName());

  // Define the new file name.
  const newFileName = sourceFile.getName() + ' ' + Utilities.formatDate(new Date(), 'GMT', 'yyyy.MM.dd');
  Logger.log('Defined new file name.');
  Logger.log('Source File ID: ' + sourceFile.getId());
  Logger.log('Source File Name: ' + sourceFile.getName());
  Logger.log('New File Name: ' + newFileName);

  // Make a copy of the file with retry logic
  let newFile = tryCopyFile(sourceFile, newFileName, folder, 5);
  
//  if (!newFile) {
//    Logger.log('Terminating due to repeated errors.');
//    return; // Exit the function if there's an error during the copy
// }

  // Grant permissions for all importrange formulas.
  grantImportRangePermissions(newFile.getId());
  Logger.log('Granted permissions for all importrange formulas.');

  // Wait for 2 minutes.
  Utilities.sleep(2 * 60 * 1000);
  Logger.log('Waited for 2 minutes.');

  // Set the spreadsheetId and nextSheetIndex in script properties.
  SCRIPT_PROP.setProperty('spreadsheetId', newFile.getId());
  SCRIPT_PROP.setProperty('nextSheetIndex', 0);
  Logger.log('Set spreadsheetId and nextSheetIndex in script properties.');

  // Replace importrange formulas with values.
  triggerReplaceImportRangeWithValues();
  Logger.log('Initiated replacement of importrange formulas with values.');
}

// Function to grant permissions for all importrange formulas.
function grantImportRangePermissions(spreadsheetId) {
  // Define the IDs of the spreadsheets that are importing ranges.
  const importRangeSheetIds = [
    "1KXmfnX5dUfDfRwQUhDhvfngOb52o35a4Fnk9zoR78xM",
    "1265dP-LCedwpxZzluG-KlCwd65lMPDfpBJD2mAc2Nq4",
    "1Rljpb6nj5Wxs9pzrTgH681z1Q8XWlMfzJHxYK9OUISM",
    "1NuWNoSa1XcdjrHq9kL1npszvViiMbHV1_0tVm_zTjSM",
    "1Bsz3Hlrn9d6RDO8aksCk4DZDhQ8QjE7rcDbSTsuTCOI",
    "11iZYly0LkpllmOyUL-5zfwghQMZ4BVW_Xbrj6KvOeW0",
    "11VOdGH66QwP33g7jugReb2-ZBxs3YJU6sKYqauwsphs",
    "1Fe6EU1s-uMFosXI2NHmN4pY-WvYO4skLXu5T-yxwC0Q"
  ];

  Logger.log('Starting to grant permissions for importrange formulas.');

  // Loop through the IDs and grant permissions for each.
  importRangeSheetIds.forEach(id => {
    try {
      Logger.log(`Granting permission for Sheet ID: ${id}`);
      addImportrangePermission(spreadsheetId, id);
      Logger.log(`Successfully granted permission for Sheet ID: ${id}`);
    } catch (e) {
      Logger.log(`Error while granting permission for Sheet ID: ${id} - ${e.toString()}`);
    }
  });

  Logger.log('Completed granting permissions for importrange formulas.');
}

// Function to grant importrange permission for a specific spreadsheet.
function addImportrangePermission(spreadsheetId, donorId) {
  Logger.log(`Initiating addImportrangePermission function for donorId: ${donorId}.`);

  // Define the URL for adding importrange permissions.
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/externaldata/addimportrangepermissions?donorDocId=${donorId}`;
  Logger.log('Defined URL for adding importrange permissions.');

  // Get the OAuth token.
  const token = ScriptApp.getOAuthToken();
  Logger.log('Obtained OAuth token.');

  // Define the parameters for the request.
  const params = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true
  };
  Logger.log('Defined parameters for the request.');

  // Send the request and log the response.
  var response = UrlFetchApp.fetch(url, params);
  Logger.log(`Response for donorId ${donorId}: ${response.getContentText()}`);
}


function tryCopyFile(sourceFile, newFileName, folder, maxAttempts) {
  let attempts = 0;
  while (attempts < maxAttempts) {
    try {
      return sourceFile.makeCopy(newFileName, folder);
    } catch (e) {
      if (attempts === maxAttempts - 1) { // last attempt
        Logger.log('Failed after ' + maxAttempts + ' attempts. Error: ' + e.toString());
        return null;
      }
      Logger.log('Error while copying, retrying... (' + (attempts+1) + '/' + maxAttempts + ')');
      Utilities.sleep(5000);  // Wait for 5 seconds before retrying
      attempts++;
    }
  }
}

// Function to replace importrange formulas with values in a spreadsheet.
function replaceImportRangeWithValues(spreadsheetId, sheetIndex = 0) {
  const batchSize = 2;

  // Sheets to be specifically processed. If empty, all sheets are processed except for those in `excludedSheets`.
  const includedSheets = ["Sellout History", "Daily Inv", "Processed PO", "Instock", "Sell In Price", "Mapping"]; 
  // Sheets to be excluded from processing.
  const excludedSheets = [];

  const maxRetries = 3;
  const startTime = new Date().getTime();
  const timeLimit = 20 * 60 * 1000;

  // Retry loop: will retry the whole batch if an error occurs, up to maxRetries times
  for(let retries = 0; retries < maxRetries; retries++) {
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheets = spreadsheet.getSheets();

      if (sheetIndex >= sheets.length) {
        Logger.log('All sheets processed.');
        deleteTriggers();
        return;
      }

      Logger.log('Processing batch starting from sheet index: ' + sheetIndex);

      for(let i = sheetIndex; i < sheetIndex + batchSize && i < sheets.length; i++){
        let sheet = sheets[i];

        // Check if the sheet should be processed based on inclusion/exclusion lists
        if (includedSheets.length > 0) {
          if (!includedSheets.includes(sheet.getName())) {
            continue; // Skip the sheet if it's not in the `includedSheets` list
          }
        } else if (excludedSheets.includes(sheet.getName())) {
          continue; // Skip the sheet if it's in the `excludedSheets` list
        }

        Logger.log('Attempting to process sheet: ' + sheet.getName());
        const range = sheet.getDataRange();
        const values = range.getValues();
        range.setValues(values);
        Logger.log('Sheet processed: ' + sheet.getName());


        const elapsedTime = new Date().getTime() - startTime;
        if (elapsedTime > timeLimit) {
          Logger.log('Time limit exceeded. Setting up trigger for next batch.');
          
          // Set up the trigger for the next batch
          SCRIPT_PROP.setProperty('nextSheetIndex', i);
          ScriptApp.newTrigger('triggerReplaceImportRangeWithValues')
            .timeBased()
            .after(5000)
            .create();
          Logger.log('Trigger set up for sheet index: ' + i);
          
          // Check and delete triggers if necessary
          const allTriggers = ScriptApp.getProjectTriggers();
          const targetTriggers = allTriggers.filter(trigger => trigger.getHandlerFunction() === 'triggerReplaceImportRangeWithValues');
          if (targetTriggers.length === 10) {
            for (let j = 0; j < 3; j++) {
              ScriptApp.deleteTrigger(targetTriggers[j]);
            }
            Logger.log('First 3 triggers deleted.');
          }

          return;
        }
      }

      Logger.log('Batch processed. Setting up trigger for next batch.');

      SpreadsheetApp.flush();

      SCRIPT_PROP.setProperty('nextSheetIndex', sheetIndex + batchSize);

      ScriptApp.newTrigger('triggerReplaceImportRangeWithValues')
        .timeBased()
        .after(5000)
        .create();

      Logger.log('Trigger set up.');

      break;
    } catch (error) {
      Logger.log(`An error occurred while accessing the spreadsheet (attempt ${retries + 1}): ${error.toString()}`);
      if (retries === maxRetries - 1) {
        Logger.log('Max retries reached. Stopping the function.');
        deleteTriggers();
        return;
      }

      Logger.log('Retrying...');
    }
  }
}


// Function to delete only the "triggerReplaceImportRangeWithValues" project triggers.
function deleteTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === 'triggerReplaceImportRangeWithValues') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Deleted triggerReplaceImportRangeWithValues trigger.');
    }
  }
}


// Function to trigger the replacement of importrange formulas with values.
function triggerReplaceImportRangeWithValues() {
  const spreadsheetId = SCRIPT_PROP.getProperty('spreadsheetId');
  const nextSheetIndex = Number(SCRIPT_PROP.getProperty('nextSheetIndex'));
  replaceImportRangeWithValues(spreadsheetId, nextSheetIndex);
}

// Run the function
function run() {
  copyAndProcessFile();
}
