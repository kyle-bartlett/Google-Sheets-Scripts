function snapshotSelloutHistory() {
  const sourceFileId = '1KXmfnX5dUfDfRwQUhDhvfngOb52o35a4Fnk9zoR78xM';
  const sourceSheetName = 'WMT';
  const sourceRange = 'A1:NZ1000'; // adjust if needed

  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sellout History');
  if (!targetSheet) {
    Logger.log('Target sheet "Sellout History" not found.');
    return;
  }

  const sourceSheet = SpreadsheetApp.openById(sourceFileId).getSheetByName(sourceSheetName);
  const values = sourceSheet.getRange(sourceRange).getValues();

  targetSheet.getRange('A5').offset(0, 0, values.length, values[0].length).setValues(values);

  const now = new Date();
  const formatted = Utilities.formatDate(now, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  targetSheet.getRange('B2').setValue(`Last Snapshot: ${formatted}`);
}