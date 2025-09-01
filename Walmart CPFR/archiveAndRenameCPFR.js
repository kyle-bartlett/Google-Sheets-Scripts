function archiveAndRenameCPFR() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const currentName = file.getName();

  // Step 1: Backup the file
  DriveApp.getFileById(file.getId()).makeCopy(`Copy of ${currentName}`);

  // Step 2: Match both WMWK## and ANKERWK##
  const wmMatch = currentName.match(/WMWK(\d{2})/);
  const ankerMatch = currentName.match(/ANKERWK(\d{2})/);

  if (!wmMatch || !ankerMatch) {
    Logger.log("Filename must contain both WMWK## and ANKERWK##.");
    return;
  }

  // Increment each week number
  let nextWMWK = parseInt(wmMatch[1], 10) + 1;
  let nextANKERWK = parseInt(ankerMatch[1], 10) + 1;

  // Zero-pad if needed
  if (nextWMWK < 10) nextWMWK = "0" + nextWMWK;
  if (nextANKERWK < 10) nextANKERWK = "0" + nextANKERWK;

  // Step 3: Replace both in filename
  const newName = currentName
    .replace(/WMWK\d{2}/, `WMWK${nextWMWK}`)
    .replace(/ANKERWK\d{2}/, `ANKERWK${nextANKERWK}`);

  // Step 4: Rename original file
  file.setName(newName);
  Logger.log(`File renamed to: ${newName}`);
}