function createOutputSpreadSheet(ssNames, propertyNames){
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('script_wk');
  const targetFolderId = inputSheet.getRange(9, 2).getValue();
  if (targetFolderId === ''){
    return;
  };
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  if (!targetFolder){
    return;
  };
  ssNames.forEach((x, idx) => {
    const sheet = SpreadsheetApp.create(x);
    const file = DriveApp.getFileById(sheet.getId());
    file.moveTo(targetFolder);
    PropertiesService.getScriptProperties().setProperty(propertyNames[idx], sheet.getId());
  });
}