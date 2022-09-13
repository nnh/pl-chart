function createOutputSpreadSheet(ssNames, propertyNames){
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('script_wk');
  const targetFolderId = inputSheet.getRange(9, 2).getValue();
  if (targetFolderId === ''){
    return;
  };
//  const ssNames = ['PL（臨床研究センター以外）_1', 'PL（臨床研究センター以外）_2', 'PL（臨床研究センター以外）_3', 'PL（臨床研究センター以外）_4'];
//  const propertyNames = ['outputSpreadsheetIdOthers1', 'outputSpreadsheetIdOthers2', 'outputSpreadsheetIdOthers3', 'outputSpreadsheetIdOthers4'];
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
//createOutputSpreadSheet(['PL（臨床研究センター）'], ['outputSpreadsheetIdClinicalResearchCenter']);
//createOutputSpreadSheet(['PL（臨床研究センター以外）_1', 'PL（臨床研究センター以外）_2', 'PL（臨床研究センター以外）_3', 'PL（臨床研究センター以外）_4'],
//                        ['outputSpreadsheetIdOthers1', 'outputSpreadsheetIdOthers2', 'outputSpreadsheetIdOthers3', 'outputSpreadsheetIdOthers4']);