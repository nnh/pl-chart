function createOutputSpreadSheet(ssNames, propertyNames){
  const targetFolder = getOutputFolder_();
  if (targetFolder === null){
    return null;
  };
  ssNames.forEach((x, idx) => {
    const sheet = SpreadsheetApp.create(x);
    const file = DriveApp.getFileById(sheet.getId());
    file.moveTo(targetFolder);
    PropertiesService.getScriptProperties().setProperty(propertyNames[idx], sheet.getId());
  });
}