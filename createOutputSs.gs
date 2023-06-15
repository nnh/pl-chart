function createOutputSpreadSheet_(ssNames, propertyNames){
  const targetFolder = getOutputFolder_();
  if (targetFolder === null){
    return null;
  };
  const spreadSheets = ssNames.map((x, idx) => {
    const sheet = SpreadsheetApp.create(x);
    const file = DriveApp.getFileById(sheet.getId());
    file.moveTo(targetFolder);
    PropertiesService.getScriptProperties().setProperty(propertyNames[idx], sheet.getId());
    return sheet;
  });
  return spreadSheets.length === 1 ? spreadSheets[0] : spreadSheets;
}