function createOutputSpreadSheet_(propertyNameAndSpreadsheetNameList){
  const targetFolder = getOutputFolder_();
  if (targetFolder === null){
    return null;
  };
  const spreadSheets = propertyNameAndSpreadsheetNameList.map(([propertyName, spreadsheetName]) => {
    const spreadsheet = SpreadsheetApp.create(spreadsheetName);
    const file = DriveApp.getFileById(spreadsheet.getId());
    file.moveTo(targetFolder);
    PropertiesService.getScriptProperties().setProperty(propertyName, spreadsheet.getId());
    return spreadsheet;
  });
  return spreadSheets.length === 1 ? spreadSheets[0] : spreadSheets;
}