/**
* Setting Project Properties
* @param none
* @return none
*/
function registerScriptProperty(){
  PropertiesService.getScriptProperties().setProperty('facilitySheetName', 'site');
  PropertiesService.getScriptProperties().setProperty('inputSheetName', 'working');
  PropertiesService.getScriptProperties().setProperty('siteSheetfacilityCodeCol', 'A');
  PropertiesService.getScriptProperties().setProperty('siteSheetfacilityNameCol', 'B');
  PropertiesService.getScriptProperties().setProperty('inputSheetfacilityNameCol', 'B');
  PropertiesService.getScriptProperties().setProperty('inputSheetyearsCol', 'C');
  PropertiesService.getScriptProperties().setProperty('inputSheetfacilityCodeCol', 'K');
  PropertiesService.getScriptProperties().setProperty('clinicalResearchCenter', '臨床研究センター');
  PropertiesService.getScriptProperties().setProperty('outputSpreadsheetIdOthers1', '');
  PropertiesService.getScriptProperties().setProperty('outputSpreadsheetIdOthers4', '');
}
/**
* Returns an array index from a column name
* @param {string} columnName Column name (e.g. 'A')
* @return {number} The corresponding array index of getValues (e.g. return 0 if 'A')
*/
function getColumnNumber(columnName){ 
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  var colNumber = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(columnName + '1').getColumn();
  colNumber--;
  return colNumber;
}
/**
* Get the sheet
* @param {string} sheetName the sheet name
* @param {spreadsheet} ss: The target spreadsheet
* @return {sheet} the sheet objects
*/
function getTargetSheet(sheetName, ss=SpreadsheetApp.getActiveSpreadsheet()){
  return ss.getSheetByName(sheetName);
}
