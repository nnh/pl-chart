/**
* Setting Project Properties
* @param none
* @return none
*/
function registerScriptProperty(){
  PropertiesService.getScriptProperties().setProperty('facilitySheetName', 'site');
  PropertiesService.getScriptProperties().setProperty('inputSheetName', 'working');
  PropertiesService.getScriptProperties().setProperty('inputSheetfacilityNameCol', 'B');
  PropertiesService.getScriptProperties().setProperty('inputSheetyearsCol', 'C');
  PropertiesService.getScriptProperties().setProperty('clinicalResearchCenter', '臨床研究センター');
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
* @return {sheet} the sheet objects
*/
function getTargetSheet(sheetName){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}
