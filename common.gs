/**
* Setting Project Properties
* @param none
* @return none
*/
function registerScriptProperty(){
  PropertiesService.getScriptProperties().setProperty('facilitySheetName', 'site');
  PropertiesService.getScriptProperties().setProperty('inputSheetName', 'working');
  PropertiesService.getScriptProperties().setProperty('scriptWorkingSheetName', 'script_wk');
  PropertiesService.getScriptProperties().setProperty('siteSheetfacilityCodeCol', 'A');
  PropertiesService.getScriptProperties().setProperty('siteSheetfacilityNameCol', 'B');
  PropertiesService.getScriptProperties().setProperty('inputSheetfacilityNameCol', 'B');
  PropertiesService.getScriptProperties().setProperty('inputSheetyearsCol', 'C');
  PropertiesService.getScriptProperties().setProperty('inputSheetfacilityCodeCol', 'K');
  PropertiesService.getScriptProperties().setProperty('clinicalResearchCenter', '臨床研究センター');
  const scriptWorkingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PropertiesService.getScriptProperties().getProperty('scriptWorkingSheetName'));;
  PropertiesService.getScriptProperties().setProperty('outputSpreadsheetIdClinicalResearchCenter', scriptWorkingSheet.getRange(1, 2).getValue());
  PropertiesService.getScriptProperties().setProperty('outputSpreadsheetIdOthers1', '');
  PropertiesService.getScriptProperties().setProperty('outputSpreadsheetIdOthers4', '');
  PropertiesService.getScriptProperties().setProperty('tempDeleteSheetName', 'temp_del');
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
class classGetColumnInfo{
  constructor(columnName){
    this.columnNumber = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(columnName + '1').getColumn();
  }
  get arrayIndex(){
    return this.getArrayIndex();
  }
  getArrayIndex(){
    var temp = this.columnNumber;
    temp--;
    return temp;
  }
  get queryColumnName(){
    return this.createQueryColumnName();
  }
  createQueryColumnName(){
    return 'Col' + this.columnNumber;
  }  
}
