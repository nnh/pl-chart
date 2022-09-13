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
/**
* Delete the output sheets
* @param {spreadsheet} ss: The target spreadsheet
* @return {sheet} the leftmost sheet
*/
function deleteSheets(ss){
  const deleteSheetName = PropertiesService.getScriptProperties().getProperty('tempDeleteSheetName');
  // Keep the leftmost sheet and delete everything else
  var tempSheets = ss.getSheets();
  const tempSheet = tempSheets[0];
  if (tempSheets.length > 1){
    tempSheets.splice(0, 1);
    tempSheets.forEach(x => ss.deleteSheet(x));
  }
  tempSheet.setName(deleteSheetName);
  return tempSheet;
}
/**
* Create a sheet at the end
* @param {string} sheetName the sheet name
* @param {spreadsheet} ss: The target spreadsheet
* @return {sheet} the sheet objects
*/
function addSheetToEnd(sheetName, ss){
  const tempIdx = ss.getNumSheets();
  const targetSheet = ss.insertSheet(sheetName, parseInt(tempIdx)); 
  return targetSheet;  
}
function execCreateChart(values, target, processFlag){
  var chartConditions;
  if (processFlag){
    // by facility
    target.name = values[0];
    target.chartSheetName = values[1];
    chartConditions = new classSetChartConditionsByFacility(target);
  } else {
    // by years
    target.name = values;
    target.chartSheetName = values;
    chartConditions = new classSetChartConditionsByYear(target);
  }
  executeCreateChart(chartConditions);
}
