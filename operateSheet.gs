// 印刷後に実行する
/**
* Show the sheets not to be printed
* @param none
* @return none
*/
function visibleWorkings(){
  const visibleSheets = new classVisibleSheets;
  visibleSheets.exec();
}
/**
* Hide the sheets not to be printed
* @param none
* @return none
*/
function hiddenWorkings(){
  const hiddenSheets = new classHiddenSheets;
  hiddenSheets.exec();
}
class classVisibleHiddenSheets{
  constructor(){
  this.targetSheets = getSheetsNonTargetPrinting();
  }
}
class classVisibleSheets extends classVisibleHiddenSheets{
  exec(){
    this.targetSheets.map(x => x.showSheet());
  }
}
class classHiddenSheets extends classVisibleHiddenSheets{
  exec(){
    this.targetSheets.map(x => x.hideSheet());
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
    tempSheets.map(x => ss.deleteSheet(x));
  }
  tempSheet.setName(deleteSheetName);
  return tempSheet;
}
/**
* Delete the sheet
* @param {string} sheetName the sheet name
* @return none
*/
function deleteTargetSheet(sheetName){
  const targetSheet = getTargetSheet(PropertiesService.getScriptProperties().getProperty(sheetName));
  if (targetSheet != null){
    ss.deleteSheet(targetSheet);
  }
}
/**
* Sort the sheets
* Order: facility name, year
* @param {spreadsheet} ss: The target spreadsheet
* @return none
*/
function sortSheets(ss){
  const targetFacilities = getTargetFacilitiesNames();
  const targetYears = getTargetYears();
  const targetSheetsName = targetFacilities.concat(targetYears);
  targetSheetsName.map(function(x, idx){
    x.activate();
    ss.moveActiveSheet(idx + 1);
  });
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
