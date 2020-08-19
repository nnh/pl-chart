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
* Get the sheets that don't print
* @param none
* @return {sheet} the sheet objects that don't print
*/
function getSheetsNonTargetPrinting(){
  const facilitySheetName = PropertiesService.getScriptProperties().getProperty('facilitySheetName');
  const inputSheetName = PropertiesService.getScriptProperties().getProperty('inputSheetName');
  const constTarget = ['仕様書', facilitySheetName, 'siteidH31.3', 'siteChangeInfo', 'segment', inputSheetName, 'out'];
  const targetSheets = getTargetSheets(constTarget);
  return targetSheets;
}
/**
* Delete the output sheets
* @param {spreadsheet} ss: The target spreadsheet
* @return none
*/
function deleteSheets(ss=SpreadsheetApp.getActiveSpreadsheet()){
  const sheetsNonTargetDeletion = getSheetsNonTargetPrinting();
  const sheetNamesNonTargetDeletion = sheetsNonTargetDeletion.map(x => x.getName());
  const targetSheets = ss.getSheets().filter(function(x){
    if (this.indexOf(x.getName()) == -1){
      return x;
    };
  }, sheetNamesNonTargetDeletion);
  targetSheets.map(x => ss.deleteSheet(x));
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
* Order: sheets not to be printed, facility name, year
* @param {spreadsheet} ss: The target spreadsheet
* @return none
*/
function sortSheets(ss=SpreadsheetApp.getActiveSpreadsheet()){
  const targetFacilities = getTargetFacilitiesNames();
  const targetYears = getTargetYears();
  const targetSheetsName = targetFacilities.concat(targetYears);
  // Move non-printable sheets to the left
  const sheetsNonTargetPrinting = getSheetsNonTargetPrinting();
  const sheetsTargetPrinting = getTargetSheets(targetSheetsName);
  const moveSheets = sheetsNonTargetPrinting.concat(sheetsTargetPrinting);
  moveSheets.map(function(x, idx){
    x.activate();
    ss.moveActiveSheet(idx + 1);
  });
}
/**
* Getting a Sheet object from an array of sheet names
* @param {string[]} sheetNames An array of sheet names
* @param {spreadsheet} ss: The target spreadsheet
* @return {sheet} the sheet objects
*/
function getTargetSheets(sheetNames, ss=SpreadsheetApp.getActiveSpreadsheet()){
  var targetSheets = sheetNames.map(x => ss.getSheetByName(x));
  // If the sheet with that name does not exist, remove it from the array
  targetSheets = targetSheets.filter(Boolean);
  return targetSheets;
}
/**
* Create a sheet at the end
* @param {string} sheetName the sheet name
* @param {spreadsheet} ss: The target spreadsheet
* @return {sheet} the sheet objects
*/
function addSheetToEnd(sheetName, ss=SpreadsheetApp.getActiveSpreadsheet()){
  const tempIdx = ss.getNumSheets();
  const targetSheet = ss.insertSheet(sheetName, parseInt(tempIdx)); 
  return targetSheet;  
}
