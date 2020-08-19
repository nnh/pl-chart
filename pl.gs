/**
* Create Charts of clinical research centers by facility and year
* @param none
* @return none
*/
function execCreateChart(){
  // Delete the output sheets
  deleteSheets();
  var target = {};
  const targetFacilitiesCodeAndName = getTargetFacilitiesCodeAndName(true);
  // Output sheets for each facility
  targetFacilitiesCodeAndName.map(function(targetFacility){
    target.name = targetFacility[0];
    target.chartSheetName = targetFacility[1];
    var chartConditions = new classSetChartConditionsByFacility(target);
    executeCreateChart(chartConditions);
  });
  // Output sheets for each year
  const targetYears = getTargetYears();
  targetYears.map(function(targetYear){
    target.name = targetYear;
    target.chartSheetName = target.name;
    var chartConditions = new classSetChartConditionsByYear(target);
    executeCreateChart(chartConditions);
  });
  // Sort the all sheets
  sortSheets();
}
/**
* 臨床研究センター以外出力
* @param none
* @return none
*/
function execCreateChartOthersMain(){
  const countMax = 33;
  const outputSS_1 = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('outputSpreadsheetIdOthers1'));
  const targetFacilitiesCodeAndName = getTargetFacilitiesCodeAndName(false);
  const outputSS1Count = countMax;
  const outputSS2Count = countMax;
  const outputSS3Count = countMax;
  const outputSS4Count = targetFacilitiesCodeAndName.length - countMax * 3;
  execCreateChartOthers(outputSS_1);
}
function execCreateChartOthers(ss){
  // 一つだけシートを残して他は全て削除する
  const deleteSheetName = 'temp_del';
  var tempSheets = ss.getSheets();
  const tempSheet = tempSheets[0];
  if (tempSheets.length > 1){
    tempSheets.splice(0, 1);
    tempSheets.map(x => ss.deleteSheet(x));
  }
  tempSheet.setName(deleteSheetName);
  
  
  // 不要シートを削除
  if (ss.getSheets().length > 1){
    ss.deleteSheet(tempSheet);
  }
}

