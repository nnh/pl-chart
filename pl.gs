/**
* Create Charts of clinical research centers by facility and year
* @param none
* @return none
*/
function execCreateChart(){
  var target = {};
  target.ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('outputSpreadsheetIdClinicalResearchCenter'));
  // Delete the output sheets
  const tempDelSheet = deleteSheets(target.ss);
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
  //sortSheets(target.ss);
}
/**
* 臨床研究センター以外出力
* @param none
* @return none
*/
function execCreateChartOthersMain(){
  const countMax = 33;
  const targetFacilitiesCodeAndName = getTargetFacilitiesCodeAndName(false);
  const targetFacilitiesCount = targetFacilitiesCodeAndName.length;
  const targetArray = [['outputSpreadsheetIdOthers1', countMax],
/*                       ['outputSpreadsheetIdOthers2', countMax],
                       ['outputSpreadsheetIdOthers3', countMax],*/
                       ['outputSpreadsheetIdOthers4', targetFacilitiesCount - countMax * 3]];
  targetArray.map(function(x, idx){
    const startIndex = idx * countMax;
    const endIndex = startIndex + x[1];
    const target = targetFacilitiesCodeAndName.slice(startIndex, endIndex);
    const condition = {ss: x[0], sheetCount: x[1], facilities: target};
    execCreateChartOthers(condition);
  });
}  
function execCreateChartOthers(condition){
  var target = condition;
  const deleteSheetName = 'temp_del';
  // 一つだけシートを残して他は全て削除する
  var tempSheets = target.ss.getSheets();
  const tempSheet = tempSheets[0];
  if (tempSheets.length > 1){
    tempSheets.splice(0, 1);
    tempSheets.map(x => target.ss.deleteSheet(x));
  }
  // 残したシートは最後に削除する
  tempSheet.setName(deleteSheetName);
  // Output sheets for each facility
  condition.facilities.map(function(targetFacility){
    target.name = targetFacility[0];
    target.chartSheetName = targetFacility[1];
    const chartConditions = new classSetChartConditionsByFacility(target);
    executeCreateChart(chartConditions);
  });
  
  
  // 不要シートを削除
  if (ss.getSheets().length > 1){
    ss.deleteSheet(tempSheet);
  }
}

