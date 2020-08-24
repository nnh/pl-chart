/**
* Create Charts of clinical research centers by facility and year
* @param none
* @return none
*/
function execCreateChartMain(){
  var target = {};
  target.ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('outputSpreadsheetIdClinicalResearchCenter'));
  // Delete the output sheets
  const tempDelSheet = deleteSheets(target.ss);
  const targetFacilitiesCodeAndName = getTargetFacilitiesCodeAndName(true);
  const targetYears = getTargetYears();
  // Output sheets for each facility
  targetFacilitiesCodeAndName.forEach(x => execCreateChart(x, target, true));
  // Output sheets for each year
  targetYears.forEach(x => execCreateChart(x, target, false));
  target.ss.deleteSheet(tempDelSheet);
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
                       ['outputSpreadsheetIdOthers2', countMax],
                       ['outputSpreadsheetIdOthers3', countMax],
                       ['outputSpreadsheetIdOthers4', targetFacilitiesCount - countMax * 3]];
  targetArray.forEach(function(x, idx){
    const startIndex = idx * countMax;
    const endIndex = startIndex + x[1];
    const target = targetFacilitiesCodeAndName.slice(startIndex, endIndex);
    const condition = {ss: SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty(x[0])), 
                       sheetCount: x[1], 
                       facilities: target};
    // Delete the output sheets
    const tempDelSheet = deleteSheets(condition.ss);
    // Output sheets for each facility
    condition.facilities.forEach(x => execCreateChart(x, condition, true));
    condition.ss.deleteSheet(tempDelSheet);
  });
}  
