/**
* Create charts for non-clinical research centers
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
