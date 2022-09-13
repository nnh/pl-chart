/**
* Create Charts of clinical research centers by facility and year
* @param none
* @return none
*/
function execCreateChartMain(){
  createOutputSpreadSheet(['PL（臨床研究センター）'], ['outputSpreadsheetIdClinicalResearchCenter']);
  let target = {};
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
