// https://developers.google.com/chart/interactive/docs/gallery/combochart

/**
* Create Charts of clinical research centers by facility and year
* @param none
* @return none
*/
function execCreateChart(){
  // Delete the output sheets
  deleteSheets();
  var target = {};
  // Output sheets for each facility
  const targetFacilities = getTargetFacilitiesValues();
  const facilityNameIdx = getColumnNumber(PropertiesService.getScriptProperties().getProperty('inputSheetfacilityNameCol'));
  targetFacilities.map(function(targetFacility){
    target.name = targetFacility[facilityNameIdx];
    var chartConditions = new classSetChartConditionsByFacility(target);
    executeCreateChart(chartConditions);
  });
  return;
  // Output sheets for each year
  const targetYears = getTargetYears();
  targetYears.map(function(targetYear){
    target.name = targetYear;
    var chartConditions = new classSetChartConditionsByYear(target);
    executeCreateChart(chartConditions);
  });
  // Sort the all sheets
  sortSheets();
}
