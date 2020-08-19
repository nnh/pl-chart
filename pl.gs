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
  const targetFacilities = getTargetFacilitiesValues();
  // Removing Duplicate Code, get the latest name of the facility
  var targetFacilitiesCodes = targetFacilities.map(x => x[0]);
  targetFacilitiesCodes = Array.from(new Set(targetFacilitiesCodes));
  const targetFacilitiesCodeAndName = targetFacilitiesCodes.map(x => getRecentFacilityName(x));
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
