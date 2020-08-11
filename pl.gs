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
  const facilityNameCol = PropertiesService.getScriptProperties().getProperty('inputSheetfacilityNameCol');
  const yearsCol = PropertiesService.getScriptProperties().getProperty('inputSheetyearsCol');
  const facilityNameIdx = getColumnNumber(facilityNameCol);
  targetFacilities.map(function(targetFacility){
    target.name = targetFacility[facilityNameIdx];
    target.dataCol = facilityNameCol;  // Output by facility name
    target.xCol = yearsCol;            // Horizontal axis:years
    target.orderbyCol = yearsCol;      // Sort the output order of the horizontal axis by year
    target.ymdCondition = '';
    executeCreateChart(target);
  });
  // Output sheets for each year
  const targetYears = getTargetYears();
  const targetSegment = PropertiesService.getScriptProperties().getProperty('clinicalResearchCenter');
  targetYears.map(function(targetYear){
    target.name = targetYear;
    target.dataCol = yearsCol;        // Output by year
    target.xCol = facilityNameCol;    // Horizontal axis:facility name
    target.orderbyCol = 'K';          // Sort the output order of the horizontal axis by facility code
    target.ymdCondition = "and J = '" + targetSegment + "' ";  // Only if the segment is a clinical research center
    executeCreateChart(target);
  });
  // Sort the all sheets
  sortSheets();
}
