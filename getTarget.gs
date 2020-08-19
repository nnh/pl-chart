/**
* Return the names of the clinical research center in a one-dimensional array
* @param none
* @return {string} The names of the clinical research center in a one-dimensional array
*/
function getTargetFacilitiesNames(){
  const colIndex = getColumnNumber(PropertiesService.getScriptProperties().getProperty('siteSheetfacilityNameCol'));
  return getTargetColValueFacilities(colIndex);
}
/**
* Returns a one-dimensional array of values at a specified index
* @param {number} Target array index 
* @return {string} A one-dimensional array of values at a specified index
*/function getTargetColValueFacilities(idx){
  var facilityValues = getTargetFacilitiesValues();
  facilityValues = facilityValues.map(x => x[idx]);
  return facilityValues;
}
/**
* Returns the value of the row where the segment is '臨床研究センター' 
* @param none
* @return {string} The values of the target rows
*/
function getTargetFacilitiesValues(){
  const target = PropertiesService.getScriptProperties().getProperty('clinicalResearchCenter');
  const segmentValuesList = getTargetFacilityList('E', target, true);
  return segmentValuesList;
}
/**
* Extract the values of the rows that match the conditions from the Site sheet
* @param {string} targetCol The column name of the extraction condition, e.g. 'C'
* @param {string} conditionString A string of conditions
* @param {boolean} condition IF true then '==', else '!='.
* @return {string} The values of the target rows
*/
function getTargetFacilityList(targetCol, conditionString, condition){
  const siteSheet = getTargetSheet(PropertiesService.getScriptProperties().getProperty('facilitySheetName'));
  var siteValues = siteSheet.getRange('A:F').getValues();
  const targetIndex = getColumnNumber(targetCol);
  siteValues = siteValues.filter(function(x){
    if (condition){
      return x[targetIndex] == conditionString;
    } else {
      return x[targetIndex] != conditionString;
    }
  });
  return siteValues; 
}
/**
* Return the years in a one-dimensional array
* @param none
* @return {string} The years in a one-dimensional array
*/
function getTargetYears(){
  const workingSheet = getTargetSheet(PropertiesService.getScriptProperties().getProperty('inputSheetName'));
  const yearsCol = PropertiesService.getScriptProperties().getProperty('inputSheetyearsCol');
  const rangeAddr = yearsCol + ':' + yearsCol;
  var years = workingSheet.getRange(rangeAddr).getValues();
  // Making an array from two-dimensional to one-dimensional
  years = years.reduce(function(x, item){
    x.push(...item);
    return x;
  }, []);
  // Removing Duplicates, Nulls and Headings
  years = years.filter(x => x != '年度');
  years = Array.from(new Set(years))
  years = years.filter(Boolean);
  return years;
}
/**
* Returns the name of the most recent facility from the facility code
* @param {number} The facility code
* @return {string} The code and the most recent facility name
*/
function getRecentFacilityName(facilityCode){
  const facilityCodeCol = PropertiesService.getScriptProperties().getProperty('siteSheetfacilityCodeCol');
  const facilityIdx = getColumnNumber(PropertiesService.getScriptProperties().getProperty('siteSheetfacilityNameCol'));
  const facilityList = getTargetFacilityList(facilityCodeCol, facilityCode, true);
  var facilityNames = facilityList.map(x => x[facilityIdx]);
  // If it's one code and one facility name, it returns the name of the facility
  if (facilityNames.length == 1){
    return [facilityCode, facilityNames[0]];
  }
  // If more than one facility name in one code, only the most recent one is returned
  const years = getTargetYears();
  const workingSheetValues = getTargetSheet(PropertiesService.getScriptProperties().getProperty('inputSheetName')).getRange('A:E').getValues();
  const inputYearsIdx = getColumnNumber(PropertiesService.getScriptProperties().getProperty('inputSheetyearsCol'));
  const inputFacilityNameIdx = getColumnNumber(PropertiesService.getScriptProperties().getProperty('inputSheetfacilityNameCol'));
  for (var i = years.length - 1; i >= 0; i--){
    var tempYearValues = workingSheetValues.filter(x => x[inputYearsIdx] == years[i]);
    var tempfacilityByYear = tempYearValues.filter(x => facilityNames.indexOf(x[inputFacilityNameIdx]) > -1);
    if (tempfacilityByYear.length > 0){
      break;
    }
  }
  if (tempfacilityByYear.length == 1){
    return [facilityCode, tempfacilityByYear[0][inputFacilityNameIdx]];
  } else {
    // Not supported if more than one facility is named with the same code in the same year
    return null;
  }
}
