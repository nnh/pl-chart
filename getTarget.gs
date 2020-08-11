/**
* Return the names of the clinical research center in a one-dimensional array
* @param none
* @return {string} The names of the clinical research center in a one-dimensional array
*/
function getTargetFacilitiesNames(){
  var facilityValues = getTargetFacilitiesValues();
  facilityValues = facilityValues.map(x => x[1]);
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
