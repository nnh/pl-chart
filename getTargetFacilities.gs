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
  const segmentValuesList = getTargetFacilityList('E', '臨床研究センター', true);
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const siteSheet = ss.getSheetByName('site');
  var siteValues = siteSheet.getRange('A:F').getValues();
  const targetIndex = getColumnNumber(siteSheet, targetCol)
  siteValues = siteValues.filter(function(x){
    if (condition){
      return x[targetIndex] == conditionString;
    } else {
      return x[targetIndex] != conditionString;
    }
  })
  return siteValues; 
}
/**
* Returns an array index from a column name
* @param {sheet} targetSheet sheet object
* @param {string} columnName Column name (e.g. 'A')
* @return {number} The corresponding array index of getValues (e.g. return 0 if 'A')
*/
function getColumnNumber(targetSheet, columnName){ 
  var colNumber = targetSheet.getRange(columnName + '1').getColumn();
  colNumber--;
  return colNumber;
}
