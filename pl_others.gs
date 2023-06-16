/**
* Create charts for non-clinical research centers
* @param none
* @return none
*/
function execCreateChartOthersMain(){
  const propertyNameAndSpreadsheetNames = new GetSpreadsheetNames().getOthersArray();
  const spreadSheets = createOutputSpreadSheet_(propertyNameAndSpreadsheetNames);
  const countMax = 33;
  const targetFacilitiesCodeAndName = getTargetFacilitiesCodeAndName(false);
  const targetFacilitiesCount = targetFacilitiesCodeAndName.length;
  const lastSpreadSheetIndex = 3;
  const targetArray = spreadSheets.map((spreadsheet, idx) => {
    const count = idx === lastSpreadSheetIndex ? targetFacilitiesCount - countMax * lastSpreadSheetIndex : countMax;
    return [spreadsheet, count];
  });
  targetArray.forEach(([spreadsheet, count], idx) => {
    const startIndex = idx * countMax;
    const endIndex = startIndex + count;
    const target = targetFacilitiesCodeAndName.slice(startIndex, endIndex);
    const condition = {ss: spreadsheet, 
                       sheetCount: count, 
                       facilities: target};
    // Delete the output sheets
    const tempDelSheet = deleteSheets(condition.ss);
    // Output sheets for each facility
    condition.facilities.forEach(x => execCreateChart(x, condition, true));
    condition.ss.deleteSheet(tempDelSheet);
    global_accessPermission = false;
  });
}  
