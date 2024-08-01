class GetResearchPerformance{
  constructor(targetList){
    this.target = targetList;
    if (targetList.year < 2019){
      this.targetYearStr = '平成' + String(targetList.year - 2000 + 12) + '年度';
    } else {
      this.targetYearStr = '';
    }
  }
  getData(){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    var researchPerformanceCondition = {};
    researchPerformanceCondition.inputSheet = ss.getSheetByName('working');
    researchPerformanceCondition.plSheet = ss.getSheetByName('research_performance' + this.target.year);
    researchPerformanceCondition.outputSheet = ss.getSheetByName('research' + this.target.year); 
    researchPerformanceCondition.targetYear = this.targetYearStr; 
    researchPerformanceCondition.seirekiYear = this.target.year;
    getResearchPerformance(researchPerformanceCondition);
  }
}
function main(){
  //const targetYears = [2017, 2015, 2014];
  const targetYears = [2016];
  targetYears.forEach(x => getByYear(x));
}
function getByYear(targetYear){
    var researchPerformanceCondition = {};
    researchPerformanceCondition.year = targetYear; 
    const researchPerformance = new GetResearchPerformance(researchPerformanceCondition);
    researchPerformance.getData();
}
function getResearchPerformance(researchPerformanceCondition){
  const inputRawdata = researchPerformanceCondition.inputSheet.getDataRange().getValues();
  if (researchPerformanceCondition.plSheet){
    var plRawdata = researchPerformanceCondition.plSheet.getDataRange().getValues();
  } else {
    // dummy
    var plRawdata = [['dummy', 0], [0, 0]];
  }
  const plSumColumn = plRawdata[0].length - 1;
  const plTargetArray = plRawdata.map(x => [x[0], x[plSumColumn]]);
  const plTargetArrayObj = Object.fromEntries(plTargetArray);
  var header = [['facility_code', '=facility_code!C1', 'revenue', 'profit', 'profit_rate', 'research']];
  var inputData = inputRawdata.filter(x => x[2] == researchPerformanceCondition.targetYear);
  inputData = inputData.map((x, idx) => {
      const targetRow = idx + 2;
      const baseYear = 2018;
      const vlookupOffset = 8 + baseYear - researchPerformanceCondition.seirekiYear;
      const getFacilityName = '=VLOOKUP(A' + targetRow + ',facility_code!A:E,3,FALSE)';
      if (plRawdata[0][0] == 'dummy'){
        var getReserch = '=IFERROR(VLOOKUP(A' + targetRow + ",'research2015-2018'!$A$2:$L$132," + vlookupOffset + ',FALSE),' + '"")';
      } else {
        var getReserch = plTargetArrayObj[x[10]];
      }
      return([x[10], getFacilityName, x[4], x[7], x[7] / x[4], getReserch]);
    });
  const outputValues = header.concat(inputData);
  researchPerformanceCondition.outputSheet.clearContents
  researchPerformanceCondition.outputSheet.getRange(1, 1, outputValues.length, outputValues[0].length).setValues(outputValues);
  // sort by facility_code
  researchPerformanceCondition.outputSheet.getRange(2, 1, researchPerformanceCondition.outputSheet.getLastRow(), researchPerformanceCondition.outputSheet.getLastColumn()).sort(1);
}