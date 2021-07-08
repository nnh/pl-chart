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
    getResearchPerformance(researchPerformanceCondition);
  }
}
function main(){
  const targetYears = [2017, 2015, 2014];
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
  var header = [['facility_code', 'revenue', 'profit', 'profit_rate', 'research']];
  var inputData = inputRawdata.filter(x => x[2] == researchPerformanceCondition.targetYear);
  inputData = inputData.map(x => [x[10], x[4], x[7], x[7] / x[4], plTargetArrayObj[x[10]]]);
  const outputValues = header.concat(inputData);
  researchPerformanceCondition.outputSheet.clearContents
  researchPerformanceCondition.outputSheet.getRange(1, 1, outputValues.length, outputValues[0].length).setValues(outputValues);
}