function main(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var researchPerformanceCondition = {};
  researchPerformanceCondition.inputSheet = ss.getSheetByName('working');
  researchPerformanceCondition.plSheet = ss.getSheetByName('research_performance2018');
  researchPerformanceCondition.outputSheet = ss.getSheetByName('research2018'); 
  researchPerformanceCondition.targetYear = '平成30年度'; 
  getResearchPerformance(researchPerformanceCondition);
}
function getResearchPerformance(researchPerformanceCondition){
  const inputRawdata = researchPerformanceCondition.inputSheet.getDataRange().getValues();
  const plRawdata = researchPerformanceCondition.plSheet.getDataRange().getValues();
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

