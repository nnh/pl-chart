function modyfiComplexChartMain() {
  const targetFolder = getOutputFolder_();
  if (targetFolder === null){
    return null;
  };
  const targetSheetNames = new GetSpreadsheetNames().getAllArray().map(x => x[1]);
  const files = targetFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  let targetFilesId = new Array();
  while (files.hasNext()){
    const file = files.next();
    if (targetSheetNames.includes(file.getName())){
      targetFilesId.push(file.getId());
    }
  }
  if (targetFilesId.length < 1){
    return;
  }
  targetFilesId.forEach(fileId => modyfiComplexChart_(fileId));
}
function modyfiComplexChart_(spreadsheetId) {  
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const chartConditions = {};
  const clinicalResearchRow = 1;
  spreadsheet.getSheets().forEach(sheet =>{
    const charts = sheet.getCharts();
    if (charts.length === 0){
      return;
    }
    chartConditions.outputSheet = spreadsheet.getSheetByName(sheet.getName());
    charts.forEach(chart => {
      const anchorRow = chart.getContainerInfo().getAnchorRow();
      const inputSheetName = anchorRow === clinicalResearchRow 
        ? `${sheet.getName()}${PropertiesService.getScriptProperties().getProperty('clinicalResearchSheetNameFooter')}` 
        : `${sheet.getName()}${PropertiesService.getScriptProperties().getProperty('ordinarySheetNameFooter')}`; 
      chartConditions.inputSheet = spreadsheet.getSheetByName(inputSheetName);
      chartConditions.outputSheetName = inputSheetName;
      [chartConditions.rangeAddress_x, chartConditions.rangeAddress_y0, chartConditions.rangeAddress_y1, chartConditions.rangeAddress_y2] = chart.getRanges().map(range => range.getA1Notation());
      const createChart = anchorRow === clinicalResearchRow ? new CreateChart(chartConditions) : new CreateChartOverAll(chartConditions);
      createChart.createInsertChart();
      createChart.removeChart(chart);
    });
  });
}
