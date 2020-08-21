/**
* Create a chart
* @param {object} chartConditions　 
* @param {sheet} inputSheet: Source of the chart　 
* @param {sheet} outputSheet: Sheet to output a chart　 
* @param {number} outputChartRow: Row to output a chart　 
* @return none 
*/
function createChart(chartConditions, inputSheet, outputSheet, outputChartRow){
  // Reference URL
  // https://developers.google.com/chart/interactive/docs/gallery/combochart
  var table_outputRow = chartConditions.tableStartRow;
  const tableStartCol = 2;
  const chartTitleFontSize = 20;
  const annotationsFontSize = 10;
  const legendFontSize = 10;
  const vAxesFontSize = 10;
  const hAxisFontSize = 8;
  const tableFontSize = 10;
  var newChart = outputSheet.newChart()
  // Data Range for a Chart
  .addRange(inputSheet.getRange(chartConditions.rangeAddress_x))
  .addRange(inputSheet.getRange(chartConditions.rangeAddress_y0))
  .addRange(inputSheet.getRange(chartConditions.rangeAddress_y1))
  .addRange(inputSheet.getRange(chartConditions.rangeAddress_y2))
  .setChartType(Charts.ChartType.COMBO)
  // The position to output a chart
  .setPosition(outputChartRow, 1, 0, 0)
  // Title
  .setOption('title', chartConditions.title)
  .setOption('titleTextStyle', {color: 'black', fontSize: chartTitleFontSize})
  // Type of a chart
  .setOption('series', {
    0: {type: 'bars', color:'#666666', labelInLegend:chartConditions.labelRevenue, targetAxisIndex: 0, 
        dataLabel: 'value', dataLabelPlacement: 'outsideEnd', annotations: {textStyle: {color: 'gray', fontSize: annotationsFontSize}}}, 
    1: {type: 'bars', color:'#cccccc', labelInLegend:chartConditions.labelCost, targetAxisIndex: 0, 
        dataLabel: 'value', dataLabelPlacement: 'outsideEnd', annotations: {textStyle: {color: 'gray', fontSize: annotationsFontSize}}},
    2: {type: 'line', color: 'black', labelInLegend:chartConditions.labelProfit, lineWidth: 2, pointSize: 7, targetAxisIndex: 1, 
        dataLabel: 'value', dataLabelPlacement: 'outsideEnd', annotations: {textStyle: {color: 'black', fontSize: annotationsFontSize}}}
  })
  //Setting a legend
  .setOption('legend', {position: 'top', textStyle: {color: 'black', fontSize: legendFontSize}}) 
  // Chart size
  .setOption('width', 700)
  .setOption('height', 462)
  // Putting a slant on the output string on the horizontal axis
  .setOption('hAxis.slantedText', true)
  .setOption('hAxis.slantedTextAngle', 60)
  // Set the font size of the horizontal axis
  .setOption('hAxis.textStyle.fontSize', hAxisFontSize)
  // Vertical axis setting
  if (chartConditions.targetCost == chartConditions.constClinicalResearch){
    newChart.setOption('vAxes', {
      0: {title:'金額(百万円)', titleTextStyle: {fontSize: vAxesFontSize}, format: 'decimal', viewWindowMode: 'explicit', viewWindow: {min: -2000, max: 2000}, textStyle: {fontSize: vAxesFontSize}},   
      1: {format: 'decimal', viewWindowMode: 'explicit', viewWindow: {min: -2000, max: 2000}, textPosition: 'none'}
    })
  } else {
    newChart.setOption('vAxes', {
      0: {title:'金額(百万円)', titleTextStyle: {fontSize: vAxesFontSize}, format: 'decimal', viewWindowMode: 'explicit', viewWindow: {min: -30000, max: 30000}, gridlines:{count: 5}, textStyle: {fontSize: vAxesFontSize}},   
      1: {format: 'decimal', viewWindowMode: 'explicit', viewWindow: {min: -30000, max: 30000}, gridlines:{count: 5}, textPosition: 'none'}
    })
  };
  // Output a chart
  outputSheet.insertChart(newChart.build());
  // Set the grid line to "Do not output"
  outputSheet.setHiddenGridlines(true);
}
/**
* Create a chart
* @param {object} target: chartConditions　 
* @return none 
*/
function executeCreateChart(target){
  // Create a sheet to output a chart
  const chart_outputSheet = addSheetToEnd(target.chartSheetName); 
  // ClinicalResearch
  var chartConditionsClinicalResearch = target;
  chartConditionsClinicalResearch.clinicalResearch = '';
  executeCreateChartCommon(chartConditionsClinicalResearch, chart_outputSheet);
  // Ordinary
  var chartConditionsOrdinary = target;
  chartConditionsOrdinary.ordinary = '';
  executeCreateChartCommon(chartConditionsOrdinary, chart_outputSheet);
}
class classSetChartConditions{
  constructor(target){
    // Output in millions of yen
    const items_unit = '1000000';
    this.constClinicalResearch = '（臨床研究）';
    this.constOrdinary = '（全体）';
    this.startRow = 2;
    this.endRow = 30;
    this.selectItems_D = 'sum(D)/' + items_unit;
    this.selectItems_E = 'sum(E)/' + items_unit;
    this.selectItems_F = 'sum(F)*-1/' + items_unit;
    this.selectItems_G = 'sum(G)*-1/' + items_unit;
    this.selectItems_H = 'sum(H)/' + items_unit;
    this.selectItems_I = 'sum(I)/' + items_unit;
    this.labelRevenue = '収益';
    this.labelCost = '費用';
    this.labelProfit = '利益';
    this.targetName = target.name;
    this.chartSheetName = target.chartSheetName;
  }
  set clinicalResearch(value){
    this.targetCost = this.constClinicalResearch;
	this.col_y0 = 'L';
	this.col_y1 = 'N';
	this.col_y2 = 'Q';
	this.revenueItem = 'sum(D)/1000000';
	this.costItem = 'sum(F)*-1/1000000';
	this.profitItem = 'sum(I)/1000000';
    this.outputRow = 1;
  }
  get clinicalResearch(){
    return this;
  }
  set ordinary(value){
    this.targetCost = this.constOrdinary;
	this.col_y0 = 'M';
	this.col_y1 = 'O';
	this.col_y2 = 'P';
	this.revenueItem = 'sum(E)/1000000';
	this.costItem = 'sum(G)*-1/1000000';
	this.profitItem = 'sum(H)/1000000';
    this.outputRow = 24;
  }
  get ordinary(){
    return this;
  }
}
class classSetChartConditionsByFacility extends classSetChartConditions{
  constructor(target){
    const yearsCol = PropertiesService.getScriptProperties().getProperty('inputSheetyearsCol');
    super(target);
    this.dataCol = PropertiesService.getScriptProperties().getProperty('inputSheetfacilityCodeCol');
    this.xCol = yearsCol;
    this.orderbyCol = yearsCol;
    this.ymdCondition = '';
    // Address of the cell range of the input data for Chart output
    this.condition = this.targetName;
    this.strB = "'unused_B'";
    this.groupBy = 'C, K';
  }
}
class classSetChartConditionsByYear extends classSetChartConditions{
  constructor(target){
    const targetSegment = PropertiesService.getScriptProperties().getProperty('clinicalResearchCenter');
    super(target);
    this.dataCol = PropertiesService.getScriptProperties().getProperty('inputSheetyearsCol');
    this.xCol = PropertiesService.getScriptProperties().getProperty('inputSheetfacilityNameCol');
    this.orderbyCol = PropertiesService.getScriptProperties().getProperty('inputSheetfacilityCodeCol');
    this.ymdCondition = "and J = '" + targetSegment + "' ";
    // Address of the cell range of the input data for Chart output
    this.condition = "'" + this.targetName + "'";
    this.strB = 'B';
    this.groupBy = 'B, C, K';
  }
}
/**
* Setting up common settings for chart creation
* @param {object} chartConditions　 
* @param {object} chart_outputSheet: Sheet to output a chart　 
* @return none 
*/
function executeCreateChartCommon(chartConditions, chart_outputSheet){
  // Address of the cell range of the input data for Chart output
  chartConditions.rangeAddress_x = chartConditions.xCol + chartConditions.startRow + ':' + chartConditions.xCol + chartConditions.endRow;
  chartConditions.rangeAddress_y0 = chartConditions.col_y0 + chartConditions.startRow + ':' + chartConditions.col_y0 + chartConditions.endRow;
  chartConditions.rangeAddress_y1 = chartConditions.col_y1 + chartConditions.startRow + ':' + chartConditions.col_y1 + chartConditions.endRow;
  chartConditions.rangeAddress_y2 = chartConditions.col_y2 + chartConditions.startRow + ':' + chartConditions.col_y2 + chartConditions.endRow;
  // The name of the sheet to output the input data for Chart output
  chartConditions.outputSheetName = chartConditions.chartSheetName + chartConditions.targetCost;
  // Chart title 
  chartConditions.title = chartConditions.outputSheetName;
  const inputSheet = createQuerySheet(chartConditions);
  createChart(chartConditions, inputSheet, chart_outputSheet, chartConditions.outputRow);
}  