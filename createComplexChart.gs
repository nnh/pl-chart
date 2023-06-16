class classSetChartConditions{
  constructor(target){
    this.ss = target.ss;
    // Output in millions of yen
    const items_unit = '1000000';
    this.constClinicalResearch = PropertiesService.getScriptProperties().getProperty('clinicalResearchSheetNameFooter');
    this.constOrdinary = PropertiesService.getScriptProperties().getProperty('ordinarySheetNameFooter');
    this.startRow = 2;
    this.endRow = 30;
    this.selectItems_D = 'sum(Col4)/' + items_unit;
    this.selectItems_E = 'sum(Col5)/' + items_unit;
    this.selectItems_F = 'sum(Col6)*-1/' + items_unit;
    this.selectItems_G = 'sum(Col7)*-1/' + items_unit;
    this.selectItems_H = 'sum(Col8)/' + items_unit;
    this.selectItems_I = 'sum(Col9)/' + items_unit;
    this.labelRevenue = '収益';
    this.labelCost = '費用';
    this.labelProfit = '利益';
    this.targetName = target.name;
    this.chartSheetName = target.chartSheetName;
    this.selectItems_L = 'Col12';
  }
  set clinicalResearch(value){
    this.targetCost = this.constClinicalResearch;
	this.col_y0 = 'L';
	this.col_y1 = 'N';
	this.col_y2 = 'Q';
	this.revenueItem = 'sum(Col4)/1000000';
	this.costItem = 'sum(Col6)*-1/1000000';
	this.profitItem = 'sum(Col9)/1000000';
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
	this.revenueItem = 'sum(Col5)/1000000';
	this.costItem = 'sum(Col7)*-1/1000000';
	this.profitItem = 'sum(Col8)/1000000';
    this.outputRow = 24;
  }
  get ordinary(){
    return this;
  }
}
class classSetChartConditionsByFacility extends classSetChartConditions{
  constructor(target){
    const xColName = PropertiesService.getScriptProperties().getProperty('inputSheetyearsCol');
    const yearsColInfo = new classGetColumnInfo(xColName);
    const yearsCol = yearsColInfo.queryColumnName;
    super(target);
    const dataColInfo = new classGetColumnInfo(PropertiesService.getScriptProperties().getProperty('inputSheetfacilityCodeCol'));
    this.dataCol = dataColInfo.queryColumnName;
    this.xColName = xColName;
    this.xCol = yearsCol;
    this.orderbyCol = 'Col12';
    this.ymdCondition = '';
    // Address of the cell range of the input data for Chart output
    this.condition = this.targetName;
    this.strB = "'unused_B'";
    this.groupBy = 'Col3, Col11, Col12';
  }
}
class classSetChartConditionsByYear extends classSetChartConditions{
  constructor(target){
    const targetSegment = PropertiesService.getScriptProperties().getProperty('clinicalResearchCenter');
    super(target);
    const dataColInfo = new classGetColumnInfo(PropertiesService.getScriptProperties().getProperty('inputSheetyearsCol'));
    this.dataCol = dataColInfo.queryColumnName;
    this.xColName = PropertiesService.getScriptProperties().getProperty('inputSheetfacilityNameCol');
    const xColInfo = new classGetColumnInfo(this.xColName);
    this.xCol = xColInfo.queryColumnName;
    const orderbyColInfo = new classGetColumnInfo(PropertiesService.getScriptProperties().getProperty('inputSheetfacilityCodeCol'));
    this.orderbyCol = orderbyColInfo.queryColumnName;
    this.ymdCondition = "and Col10 = '" + targetSegment + "' ";
    // Address of the cell range of the input data for Chart output
    this.condition = "'" + this.targetName + "'";
    this.strB = 'Col2';
    this.groupBy = 'Col2, Col3, Col11';
    this.selectItems_L = null;
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
  chartConditions.rangeAddress_x = chartConditions.xColName + chartConditions.startRow + ':' + chartConditions.xColName + chartConditions.endRow;
  chartConditions.rangeAddress_y0 = chartConditions.col_y0 + chartConditions.startRow + ':' + chartConditions.col_y0 + chartConditions.endRow;
  chartConditions.rangeAddress_y1 = chartConditions.col_y1 + chartConditions.startRow + ':' + chartConditions.col_y1 + chartConditions.endRow;
  chartConditions.rangeAddress_y2 = chartConditions.col_y2 + chartConditions.startRow + ':' + chartConditions.col_y2 + chartConditions.endRow;
  // The name of the sheet to output the input data for Chart output
  chartConditions.outputSheetName = chartConditions.chartSheetName + chartConditions.targetCost;
  // Chart title 
  chartConditions.title = chartConditions.outputSheetName;
  const inputSheet = createQuerySheet(chartConditions);
  chartConditions.inputSheet = inputSheet;
  return chartConditions;
//  createChart(chartConditions, inputSheet, chart_outputSheet, chartConditions.outputRow);
}
/**
* Create a chart.
*/ 
class CreateChart{
  constructor(chartConditions){
    this.inputSheet = chartConditions.inputSheet;
    this.chartTitle = chartConditions.outputSheetName;
    this.outputSheet = chartConditions.outputSheet;
    this.rangeAddress_x = chartConditions.rangeAddress_x;
    this.rangeAddress_y0 = chartConditions.rangeAddress_y0;
    this.rangeAddress_y1 = chartConditions.rangeAddress_y1;
    this.rangeAddress_y2 = chartConditions.rangeAddress_y2;
    this.chartRangeMax = 2000;
    this.chartRangeMin = -2000;
    this.inputSheetLastRow = this.inputSheet.getLastRow().toFixed();
    this.outputRow = 1;
    // Set the grid line to "Do not output"
    this.outputSheet.setHiddenGridlines(true);
  };
  setChartOptions(chart){
    return chart
      .asComboChart()
      .addRange(this.inputSheet.getRange(this.rangeAddress_x))
      .addRange(this.inputSheet.getRange(this.rangeAddress_y0))
      .addRange(this.inputSheet.getRange(this.rangeAddress_y1))
      .addRange(this.inputSheet.getRange(this.rangeAddress_y2))
      .setPosition(this.outputRow, 1, 0, 0)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(0)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('legend.position', 'top')
      .setOption('title', this.chartTitle)  
      .setOption('annotations.domain.textStyle.color', '#000000')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.fontSize', 10)
      .setOption('legend.textStyle.color', '#000000')
      .setOption('titleTextStyle.fontSize', 20)
      .setOption('titleTextStyle.color', '#000000')
      .setOption('annotations.total.textStyle.color', '#000000')
      .setOption('hAxis.slantedText', true)
      .setOption('hAxis.slantedTextAngle', 60)
      .setOption('hAxis.textStyle.fontSize', 8)
      .setOption('hAxis.textStyle.color', '#000000')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setOption('vAxes.0.titleTextStyle.color', '#000000')
      .setOption('series.0.viewWindowMode', 'explicit')
      .setOption('series.0.dataLabel', 'value')
      .setOption('series.1.dataLabel', 'value')
      .setOption('series.2.dataLabel', 'value')
      .setOption('series.0.labelInLegend', '収益')
      .setOption('series.1.labelInLegend', '費用')
      .setOption('series.2.labelInLegend', '利益')
      .setOption('series.0.type', 'bars')
      .setOption('series.1.type', 'bars')
      .setOption('series.2.pointSize', 7)
      .setYAxisTitle('単位（百万円）')
      .setOption('height', 462)
      .setOption('width', 700)
      .setRange(this.chartRangeMin, this.chartRangeMax)
      .setColors(["gray", "silver", "black"])
      .build();
  }
  createChartObject(){
    return this.setChartOptions(this.outputSheet.newChart());
  };
  insertChart(chart){
    this.outputSheet.insertChart(chart);
  };
  createInsertChart(){
    this.insertChart(this.createChartObject());
  };
  removeChart(chart){
    this.outputSheet.removeChart(chart);
  }
};
class CreateChartOverAll extends CreateChart{
  constructor(chartConditions){
    super(chartConditions);
    this.chartRangeMax = 30000;
    this.chartRangeMin = -30000;
    this.outputRow = 24;
  };
}
function executeCreateChart(target){
  const chart_outputSheet = addSheetToEnd(target.chartSheetName, target.ss); 
  // ClinicalResearch
  let chartConditionsClinicalResearch = target;
  chartConditionsClinicalResearch.clinicalResearch = '';
  let chartConditions = executeCreateChartCommon(chartConditionsClinicalResearch, chart_outputSheet);
  chartConditions.outputSheet = chart_outputSheet;
  new CreateChart(chartConditions).createInsertChart();
  // Ordinary
  let chartConditionsOrdinary = target;
  chartConditionsOrdinary.ordinary = '';
  chartConditions = executeCreateChartCommon(chartConditionsOrdinary, chart_outputSheet);
  chartConditions.outputSheet = chart_outputSheet;
  new CreateChartOverAll(chartConditions).createInsertChart();
}; 