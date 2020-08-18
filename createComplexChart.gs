// inputSheetの情報からグラフを作成する
function createChart(chartConditions, inputSheet, outputSheet, outputChartRow){
  if (chartConditions.targetName == '近畿中央胸部疾患センター'){
    return;
  }
  var table_outputRow = chartConditions.tableStartRow;
  const tableStartCol = 2;
  // chart object生成
  const chartTitleFontSize = 20;
  const annotationsFontSize = 10;
  const legendFontSize = 10;
  const vAxesFontSize = 10;
  const hAxisFontSize = 8;
  const tableFontSize = 10;
  var newChart = outputSheet.newChart()
  // グラフ作成用データ範囲
  .addRange(inputSheet.getRange(chartConditions.rangeAddress_x))
  .addRange(inputSheet.getRange(chartConditions.rangeAddress_y0))
  .addRange(inputSheet.getRange(chartConditions.rangeAddress_y1))
  .addRange(inputSheet.getRange(chartConditions.rangeAddress_y2))
  .setChartType(Charts.ChartType.COMBO)  // 複合グラフ
  .setPosition(outputChartRow, 1, 0, 0)  // 出力セル
  // グラフのタイトル
  .setOption('title', chartConditions.title)
  .setOption('titleTextStyle', {color: 'black', fontSize: chartTitleFontSize})
  // グラフの種類
  .setOption('series', {
    0: {type: 'bars', color:'#666666', labelInLegend:chartConditions.labelRevenue, targetAxisIndex: 0, 
        dataLabel: 'value', dataLabelPlacement: 'outsideEnd', annotations: {textStyle: {color: 'gray', fontSize: annotationsFontSize}}}, 
    1: {type: 'bars', color:'#cccccc', labelInLegend:chartConditions.labelCost, targetAxisIndex: 0, 
        dataLabel: 'value', dataLabelPlacement: 'outsideEnd', annotations: {textStyle: {color: 'gray', fontSize: annotationsFontSize}}},
    2: {type: 'line', color: 'black', labelInLegend:chartConditions.labelProfit, lineWidth: 2, pointSize: 7, targetAxisIndex: 1, 
        dataLabel: 'value', dataLabelPlacement: 'outsideEnd', annotations: {textStyle: {color: 'black', fontSize: annotationsFontSize}}}
  })
  .setOption('legend', {position: 'top', textStyle: {color: 'black', fontSize: legendFontSize}})  //凡例の設定
  // グラフのサイズ
  .setOption('width', 700)
  .setOption('height', 462)
  // 横軸に出力する文字列に傾斜をつける
  .setOption('hAxis.slantedText', true)
  .setOption('hAxis.slantedTextAngle', 60)
  // 横軸のフォントサイズ設定
  .setOption('hAxis.textStyle.fontSize', hAxisFontSize)
  // 縦軸の設定
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
  // グラフを出力
  outputSheet.insertChart(newChart.build());
  // グリッド線を出力しないにする
  outputSheet.setHiddenGridlines(true);
}
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
    this.selectItems_D = 'D/' + items_unit;
    this.selectItems_E = 'E/' + items_unit;
    this.selectItems_F = 'F*-1/' + items_unit;
    this.selectItems_G = 'G*-1/' + items_unit;
    this.selectItems_H = 'H/' + items_unit;
    this.selectItems_I = 'I/' + items_unit;
    this.labelRevenue = '収益';
    this.labelCost = '費用';
    this.labelProfit = '利益';
    this.targetName = target.name;
  }
  set clinicalResearch(value){
    this.targetCost = this.constClinicalResearch;
	this.col_y0 = 'L';
	this.col_y1 = 'N';
	this.col_y2 = 'Q';
	this.revenueItem = 'D/1000000';
	this.costItem = 'F*-1/1000000';
	this.profitItem = 'I/1000000';
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
	this.revenueItem = 'E/1000000';
	this.costItem = 'G*-1/1000000';
	this.profitItem = 'H/1000000';
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
  }
}
function executeCreateChartCommon(chartConditions, chart_outputSheet){
  // Address of the cell range of the input data for Chart output
  chartConditions.rangeAddress_x = chartConditions.xCol + chartConditions.startRow + ':' + chartConditions.xCol + chartConditions.endRow;
  chartConditions.rangeAddress_y0 = chartConditions.col_y0 + chartConditions.startRow + ':' + chartConditions.col_y0 + chartConditions.endRow;
  chartConditions.rangeAddress_y1 = chartConditions.col_y1 + chartConditions.startRow + ':' + chartConditions.col_y1 + chartConditions.endRow;
  chartConditions.rangeAddress_y2 = chartConditions.col_y2 + chartConditions.startRow + ':' + chartConditions.col_y2 + chartConditions.endRow;
  // The name of the sheet to output the input data for Chart output
  chartConditions.outputSheetName = chartConditions.targetName + chartConditions.targetCost;
  // Chart title 
  chartConditions.title = chartConditions.outputSheetName;
  const inputSheet = createQuerySheet(chartConditions);
  createChart(chartConditions, inputSheet, chart_outputSheet, chartConditions.outputRow);
}  