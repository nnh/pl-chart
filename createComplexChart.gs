// inputSheetの情報からグラフを作成する
function createChart(chartConditions, inputSheet, outputSheet, outputChartRow){
  if (chartConditions.targetName == '近畿中央胸部疾患センター'){
    return;
  }
  //
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

function getCommonConstant(target){
  // Output in millions of yen
  const items_unit = '1000000';
  var chartConditions = {};
  chartConditions.constClinicalResearch = '（臨床研究）';
  chartConditions.constOrdinary = '（全体）';
  chartConditions.startRow = 2;
  chartConditions.endRow = 30;
  chartConditions.selectItems_D = 'D/' + items_unit;
  chartConditions.selectItems_E = 'E/' + items_unit;
  chartConditions.selectItems_F = 'F*-1/' + items_unit;
  chartConditions.selectItems_G = 'G*-1/' + items_unit;
  chartConditions.selectItems_H = 'H/' + items_unit;
  chartConditions.selectItems_I = 'I/' + items_unit;
  chartConditions.labelRevenue = '収益';
  chartConditions.labelCost = '費用';
  chartConditions.labelProfit = '利益';
  chartConditions.targetName = target.name;
  // The name of the sheet to output the Chart
  chartConditions.chartSheetName = chartConditions.targetName;
  chartConditions.xCol = target.xCol;
  // Address of the cell range of the input data for Chart output
  chartConditions.rangeAddress_x = chartConditions.xCol + chartConditions.startRow + ':' + chartConditions.xCol + chartConditions.endRow;
  chartConditions.condition = chartConditions.targetName;
  chartConditions.dataCol = target.dataCol;
  chartConditions.orderbyCol = target.orderbyCol;
  chartConditions.ymdCondition = target.ymdCondition;
  return chartConditions;
}
function setChartRange(chartConditions){
  // Address of the cell range of the input data for Chart output
  chartConditions.rangeAddress_y0 = chartConditions.col_y0 + chartConditions.startRow + ':' + chartConditions.col_y0 + chartConditions.endRow;
  chartConditions.rangeAddress_y1 = chartConditions.col_y1 + chartConditions.startRow + ':' + chartConditions.col_y1 + chartConditions.endRow;
  chartConditions.rangeAddress_y2 = chartConditions.col_y2 + chartConditions.startRow + ':' + chartConditions.col_y2 + chartConditions.endRow;
  // The name of the sheet to output the input data for Chart output
  chartConditions.outputSheetName = chartConditions.targetName + chartConditions.targetCost;
  // Chart title 
  chartConditions.title = chartConditions.outputSheetName;
  return chartConditions;
}
function executeCreateChart(target){
  var chartConditions = getCommonConstant(target);
  // グラフ出力シート作成
  // 臨床研究
  chartConditions.targetCost = chartConditions.constClinicalResearch;
  const chart_outputSheet = addSheetToEnd(chartConditions.chartSheetName); 
  chartConditions.col_y0 = 'L';
  chartConditions.col_y1 = 'N';
  chartConditions.col_y2 = 'Q';
  chartConditions.revenueItem = 'D/1000000';
  chartConditions.costItem = 'F*-1/1000000';
  chartConditions.profitItem = 'I/1000000';
  setChartRange(chartConditions);
  const inputSheet_clinical_research = createQuerySheet(chartConditions)
  createChart(chartConditions, inputSheet_clinical_research, chart_outputSheet, 1);
  // 経常
  chartConditions.targetCost = chartConditions.constOrdinary;
  chartConditions.col_y0 = 'M';
  chartConditions.col_y1 = 'O';
  chartConditions.col_y2 = 'P';
  chartConditions.revenueItem = 'E/1000000';
  chartConditions.costItem = 'G*-1/1000000';
  chartConditions.profitItem = 'H/1000000';
  setChartRange(chartConditions);
  const inputSheet_ordinary = createQuerySheet(chartConditions)
  createChart(chartConditions, inputSheet_ordinary, chart_outputSheet, 24);
}

