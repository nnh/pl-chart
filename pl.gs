// https://developers.google.com/chart/interactive/docs/gallery/combochart

// グラフを作成する
function execCreateChart(){
  // 出力シートの削除
  deleteSheets();
  var target = {};
  //　施設毎シート
  const targetFacilities = getTargetFacility();
  targetFacilities.map(function(targetFacility){
    target.name = targetFacility;
    target.dataCol = 'B';     // 施設名毎に出力
    target.xCol = 'C';        // 横軸：年度
    target.orderbyCol = 'C';  // 横軸の出力順は年度でソートする
    target.ymdCondition = '';
    executeCreateChart(target);
  });
  //　年度毎シート
  const targetYears = getTargetYears();
  targetYears.map(function(targetYear){
    target.name = targetYear;
    target.dataCol = 'C';     // 年度毎に出力
    target.xCol = 'B';        // 横軸：施設名
    target.orderbyCol = 'K';  // 横軸の出力順は施設IDでソートする
    target.ymdCondition = "and J = '臨床研究センター' ";  // セグメントが臨床研究センターの行のみ対象とする
    executeCreateChart(target);
  });
  // シートの並べ替え
  sortSheets();
  // 印刷のためsiteidH31.3〜sampleまでを非表示にする
  // 「近畿中央胸部疾患センター」シートは不要なので非表示にする
  var hiddenSheet = getTargetSheets(['近畿中央胸部疾患センター']);
  const nonTargetPrintingSheets = getSheetsNonTargetPrinting();
  hiddenSheet = hiddenSheet.concat(nonTargetPrintingSheets);
  setVisibleHiddenSheets(hiddenSheet, false);
}
// シート並び替え
function sortSheets(){
  const targetFacilities = getTargetFacilitiesNames();
  const targetYears = getTargetYears();
  const targetSheetsName = targetFacilities.concat(targetYears);
  // siteidH31.3〜sampleまでを左端にする
  const sheetsNonTargetPrinting = getSheetsNonTargetPrinting();
  // グラフ
  const sheetsTargetPrinting = getTargetSheets(targetSheetsName);
  const moveSheets = sheetsNonTargetPrinting.concat(sheetsTargetPrinting);
  moveSheets.map(function(x, idx){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    x.activate();
    ss.moveActiveSheet(idx + 1);
  });
}
// シート名の配列からシートオブジェクトを取得する
function getTargetSheets(sheetNames){
  var targetSheets = sheetNames.map(function(x){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheetByName(x);
  });
  // そのシート名のシートが存在しなければ配列から削除
  targetSheets = targetSheets.filter(Boolean);
  return targetSheets;
}
// 印刷対象外シートを取得する
function getSheetsNonTargetPrinting(){
  const constTarget = ['仕様書', 'site', 'siteidH31.3', 'siteChangeInfo', 'segment', 'working', 'out'];
  var targetSheets = getTargetSheets(constTarget);
  return targetSheets;
}
// シートの表示・非表示
// 引数はシートオブジェクト
// true: 表示
// false: 非表示
function setVisibleHiddenSheets(targetSheets, visibleFlg){
  targetSheets.map(function(x){
    if (visibleFlg){
      x.showSheet();
    } else {
      x.hideSheet();
    }
  });
}
// シート削除
function deleteTargetSheet(sheetName){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(sheetName)
  if (targetSheet != null){
    ss.deleteSheet(targetSheet);
  }
}
// 再末尾にシート作成
function addSheetToEnd(sheetName){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempIdx = ss.getNumSheets();
  const targetSheet = ss.insertSheet(sheetName, parseInt(tempIdx)); 
  return targetSheet;  
}
// グラフの元データを作成する
function createQuerySheet(chartConditions){
  // 入力シート指定
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheetName = 'working';
  const inputSheet = ss.getSheetByName(inputSheetName);
  // 出力シート作成
  const targetSheet = addSheetToEnd(chartConditions.outputSheetName); 
  // グラフ作成元データの出力
  var strQuery = '=query(' + inputSheetName + '!A:K, "select A, B, C, D, E, F, G, H, I, J, K, ' + chartConditions.selectItems_D + ',' + chartConditions.selectItems_E + ',' + chartConditions.selectItems_F + ',' + 
                                                                  chartConditions.selectItems_G + ',' + chartConditions.selectItems_H + ',' + chartConditions.selectItems_I  +
                                                          ' where ' + chartConditions.dataCol + ' = ' + "'" + chartConditions.condition + 
                                                                   "' and " + chartConditions.xCol + ' is not null ' + chartConditions.ymdCondition + 
                                                          ' order by ' + chartConditions.orderbyCol + 
                                                          ' label ' + chartConditions.revenueItem + " '" + chartConditions.labelRevenue + "', " + 
                                                                      chartConditions.costItem + " '" + chartConditions.labelCost + "'," +  
                                                                      chartConditions.profitItem + " '" + chartConditions.labelProfit + "'" +  
                                                            '")';
  // 近畿中央呼吸器センターと近畿中央胸部疾患センター
  if (chartConditions.targetName == '近畿中央呼吸器センター'){
    strQuery = '=query(' + inputSheetName + '!A:K, "select A, B, C, D, E, F, G, H, I, J, K, ' + chartConditions.selectItems_D + ',' + chartConditions.selectItems_E + ',' + chartConditions.selectItems_F + ',' + 
                                                                  chartConditions.selectItems_G + ',' + chartConditions.selectItems_H + ',' + chartConditions.selectItems_I  +
                                                          ' where ' + " B = '近畿中央呼吸器センター' or B = '近畿中央胸部疾患センター' " + 
                                                                   " and " + chartConditions.xCol + ' is not null ' + chartConditions.ymdCondition + 
                                                          ' order by ' + chartConditions.orderbyCol + 
                                                          ' label ' + chartConditions.revenueItem + " '" + chartConditions.labelRevenue + "', " + 
                                                                      chartConditions.costItem + " '" + chartConditions.labelCost + "'," +  
                                                                      chartConditions.profitItem + " '" + chartConditions.labelProfit + "'" +  
                                                            '")';
  }
  
  targetSheet.getRange(1, 1).setFormula(strQuery);
  // グラフデータの書式の設定
  targetSheet.getRange('D:Q').setNumberFormat('#,##0');
  // 非表示にする
  targetSheet.hideSheet();
  return targetSheet;
}
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

function getCommonConstant(){
  const items_unit = '1000000';  // 百万単位で出力する
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
  return chartConditions;
}
function setChartRange(chartConditions){
  // データ範囲を指定
  chartConditions.rangeAddress_y0 = chartConditions.col_y0 + chartConditions.startRow + ':' + chartConditions.col_y0 + chartConditions.endRow;
  chartConditions.rangeAddress_y1 = chartConditions.col_y1 + chartConditions.startRow + ':' + chartConditions.col_y1 + chartConditions.endRow;
  chartConditions.rangeAddress_y2 = chartConditions.col_y2 + chartConditions.startRow + ':' + chartConditions.col_y2 + chartConditions.endRow;
  // グラフのタイトルと出力シート名もついでにここで  
  chartConditions.outputSheetName = chartConditions.targetName + chartConditions.targetCost;
  chartConditions.title = chartConditions.outputSheetName;
  return chartConditions;
}
function executeCreateChart(target){
  var chartConditions = getCommonConstant();
  // グラフ出力シート作成
  // 臨床研究
  chartConditions.targetCost = chartConditions.constClinicalResearch;
  chartConditions.targetName = target.name;
  chartConditions.chartSheetName = chartConditions.targetName;  // グラフを出力するシート名
  const chart_outputSheet = addSheetToEnd(chartConditions.chartSheetName); 
  chartConditions.xCol = target.xCol;
  chartConditions.rangeAddress_x = chartConditions.xCol + chartConditions.startRow + ':' + chartConditions.xCol + chartConditions.endRow;
  chartConditions.col_y0 = 'L';
  chartConditions.col_y1 = 'N';
  chartConditions.col_y2 = 'Q';
  chartConditions.revenueItem = 'D/1000000';
  chartConditions.costItem = 'F*-1/1000000';
  chartConditions.profitItem = 'I/1000000';
  setChartRange(chartConditions);
  chartConditions.condition = chartConditions.targetName;
  chartConditions.dataCol = target.dataCol;
  chartConditions.orderbyCol = target.orderbyCol;
  chartConditions.ymdCondition = target.ymdCondition;
  chartConditions.max_y2 = '';
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

// 年度を取得
function getTargetYears(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workingSheet = ss.getSheetByName('working');
  var years = workingSheet.getRange('C:C').getValues();
    // 配列の次元を落とす
  years = years.reduce(function(x, item){
    x.push(...item);
    return x;
  }, []);
  // 重複、null、見出しを削除
  years = years.filter(function(x){
    return x != '年度';
  })
  years = Array.from(new Set(years))
  years = years.filter(Boolean);
  return years;
}
