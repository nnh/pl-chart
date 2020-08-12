// 
/**
* Creating the source data for the chart
* @param {object} chartConditions 
* @return {sheet} The sheet object created 
*/
function createQuerySheet(chartConditions){
  const inputSheetName = PropertiesService.getScriptProperties().getProperty('inputSheetName');
  // Create output sheet
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
  // Formatting Chart Data
  targetSheet.getRange('D:Q').setNumberFormat('#,##0');
  // Hide this sheet
  targetSheet.hideSheet();
  return targetSheet;
}
