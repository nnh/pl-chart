/**
* Creating the source data for the chart
* @param {object} chartConditions 
* @return {sheet} The sheet object created 
*/
function createQuerySheet(chartConditions){
  const inputSSId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const inputSheetName = PropertiesService.getScriptProperties().getProperty('inputSheetName');
  // Create output sheet
  const targetSheet = addSheetToEnd(chartConditions.outputSheetName, chartConditions.ss); 
  // Output the source data of a chart
  var strQuery = '=query(importrange(' + '"' + inputSSId + '", ' + '"' + inputSheetName + '!A:K"), "' +
                        'select ' + "'" + 'unused_A' + "'," +
                                      chartConditions.strB + ',' +
                                      'Col3, sum(Col4), sum(Col5), sum(Col6), sum(Col7), sum(Col8), sum(Col9), ' + 
                                      "'" + 'unused_J' + "'," +
                                      'Col11, ' +
                                      chartConditions.selectItems_D + ',' + 
                                      chartConditions.selectItems_E + ',' + 
                                      chartConditions.selectItems_F + ',' + 
                                      chartConditions.selectItems_G + ',' + 
                                      chartConditions.selectItems_H + ',' + chartConditions.selectItems_I  + ' ' +
                        'where ' + chartConditions.dataCol + ' = ' + chartConditions.condition + 
                                   " and " + chartConditions.xCol + ' is not null ' + 
                                   chartConditions.ymdCondition + ' ' + 
                        'group by ' + chartConditions.groupBy + ' ' +
                        'order by ' + chartConditions.orderbyCol + ' ' + 
                        'label ' + chartConditions.revenueItem + " '" + chartConditions.labelRevenue + "', " + 
                                   chartConditions.costItem + " '" + chartConditions.labelCost + "'," +  
                                   chartConditions.profitItem + " '" + chartConditions.labelProfit + "'" + '")';
  targetSheet.getRange(1, 1).setFormula(strQuery);
  // Formatting Chart Data
  targetSheet.getRange('D:Q').setNumberFormat('#,##0');
  // Hide this sheet
  targetSheet.hideSheet();
  return targetSheet;
}
