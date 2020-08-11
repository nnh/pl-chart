// 出力シート削除
function deleteSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var targetSheets = sheets.filter(function(x){
    if (x.getName().slice(-6) == '（臨床研究）' || x.getName().slice(-4) == '（全体）' || x.getName().substr(0, 2) == '平成'){
      return x
    }
  });
  targetSheets.map(function(x){
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(x);
  })
  const targetSheetNames = getTargetFacilitiesNames();
  targetSheetNames.map(function(x){
    deleteTargetSheet(x);
  });
}
// ゴミ削除用
function deleteSheets2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var targetSheets = sheets.filter(function(x){
    if (x.getName().substr(0, 3) == 'シート'){
      return x
    }
  });
  targetSheets.map(function(x){
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(x);
  })
}