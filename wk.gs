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