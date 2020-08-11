// 印刷後に実行する
// site〜outまでを表示する
function visibleWorkings(){
  visibleHiddenWorkSheets(true);
}
// site〜outまでを非表示にする
function hiddenWorkings(){
  visibleHiddenWorkSheets(false);
}
/**
* Show/Hide Sheets
* @param {boolean}  
* @return {string} The names of the clinical research center in a one-dimensional array
*/
function visibleHiddenWorkSheets(condition){
  class visibleHiddenSheets{
    constructor(condition){
      this.condition = condition;
      const targetSheets = getSheetsNonTargetPrinting();
      setVisibleHiddenSheets(targetSheets, this.condition);
    }
  }
  return new visibleHiddenSheets(condition);  
}
