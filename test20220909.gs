function issue_5_(){
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('script_wk');
  inputSheet.getRange('A7').setValue('入力用PDFのURLをC列に入力');
  inputSheet.getRange('B7').setValue('=REGEXREPLACE(REGEXEXTRACT(C7, "https://.*/"), "/$", "")');
  inputSheet.getRange('A8').setValue('PDF格納フォルダのURLをC列に入力');
  inputSheet.getRange('B8').setValue('=REGEXREPLACE(C8, "https://drive.google.com/drive/folders/", "")');
}
function issue_4_(){
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('script_wk');
  inputSheet.getRange('A9').setValue('出力ファイル格納フォルダのURLをC列に入力');
  inputSheet.getRange('B9').setValue('=REGEXREPLACE(C9, "https://drive.google.com/drive/folders/", "")');
  inputSheet.getRange('A2:C6').clearContent();
  inputSheet.getRange('A2').setValue('未使用');
  inputSheet.getRange('A3').setValue('未使用');
  inputSheet.getRange('A4').setValue('未使用');
  inputSheet.getRange('A5').setValue('未使用');
  inputSheet.getRange('A6').setValue('未使用');
}
function test20220909main(){
  issue_4_();
  issue_5_();
}