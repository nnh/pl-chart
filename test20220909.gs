function issue_5(){
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('script_wk');
  inputSheet.getRange('A7').setValue('入力用PDFのURLをC列に入力');
  inputSheet.getRange('B7').setValue('=REGEXREPLACE(REGEXEXTRACT(C7, "https://.*/"), "/$", "")');
  inputSheet.getRange('A8').setValue('PDF格納フォルダのURLをC列に入力');
  inputSheet.getRange('B8').setValue('=REGEXREPLACE(C8, "https://drive.google.com/drive/folders/", "")');
}