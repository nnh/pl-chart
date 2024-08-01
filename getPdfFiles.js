/**
 * The PDF of the graph input source is retrieved from the specified URL.
 * @param none.
 * @return none.
 */
function getPdfFilesMain(){
  const getScriptWkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('script_wk');
  if (!getScriptWkSheet){
    return;
  };
  let getPdfArg = {};
  getPdfArg.savePdfFolderId = getScriptWkSheet.getRange('B8').getValue();
  if (getPdfArg.savePdfFolderId === ''){
    return;
  };
  const getPdfParentUrl = getScriptWkSheet.getRange('C7').getValue();
  if (getPdfParentUrl === ''){
    return;
  };
  getPdfArg.fqdn = getScriptWkSheet.getRange('B7').getValue();
  if (getPdfArg.fqdn === ''){
    return;
  };
  const getPdf = new GetPdf(getPdfArg);
  const parentContents = getPdf.getParent(getPdfParentUrl);
  const urls = getPdf.getPdfUrls(parentContents);
  urls.forEach(x => {
    getPdf.getPdf(x);
  });
};
class GetPdf{
  constructor(getPdfArg){
    this.moveToFolderId = getPdfArg.savePdfFolderId;
    this.fqdn = getPdfArg.fqdn;
  };
  getPdf(targetUrl){
    const pdf = UrlFetchApp.fetch(targetUrl).getAs('application/pdf');
    const folder = DriveApp.getFolderById(this.moveToFolderId);
    const createPdf = DriveApp.createFile(pdf);
    createPdf.moveTo(folder);
    Utilities.sleep(20);
  };
  getParent(targetUrl){
    const res = UrlFetchApp.fetch(targetUrl);
    const resContents = res.getContentText();
    return resContents;
  };
  getPdfUrls(targetContents){
    const files = targetContents.match(/\/files\/[0-9]*\.pdf/g);
    const res = files.map(x => this.fqdn + x); 
    return res;
  };
};
