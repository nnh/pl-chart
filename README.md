# pl-chart
## 概要
GoogleSpreadsheetの'working'シートの内容から複合グラフを作成し、指定したGoogleSpreadsheetに出力します。  
## 使用手順
・出力するGoogleSpreadsheetを任意のフォルダに作成し、ファイルのIDを入力用GoogleSpreadsheetの'script_wk'シートのB列に記載してください。  
・スクリプトエディタでGoogle Apps Script GitHub アシスタントを使用し、グラフソースデータが入力されているGoogleSpreadsheetにこのリポジトリをpullしてください。  
・初回実施時のみ、common.gsのregisterScriptPropertyを実行してください。  
・pl_clinicalResearchCenter.gsのexecCreateChartMainを実行すると臨床研究センターのグラフが出力されます。  
・pl_others.gsのexecCreateChartOthersMainを実行すると臨床研究センター以外のグラフが出力されます。  
