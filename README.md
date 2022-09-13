# pl-chart
## 概要
GoogleSpreadsheetの'working'シートの内容から複合グラフを作成し、新規作成したGoogleSpreadsheetに出力します。  
## 使用手順
- グラフソースデータが入力されているGoogleSpreadsheetを開き、メニュー > 拡張機能 > Apps Scriptでスクリプトエディタを開きます。Google Apps Script GitHub アシスタント等を使用し、このリポジトリをpullしてください。pull後の初回実施時のみ、common.gsのregisterScriptPropertyを実行してください。  
- 'script_wk'シートのセルC7に、入力用PDFの取得元URLを記載してください。  
- 'script_wk'シートのセルC8に、入力用PDFを保存する先のGoogleドライブのURLを記載してください。   
- getPdfFiles.gsを開き、getPdfFilesMainを実行してください。  
- 保存されたPDFをダウンロードし、下記スクリプトを手順に沿って実行してください。  
  独立行政法人PL  
  https://gist.github.com/MarikoOhtsuka/54d571d79b4aefd5dec916159f444941  
- 上記スクリプト実行後に出力されたCSVファイルの内容をコピーし、'out'シートの末尾に追加してください。  
- 'script_wk'シートのセルC9に、出力ファイルを保存する先のGoogleドライブのURLを記載してください。   
- GoogleSpreadsheetのメニュー > PL表出力からグラフを出力します。処理の途中でポップアップメッセージが、臨床研究センター分は1回、臨床研究センター以外分は4回出力されます。詳細は「出力ファイルのアクセス許可作業について」の内容をご確認ください。    
  - 臨床研究センターをクリックすると、臨床研究センターセグメントの施設のグラフが出力されます。
  - 臨床研究センター以外をクリックすると、臨床研究センター以外のセグメントの施設のグラフが出力されます。
## 出力ファイルのアクセス許可作業について  
グラフ出力処理の途中で「出力ファイルを開き、アクセス許可を実施してからOKをクリックしてください。」というメッセージが出ます。  
'script_wk'シートのセルC8で指定したフォルダにあるGoogleSpreadsheetを開いてください。  
- 臨床研究センターの場合、「PL（臨床研究センター）」  
- 臨床研究センター以外の場合、「PL（臨床研究センター以外）」の1~4  

一番右のシートに下記のような情報が出ていたら、「アクセスを許可」をクリックしてください。  
![スクリーンショット 2022-09-13 15 10 38](https://user-images.githubusercontent.com/24307469/189823538-34838f24-5ba6-4428-a80e-a429f60b66d4.png)  
その後にポップアップのOKをクリックして、処理を続行してください。  
この作業を行わなかった場合、グラフの色等の設定が反映されなくなります。  
## 備考
GoogleSpreadsheetの制限のため、１ファイルにつき最高33施設までの情報を出力します。  
