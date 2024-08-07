# pl-chart

## 概要

GoogleSpreadsheet の'working'シートの内容から複合グラフを作成し、新規作成した GoogleSpreadsheet に出力します。

## 使用手順

- グラフソースデータが入力されている GoogleSpreadsheet を開き、メニュー > 拡張機能 > Apps Script でスクリプトエディタを開きます。Google Apps Script GitHub アシスタント等を使用し、このリポジトリを pull してください。pull 後の初回実施時のみ、common.gs の registerScriptProperty を実行してください。
- 'script_wk'シートのセル C7 に、入力用 PDF の取得元 URL を記載してください。
- 'script_wk'シートのセル C8 に、入力用 PDF を保存する先の Google ドライブの URL を記載してください。
- getPdfFiles.gs を開き、getPdfFilesMain を実行してください。
- 保存された PDF をダウンロードし、下記スクリプトを手順に沿って実行してください。  
  独立行政法人 PL  
  https://gist.github.com/MarikoOhtsuka/54d571d79b4aefd5dec916159f444941
- 上記スクリプト実行後に出力された CSV ファイルの内容をコピーし、'out'シートの末尾に追加してください。
- 'working'シートに入っている数式を'out'シートに追加した行数分コピーして追加してください。
- 'script_wk'シートのセル C9 に、出力ファイルを保存する先の Google ドライブの URL を記載してください。
- GoogleSpreadsheet のメニュー > PL 表出力からグラフを出力します。処理の途中でポップアップメッセージが、臨床研究センター分は 1 回、臨床研究センター以外分は 4 回出力されます。詳細は「出力ファイルのアクセス許可作業について」の内容をご確認ください。
  - 臨床研究センターをクリックすると、臨床研究センターセグメントの施設のグラフが出力されます。
  - 臨床研究センター以外をクリックすると、臨床研究センター以外のセグメントの施設のグラフが出力されます。

## 出力ファイルのアクセス許可作業について

グラフ出力処理の途中で「出力ファイルを開き、アクセス許可を実施してから OK をクリックしてください。」というメッセージが出ます。  
'script_wk'シートのセル C8 で指定したフォルダにある GoogleSpreadsheet を開いてください。

- 臨床研究センターの場合、「PL（臨床研究センター）」
- 臨床研究センター以外の場合、「PL（臨床研究センター以外）」の 1~4
  一番右のシートに下記のような情報が出ていたら、「アクセスを許可」をクリックしてください。  
  ![スクリーンショット 2022-09-13 15 10 38](https://user-images.githubusercontent.com/24307469/189823538-34838f24-5ba6-4428-a80e-a429f60b66d4.png)  
  その後にポップアップの OK をクリックして、処理を続行してください。  
  この作業を行わなかった場合、グラフの色等の設定が反映されなくなります。

## グラフの書式設定について

- GoogleSpreadsheet のメニュー > PL 表出力 > グラフ整形を実行すると、'script_wk'シートのセル C9 で指定したフォルダ配下のスプレッドシートにあるグラフの書式を全て更新します。

## 備考

GoogleSpreadsheet の制限のため、１ファイルにつき最高 33 施設までの情報を出力します。
