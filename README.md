# vpn-room-log-management
## 概要
VPNログ、入退室ログを整形するExcel用マクロです。
## ファイルのダウンロード
Zipファイルをダウンロードします。  
![download](https://user-images.githubusercontent.com/24307469/126087131-7fd36292-2220-4a86-85ca-53925e1b74e0.png).  
ダウンロードしたファイルを右クリックして「すべて展開」します。
## 実行方法
1. 「VPN CardLogs template.xlsm」の名前を「VPN CardLogs YYYYMM.xlsm」（YYYYMMは対象年月）に変更して開きます。  
1. Sheet1シートの実行ボタンをクリックすると処理が実行されます。  
1. デフォルトでは処理日の前月のフォルダを処理対象とします。  
（例：2020/12/11に処理を実行した場合、"202011"フォルダを対象とする。）  
それ以外の月を処理する場合は、common module内で下記の記載を検索し、""に対象のフォルダ名を記入してから実行してください。  
`Const yyyymm As String = ""` 
1. DC入退室レポートは\\aronas\Archives\Log\DC入退室\CardLogs YYYYMM.pdf  
VPN接続レポートは\\aronas\Archives\Log\VPN\VPN Logs YYYYMM.pdf  
に出力されます。  
## 仕様
### 抽出条件
土、日、祝日または22:00:00～翌4:59:59までに入退室、VPNアクセスがあった場合  
### 入出力ファイル
SEアシスタントマニュアル（月次）に記載

