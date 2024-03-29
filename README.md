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
1. DC入退室レポートは/Box/Datacenter/ISR/Attendance Management/YYYYMM/CardLogs YYYYMM.pdf  
VPN接続レポートは/Box/Datacenter/ISR/Attendance Management/YYYYMM/VPN Logs YYYYMM.pdf  
実行ファイルは\\aronas\ISR\SystemAssistant\monthlyOperations\vpnAndRoomAccessLogs\VPN CardLogs YYYYMM.xlsm  
に出力されます。  
1. 祝日が想定通りに出力されていない場合は、祝日入力シート（事務局用）.xlsxの更新を行って再度処理を実行してください。  
処理手順は「D013-4 SEアシスタントマニュアル（随時）」祝日マスタの更新を参照ください。  
## 仕様
### 抽出条件
下記マニュアルを参照ください。  
D013-2 SEアシスタントマニュアル（月次） - Google ドキュメント
https://docs.google.com/document/d/1w9B8Mk3RhhhBOY8wgcM_tpW3yeKvV8Qhy4F7QZKJV0I
### 入力ファイル
- \\aronas\Archives\Log\DC入退室\rawdata\YYYYMM （YYYYMMは対象年月） 配下にある拡張子が"csv"のファイル
- \\aronas\Archives\Log\VPN\rawdata 配下にある"access.log", "access.log- YYYYMMDD" （YYYYMMの範囲は対象年月の±１ヶ月） 
### 出力ファイル
- PDFファイル
    - DC入退室レポート  
        room_listシートの情報が出力される。  
    - VPN接続レポート  
        １ページ目にovertimeシート、それ以降にcheckシートの情報が出力される。  
- EXCELファイル  
    - overtime  
        抽出条件に該当したVPN接続ログが出力される。  
    - holiday  
        祝日入力シート（事務局用）.xlsx内の祝日情報が出力される。  
    - check  
        VPNの接続開始ログが出力される。  
    - summary_connected_from    
        ユーザー毎のVPN接続元IPアドレスが出力される。  
    - connected_from  
        作業用シート  
    - room_input  
        入力対象の全ての入退室情報が出力される。  
    - room_list  
        抽出条件に該当した入退室ログが出力される。  
    - vpn_input  
        入力対象の全てのVPN接続情報が出力される。  
### ディレクトリ構成とファイルの概要
```
.
├── README.md ... このReadMeです
├── programs
│   ├── VPN CardLogs template.xlsm ... プログラム本体です
│   ├── common.bas ... 共通処理モジュール
│   ├── iconv.sh ... ファイルの文字コード変換用スクリプト
│   ├── room_management.bas ... 入退室ログ出力処理モジュール
│   └── vpn_management.bas ... VPNログ出力処理モジュール
└── public_holiday
    └── 祝日入力シート（事務局用）.xlsx ... 祝日情報
```
### プログラム修正時の手順
1. 修正したモジュールをエクスポートしてください。
2. エクスプローラで右クリックし、「Git bash here」をクリックしてください。
3. 開いたGit bashで下記のコマンドを入力してください。
```
sh iconv.sh
```
4. エクスポートしたモジュールをテキストエディタなどで開き、文字コードがUTF-8であることを確認してからcommitし、pushしてください。


