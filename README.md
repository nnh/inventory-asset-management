# inventory-asset-management
## 概要
資産管理の端末一覧とSINET接続依頼申請書を突き合わせて確認が必要な情報を抽出するスクリプトです。
## 事前設定（初回のみ）
- 出力用のスプレッドシートを作成してください。  
- function registerScriptProperty()内の下記箇所を設定して保存し、registerScriptPropertyを実行してください。
``````
  // SINET接続許可依頼書のスプレッドシートのURL
  const sinetUrl = 'https://docs.google.com/spreadsheets/d/...';
  // 資産管理のスプレッドシートのURL
  const shisankanriUrl = 'https://docs.google.com/spreadsheets/d/...';
  // 出力用スプレッドシートのURL
  const outputSheetUrl = 'https://docs.google.com/spreadsheets/d/...';
``````  
## 実行手順
getShisanAndSinetDataを実行してください。  
出力用スプレッドシートの一番左〜左から4番目のシートに下記の抽出結果が出力されます。
- 資産管理の機器一覧とSINET接続許可を突き合わせし、SINET接続許可に存在しない端末を抽出する。
- SINET接続申請内で重複しているコンピューター名を抽出する。
- 資産管理の機器一覧とSINET接続許可を突き合わせし、使用者名が異なる端末を抽出する。
- 資産管理の機器一覧とSINET接続許可を突き合わせし、機器一覧に存在しない端末を抽出する。
