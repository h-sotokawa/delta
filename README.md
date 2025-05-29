# 管理システム (Google Apps Script Web App)

## 概要
Google Apps Script (GAS) を使用した代替機管理システムです。スプレッドシートのデータを閲覧・編集できるWebアプリケーションとして動作します。

## 主な機能
- **ダッシュボード**: 各機能へのナビゲーション
- **スプレッドシートビューアー**: 各拠点のデータを閲覧・編集
- **ドラッグ&ドロップ**: テーブルの列を自由に並び替え
- **ステータス管理**: 貸出機器のステータスをリアルタイムで更新

## ファイル構造
```
├── Index.html                    # メインのHTMLファイル
├── Code.gs                       # Google Apps Scriptのサーバーサイドコード
├── styles.html                   # グローバルスタイルシート
├── dashboard-styles.html         # ダッシュボード専用スタイル
├── spreadsheet-styles.html       # スプレッドシート専用スタイル
├── navigation.html               # ナビゲーションバーコンポーネント
├── dashboard.html                # ダッシュボードページ
├── spreadsheet.html              # スプレッドシートビューアーページ
├── coming-soon.html              # 準備中ページテンプレート
├── main.html                     # メインJavaScriptコード
├── spreadsheet-functions.html    # スプレッドシート関連JavaScript
├── drag-drop-functions.html      # ドラッグ&ドロップ関連JavaScript
└── README.md                     # このファイル
```

## セットアップ方法
1. Google Apps Scriptプロジェクトを作成
2. すべてのファイルをプロジェクトにアップロード
3. スクリプトプロパティに各拠点のスプレッドシートIDを設定
   - `SPREADSHEET_ID_SOURCE_OSAKA_DESKTOP`
   - `SPREADSHEET_ID_SOURCE_OSAKA_LAPTOP`
   - `SPREADSHEET_ID_SOURCE_KOBE`
   - `SPREADSHEET_ID_SOURCE_HIMEJI`
   - `SPREADSHEET_ID_SOURCE_OSAKA_PRINTER`
   - `SPREADSHEET_ID_SOURCE_HYOGO_PRINTER`
4. Webアプリとしてデプロイ

## 開発時の注意点
- デバッグモードは `main.html` の `DEBUG_MODE` 変数で制御
- 本番環境では `DEBUG_MODE = false` に設定してください
- 新しいページを追加する場合は、`navigation.html` にリンクを追加してください

## 今後の拡張予定
- フォーム作成機能
- レポート機能
- データインポート/エクスポート機能
- システム設定機能

## ライセンス
内部利用のみ