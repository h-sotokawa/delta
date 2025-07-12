# QRコード中間ページシステム

## 概要
代替機のQRコードから共通フォームへ遷移するための中間ページシステムです。

## ファイル構成

### 1. qrcode.html（テスト用）
- ローカルテスト用のHTMLファイル
- Google Sitesに埋め込む場合のテンプレート
- **注意**: Google Sitesではiframe制限により動作しません

### 2. gas-qr-redirect-final.gs（本番用・推奨）
- Google Apps Script Web Appとして使用
- URLパラメータを確実に取得可能
- 自動リダイレクト機能付き

## 実装方法

### 推奨：Google Apps Script Web App

1. [Google Apps Script](https://script.google.com/)で新規プロジェクト作成
2. `gas-qr-redirect-final.gs`のコードをコピー
3. デプロイ:
   - デプロイ → 新しいデプロイ
   - 種類: ウェブアプリ
   - アクセス: 全員
4. 生成されたURLをQRコードに使用

### QRコードURL形式
```
https://script.google.com/macros/s/[SCRIPT_ID]/exec?id=[拠点管理番号]
```

### 拠点管理番号の形式
```
拠点_カテゴリ_モデル_製造番号_連番
例: OSA_desktop_OptiPlex_ABC123_001
```

### カテゴリ別振り分け
- **端末用フォーム**: desktop, laptop, server
- **プリンタ用フォーム**: printer, other

## メンテナンス

### フォームURL変更時
1. GASプロジェクトを開く
2. フォームURLを更新
3. デプロイ → デプロイを管理 → 編集 → 新バージョン
4. **URLは変わりません**

## 注意事項
- Google Sitesの埋め込みHTMLではURLパラメータ取得に制限があります
- GAS Web Appを使用することで確実に動作します
- デプロイURLは「デプロイを管理」から更新することで固定できます