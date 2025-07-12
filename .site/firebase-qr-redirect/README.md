# Firebase Hosting によるQRコード中間ページ

## メリット

1. **超高速**: 静的ホスティングなので即座にリダイレクト
2. **無料枠で十分**: 
   - 10GB/月のホスティング容量
   - 360MB/日のデータ転送量
   - カスタムドメイン対応
3. **高可用性**: Googleのグローバルインフラ
4. **HTTPS対応**: 自動的にSSL証明書付き

## セットアップ手順

### 1. Firebaseプロジェクトの作成

```bash
# Firebase CLIのインストール（未インストールの場合）
npm install -g firebase-tools

# Firebaseにログイン
firebase login

# プロジェクトの初期化
firebase init hosting
```

### 2. プロジェクト設定

初期化時の選択：
- **Project**: 新規作成または既存プロジェクトを選択
- **Public directory**: `.` （カレントディレクトリ）
- **Single-page app**: Yes
- **GitHub Actions**: No（必要に応じて）

### 3. デプロイ

```bash
# デプロイ実行
firebase deploy --only hosting
```

デプロイ後、以下のようなURLが発行されます：
```
https://your-project-id.web.app
https://your-project-id.firebaseapp.com
```

### 4. カスタムドメイン設定（オプション）

Firebaseコンソールから：
1. Hosting → カスタムドメインを追加
2. ドメインを入力（例：`qr.your-company.com`）
3. DNS設定を行う

## 使用方法

QRコードには以下のURLを設定：
```
https://your-project-id.web.app/?id=拠点管理番号
```

例：
```
https://your-project-id.web.app/?id=OSAKA_server_ThinkPad_ABC123_001
```

## 動作フロー

1. QRコードを読み取る
2. Firebase Hostingのページにアクセス
3. JavaScriptが拠点管理番号を解析
4. 適切なGoogle Formsへ即座にリダイレクト

## コスト

無料枠の範囲：
- **ストレージ**: 10GB/月
- **データ転送**: 360MB/日
- **ファイル数**: 無制限

通常の使用では無料枠で十分対応可能です。

## 高度な機能（オプション）

### ログ収集
Firebase Analyticsと連携してアクセスログを収集可能

### A/Bテスト
Firebase A/B Testingでリダイレクト速度の最適化

### エラー監視
Firebase Crashlyticsでエラーを監視