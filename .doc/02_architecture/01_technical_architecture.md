# 技術アーキテクチャ設計書

## 1. アーキテクチャ概要

### 1.1 システム構成

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   Client Side   │    │   Server Side   │    │   Data Layer    │
│                 │    │                 │    │                 │
│ HTML5 + CSS3    │◄──►│ Google Apps     │◄──►│ Google Sheets   │
│ JavaScript      │    │ Script          │    │ Properties      │
│ (ES6+)          │    │                 │    │ Service         │
└─────────────────┘    └─────────────────┘    └─────────────────┘
```

### 1.2 技術スタック

- **フロントエンド**: HTML5、CSS3、Vanilla JavaScript (ES6+)
- **バックエンド**: Google Apps Script
- **データベース**: Google Sheets + PropertiesService
- **認証**: Google OAuth 2.0
- **デプロイ**: Google Apps Script Web App
- **UI Framework**: Font Awesome (アイコン)、Custom CSS Grid

## 2. コンポーネント設計

### 2.1 フロントエンドコンポーネント

#### 2.1.1 ページコンポーネント

```
src/
├── Index.html                 # メインエントリーポイント
├── dashboard.html             # ダッシュボード画面
├── spreadsheet.html           # スプレッドシートビューアー
├── url-generator.html         # URL生成画面
├── model-master.html          # 機種マスタ管理
├── location-master.html       # 拠点マスタ管理
├── settings.html              # システム設定
└── navigation.html            # 共通ナビゲーション
```

#### 2.1.2 スタイルコンポーネント

```
styles/
├── styles.html                # 共通スタイル
├── dashboard-styles.html      # ダッシュボード専用
├── spreadsheet-styles.html    # スプレッドシート専用
├── url-generator-styles.html  # URL生成専用
├── model-master-styles.html   # 機種マスタ専用
├── location-master-styles.html # 拠点マスタ専用
└── settings-styles.html       # 設定専用
```

#### 2.1.3 機能コンポーネント

```
functions/
├── main.html                      # メイン制御・ページ遷移
├── spreadsheet-functions.html     # スプレッドシート機能
├── url-generator-functions.html   # URL生成機能
├── model-master-functions.html    # 機種マスタ機能
├── location-master-functions.html # 拠点マスタ機能
├── settings-functions.html        # 設定機能
└── drag-drop-functions.html       # ドラッグ&ドロップ機能
```

### 2.2 バックエンドコンポーネント

#### 2.2.1 メインコンポーネント

```
Code.gs
├── doGet()                    # エントリーポイント
├── include()                  # HTMLインクルード機能
├── スプレッドシート操作関数     # データCRUD処理
├── マスタ管理関数             # マスタデータ管理
├── URL生成関数               # URL生成ロジック
├── 設定管理関数              # システム設定管理
├── ユーティリティ関数         # 共通処理
└── ログ・パフォーマンス管理    # 監視・ログ機能
```

#### 2.2.2 専用モジュール

```
url-generation.gs             # URL生成専用ロジック
├── generateCommonFormUrl()   # 共通フォームURL生成
├── addLocationNumberParameter() # パラメータ追加
├── saveUrlToMaster()         # マスタ保存処理
└── QR関連関数                # QRコード処理
```

## 3. データアーキテクチャ

### 3.1 データストレージ構成

#### 3.1.1 Google Sheets 構成

```
メインスプレッドシート
├── ステータス収集シート（フォーム回答受信）
│   ├── 端末ステータス収集       # 端末フォームの直接送信先
│   ├── プリンタステータス収集   # プリンタフォームの直接送信先
│   └── その他ステータス収集     # その他フォームの直接送信先
│
├── マスタシート（データ管理）
│   ├── 拠点マスタ              # 拠点情報管理
│   ├── 機種マスタ              # 機種情報管理
│   ├── データタイプマスタ      # データタイプ管理
│   ├── 端末マスタ              # 端末機器管理（収集シートから自動取得）
│   ├── プリンタマスタ          # プリンタ機器管理（収集シートから自動取得）
│   └── その他マスタ            # その他機器管理（収集シートから自動取得）
```

#### 3.1.2 PropertiesService 構成

```
システム設定
├── SPREADSHEET_ID_MAIN  # メインスプレッドシートID
├── TERMINAL_COMMON_FORM_URL    # 端末用フォームURL
├── PRINTER_COMMON_FORM_URL     # プリンタ用フォームURL
├── QR_REDIRECT_URL            # QR中間ページURL
├── DEBUG_MODE                  # デバッグモード（true/false）
│
├── ログ通知設定（debugMode=trueの場合のみ）
│   ├── ERROR_NOTIFICATION_EMAIL    # エラー通知メールアドレス
│   └── ALERT_NOTIFICATION_EMAIL    # アラート通知メールアドレス
│
└── ステータス変更通知設定
    └── STATUS_CHANGE_NOTIFICATION_ENABLED  # 全体の有効/無効
        # 拠点別の通知ON/OFFは拠点マスタで管理
```

### 3.2 データフロー設計

#### 3.2.1 フォームデータフローと通知

```
[Google Forms回答]
      ↓
[ステータス収集シート]  ←─── onFormSubmitトリガー設定
      ↓                            ↓
[自動データ連携]              [ステータス変更検知]
      ↓                            ↓
[マスタシート更新]            [通知メール送信]
（QUERY/数式で自動）          （条件に応じて）

※ 重要: 通知はフォーム送信時のみ。スプレッドシート直接編集では通知なし
```

#### 3.2.2 URL 生成フロー

```
User Input → Form Validation → LocationNumber Generation
    ↓
Settings Validation → URL Generation → Master Data Save
    ↓
Response to Client → UI Update → Success/Error Display
```

#### 3.2.3 マスタ管理フロー

```
User Action → Form Submission → Data Validation
    ↓
Sheets API Call → Data Update → DateTime Recording
    ↓
Response Handling → UI Refresh → Notification Display
```

#### 3.2.4 データ取得フロー

```
Page Load → Background Preload → Cache Check
    ↓
Sheets Data Fetch → Data Processing → Client Update
    ↓
Real-time Sync → Performance Logging → Error Handling
```

#### 3.2.5 フォームデータフローと通知

```
Google Forms 回答送信
    ↓
ステータス収集シート（端末/プリンタ/その他）
    ↓
onFormSubmitトリガー発火
    ↓
ステータス変更検知処理
    ├─ 拠点管理番号から拠点ID抽出
    ├─ 前回ステータスとの比較
    └─ 変更有無の判定
    ↓
通知条件確認
    ├─ グローバル通知設定（PropertiesService）
    └─ 拠点別通知設定（拠点マスタ）
    ↓
通知送信（条件を満たす場合）
    ├─ メール作成・送信
    └─ 送信ログ記録
    ↓
統合ビューリアルタイム更新
    ├─ カテゴリ判定（端末系 or プリンタ・その他系）
    ├─ マスタデータ取得・結合
    ├─ 計算フィールド生成
    ├─ 統合ビューシート部分更新
    └─ 検索インデックス更新
```

### 3.2.1 統合ビュー更新アーキテクチャ

#### リアルタイム更新システム

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   Data Source   │    │  Update Engine  │    │ Integrated View │
│                 │    │                 │    │                 │
│ ○ 端末マスタ    │◄──►│ ○ onFormSubmit │◄──►│ ○ 端末系統合    │
│ ○ プリンタマスタ │    │ ○ onChange     │    │ ○ プリンタ系統合 │
│ ○ その他マスタ   │    │ ○ timeBased    │    │ ○ 検索インデックス│
│ ○ 拠点マスタ    │    │                 │    │                 │
│ ○ 機種マスタ    │    │ ○ 部分更新処理  │    │ ○ 計算フィールド │
│ ○ 収集シート    │    │ ○ エラー処理   │    │ ○ 参照データ    │
└─────────────────┘    └─────────────────┘    └─────────────────┘
```

#### 更新トリガー設定

| トリガー | 実行タイミング | 処理対象 | 更新範囲 |
|----------|----------------|----------|----------|
| onFormSubmit | フォーム送信時 | 該当機器のみ | 部分更新 |
| onChange | マスタシート変更時 | 関連データのみ | 部分更新 |
| timeBased | 深夜2:00 | 全データ | 完全再構築 |

#### パフォーマンス最適化

- **部分更新**: 変更された機器のみ処理
- **並列処理**: 端末系・プリンタ系の同時更新
- **インデックス活用**: 高速検索による行特定
- **キャッシュ活用**: マスタデータの一時保存

### 3.3 日時処理仕様

#### 3.3.1 基本仕様

- **タイムゾーン**: Asia/Tokyo (JST) で統一
- **保存形式**: Google Apps Script のデフォルト形式で保存
- **表示形式**: `yyyy/MM/dd HH:mm:ss`
- **データ型**: Google Apps Script のデフォルト日時型として処理

#### 3.3.2 バックエンド実装（Google Apps Script）

##### 日時生成ロジック

- `Utilities.formatDate()` を使用した現在日時取得
- 日時フォーマット: `yyyy/MM/dd HH:mm:ss`
- カスタム日時フォーマット機能

##### Google Sheets 保存処理

- データ配列に作成日時、更新日時を追加
- シートの最終行に挿入
- 日時は Google Apps Script のデフォルト形式で保存
- 更新時は更新日時のみを更新

##### データ取得処理

- シートからデータを配列として取得
- ヘッダー行とデータ行を分離
- 日時フィールドは Google Apps Script のデフォルト形式として保持

#### 3.3.3 フロントエンド実装（JavaScript）

##### 日時パース機能

- 日時文字列を Date オブジェクトに変換
- `yyyy/MM/dd HH:mm:ss` 形式の正規表現パース
- パースエラー時は null を返す

##### 日時フォーマット機能

- Date オブジェクトを文字列に変換
- 既存の文字列データはそのまま表示
- ゼロパディングで一貫したフォーマット
- エラー時の適切なハンドリング

##### 相対時間表示機能

- 「1 時間前」「昨日」などの人間的な表示
- 現在時刻との差分計算
- 表示長が長い場合の省略表示

##### UI での使用方法

- テーブルの更新日時表示
- タイムスタンプのリアルタイム更新
- ソート機能との組み合わせ

#### 3.3.4 データ処理層での日時ハンドリング

##### バリデーション機能

- 日時形式の正当性チェック
- 範囲指定での日付検証
- 頗略チェックとエラーハンドリング

##### フィルタリング機能

- 今日のデータを抽出するフィルタリング
- 日付文字列の前方一致で判定
- 作成日、更新日の範囲検索

##### ソート機能

- 日時フィールドでのソート機能
- 昇順・降順の選択可能
- null データの適切な処理

#### 3.3.5 タイムゾーン考慮事項

##### クライアント側のタイムゾーン対応

- ユーザーのローカルタイムゾーンを自動検出
- JST とローカル時間の相互変換
- JST のオフセット（+9 時間）を考慮した計算
- ローカル時間でのデータ送信時の JST 変換

#### 3.3.6 実装上の注意点

1. **一貫性の維持**

   - すべての日時は JST（Asia/Tokyo）で統一
   - フォーマットは `yyyy/MM/dd HH:mm:ss` で統一

2. **パフォーマンス考慮**

   - 大量データの日時処理は非同期で実行
   - 必要に応じてキャッシュを活用

3. **エラーハンドリング**
   - 不正な日時形式に対する適切な処理
   - null や空文字列の適切な扱い

## 4. セキュリティアーキテクチャ

### 4.1 認証・認可

```
User Request → Google OAuth → Token Validation
    ↓
Organization Check → Permission Verification → Access Grant
    ↓
Session Management → Activity Logging → Resource Access
```

### 4.2 データ保護

- **通信**: HTTPS 強制、CSP 設定
- **ストレージ**: Google 暗号化、アクセス制御
- **処理**: サニタイゼーション、SQL インジェクション対策

### 4.3 操作制御

- **管理者機能**: 3 段階確認プロセス
- **データ変更**: 変更履歴記録
- **エラーハンドリング**: 情報漏洩防止

## 5. パフォーマンスアーキテクチャ

### 5.1 フロントエンド最適化

#### 5.1.1 リソース管理

- **CSS/JS**: インライン埋め込みによる HTTP Request 削減
- **画像**: Base64 エンコード、WebP 対応
- **フォント**: Google Fonts CDN 活用

#### 5.1.2 JavaScript 最適化

- **非同期処理**: Promise/async-await 活用
- **イベント**: debounce/throttle 実装
- **DOM 操作**: 最小限の DOM 更新

### 5.2 バックエンド最適化

#### 5.2.1 データ処理最適化

- **バッチ処理**: 複数操作の一括実行
- **キャッシュ**: 頻繁アクセスデータのメモリ保持
- **並列処理**: 独立タスクの並行実行

#### 5.2.2 API 呼び出し最適化

- **レート制限**: Google Apps Script 制限内での処理
- **リトライ**: 一時的障害時の再試行機構
- **タイムアウト**: 長時間処理の適切な制御

### 5.3 データアクセス最適化

#### 5.3.1 Sheets API 最適化

- **範囲指定**: 必要データのみの取得
- **バッチ更新**: 複数セルの一括更新
- **インデックス活用**: 検索性能の向上

#### 5.3.2 プリロード戦略

- **背景読み込み**: ユーザー操作前のデータ準備
- **予測読み込み**: 次に必要なデータの先読み
- **キャッシュ戦略**: LRU 方式でのメモリ管理

## 6. 監視・ログアーキテクチャ

### 6.1 ログ設計

#### 6.1.1 アプリケーションログ

##### ログレベル定義
- **ERROR**: システムエラー、例外処理、致命的なエラー
- **WARN**: 警告、非推奨機能の使用、パフォーマンス低下
- **INFO**: 正常処理、ユーザーアクション、システムイベント
- **DEBUG**: デバッグ情報、詳細な処理フロー（開発環境のみ）

##### ログ項目
```javascript
{
  timestamp: "2024/07/02 10:30:00",
  level: "INFO",
  userId: "user@example.com",
  sessionId: "abc123...",
  action: "URL_GENERATION",
  module: "url-generator",
  message: "拠点管理番号生成完了",
  data: {
    locationNumber: "OSA_SV_ThinkPad_ABC123_001",
    executionTime: 1.2
  },
  userAgent: "Mozilla/5.0...",
  ipAddress: "192.168.1.1"
}
```

##### ログ保存先
- **短期ログ**: Google Apps Script標準ログ（Stackdriver Logging）
- **長期ログ**: Google Sheetsの専用ログシート（90日間保存）
- **エラーログ**: 即座にメール通知 + Sheets記録

#### 6.1.2 パフォーマンスログ

##### 計測項目
- **API応答時間**: 各API関数の実行時間
- **Sheets操作時間**: データ読み書きの処理時間
- **UI表示時間**: ページ読み込みからレンダリング完了まで
- **メモリ使用量**: スクリプト実行時のメモリ消費

##### パフォーマンスログ形式
```javascript
{
  timestamp: "2024/07/02 10:30:00",
  metric: "API_RESPONSE_TIME",
  function: "getLocationAllDevices",
  executionTime: 2.5,  // 秒
  dataSize: 150,       // レコード数
  memoryUsed: 45,      // MB
  cacheHit: false,
  details: {
    sheetsReadTime: 1.8,
    dataProcessingTime: 0.5,
    responseTime: 0.2
  }
}
```

##### 閾値設定
- **警告閾値**: API応答3秒以上、メモリ使用80MB以上
- **エラー閾値**: API応答5秒以上、メモリ使用95MB以上
- **自動アラート**: 閾値超過時にメール通知

### 6.2 監視ポイント

#### 6.2.1 システム稼働監視

##### ヘルスチェック
- **監視間隔**: 5分ごとの定期チェック
- **チェック項目**:
  - Google Apps Script実行環境の応答
  - Google Sheetsへのアクセス可否
  - PropertiesServiceの読み書き確認
  - メモリ使用量の確認
- **判定基準**:
  - 3回連続失敗でアラート発生
  - 応答時間が10秒を超えた場合に警告
- **通知方法**: 
  - メールでの即時通知
  - ログシートへの記録

##### 可用性監視
- **稼働率目標**: 99.5%以上
- **計測方法**: 
  - 5分ごとのping監視
  - ユーザーアクセス成功率
- **ダウンタイム記録**:
  - 開始・終了時刻
  - 影響範囲（機能別）
  - 原因と対応内容

#### 6.2.2 パフォーマンス監視

##### API応答時間
- **計測対象**: 
  - 全API関数の実行時間
  - データベース（Sheets）アクセス時間
  - 外部API呼び出し時間
- **集計項目**:
  - 平均応答時間
  - 最大応答時間
  - 95パーセンタイル値
- **監視閾値**:
  - 平均3秒超過で警告
  - 最大10秒超過でエラー

##### スループット監視
- **計測項目**:
  - 1分あたりのリクエスト数
  - 同時実行数
  - 処理成功率
- **容量閾値**:
  - 100リクエスト/分で警告
  - 150リクエスト/分で制限

##### リソース使用状況
- **メモリ使用量**:
  - 現在使用量の監視
  - ピーク時使用量の記録
  - メモリリーク検出
- **実行時間**:
  - スクリプト実行時間（6分制限）
  - 各関数の実行時間分布

#### 6.2.3 エラー監視

##### エラー検知
- **監視レベル**:
  - FATAL: システム停止レベル
  - ERROR: 機能エラー
  - WARN: 警告レベル
- **検知条件**:
  - エラー率が5%を超過
  - 特定エラーの繰り返し発生
  - 新規エラーパターンの検出

##### アラート設定
- **即時通知条件**:
  - FATALエラー発生
  - エラー率10%超過
  - システム停止
- **通知先**:
  - システム管理者メール
  - Slackチャンネル（オプション）
  - 監視ダッシュボード

##### エラー分析
- **自動分類**:
  - エラータイプ別集計
  - 発生頻度分析
  - 影響範囲特定
- **根本原因分析**:
  - エラースタックトレース
  - 実行コンテキスト
  - 再現手順の記録

#### 6.2.4 利用状況分析

##### アクセス分析
- **ユーザー行動**:
  - ページビュー数
  - 機能別利用回数
  - 滞在時間
  - 離脱率
- **利用パターン**:
  - 時間帯別アクセス
  - 曜日別傾向
  - 拠点別利用状況

##### 機能利用分析
- **機能別統計**:
  - URL生成回数
  - マスタ更新頻度
  - データ検索回数
  - エクスポート利用
- **パフォーマンス相関**:
  - 機能別応答時間
  - データ量との相関
  - 同時利用の影響

##### ユーザー満足度
- **計測指標**:
  - エラー遭遇率
  - 再試行率
  - タスク完了率
- **改善指標**:
  - 機能改善要望
  - 問題報告頻度
  - 利用継続率

## 7. デプロイアーキテクチャ

### 7.1 デプロイ戦略

#### 7.1.1 環境構成

```
Environment
├── develop        # 開発環境（GAS Test Deploy）
└── main           # 本番環境（GAS Web App）
```

#### 7.1.2 デプロイフロー

```
Code Commit → Code Review → Testing → Staging → Production
    ↓            ↓           ↓         ↓          ↓
  GitHub      Pull Request  Auto Test  Manual Test Deployment
```

### 7.2 バージョン管理

- **Git**: ソースコード管理
- **Semantic Versioning**: バージョン体系
- **Release Notes**: リリース内容記録
- **Rollback**: 緊急時巻き戻し手順

### 7.3 設定管理

- **環境別設定**: 環境ごとの設定値管理
- **シークレット管理**: 機密情報の安全な管理
- **設定変更**: 変更履歴とロールバック

## 8. 拡張性アーキテクチャ

### 8.1 モジュール設計

- **疎結合**: 各機能の独立性確保
- **インターフェース**: 標準化された API 設計
- **プラグイン**: 新機能の追加容易性

### 8.2 データ拡張

- **スキーマ進化**: 非破壊的スキーマ変更
- **マイグレーション**: データ移行ツール
- **バックアップ**: 変更前データ保護

### 8.3 統合拡張

- **API 設計**: RESTful API 準拠
- **Webhook**: 外部システム連携
- **エクスポート**: 標準形式でのデータ出力

## 9. フロントエンドアーキテクチャの注意事項

### 9.1 DOM要素の一意性確保

#### 9.1.1 問題の背景
- SPAアプリケーションでは、複数のページが同時にDOM上に存在する
- 同じIDを持つ要素が複数存在すると、`document.getElementById()`は最初に見つかった要素を返す
- これにより、意図しない要素が選択され、予期しない動作を引き起こす

#### 9.1.2 実装ガイドライン

##### ID命名規則
```html
<!-- ❌ 避けるべき実装 -->
<select id="location">  <!-- スプレッドシートページ -->
<select id="location">  <!-- URL生成ページ -->

<!-- ✅ 推奨実装 -->
<select id="spreadsheet-location">  <!-- スプレッドシートページ -->
<select id="url-generator-location">  <!-- URL生成ページ -->
```

##### 要素取得の最適化
```javascript
// ❌ 避けるべき実装
const locationSelect = document.getElementById('location');

// ✅ 推奨実装1: 属性セレクターを使用
const locationSelect = document.querySelector('select#location[onchange*="updateLocationNumber"]');

// ✅ 推奨実装2: ページコンテキストを考慮
const urlGeneratorPage = document.getElementById('url-generator');
const locationSelect = urlGeneratorPage.querySelector('#location');

// ✅ 推奨実装3: 一意のIDを使用
const locationSelect = document.getElementById('url-generator-location');
```

#### 9.1.3 既存コードの修正方針

##### 段階的な修正アプローチ
1. **即時対応**: セレクターを使用した回避策
   ```javascript
   // onchange属性やdata属性を利用した一意性の確保
   document.querySelector('select#location[onchange*="特定の関数名"]');
   ```

2. **中期対応**: ID命名規則の統一
   - ページプレフィックスの追加: `{page}-{element}`
   - 例: `spreadsheet-location`, `url-generator-location`

3. **長期対応**: コンポーネント化
   - 各ページを独立したコンポーネントとして設計
   - スコープ付きIDまたはクラス名の使用

##### デバッグ支援機能
```javascript
// 同じIDを持つ要素の検出
function detectDuplicateIds() {
  const allElements = document.querySelectorAll('[id]');
  const idMap = {};
  
  allElements.forEach(el => {
    const id = el.id;
    if (!idMap[id]) {
      idMap[id] = [];
    }
    idMap[id].push(el);
  });
  
  Object.entries(idMap).forEach(([id, elements]) => {
    if (elements.length > 1) {
      console.warn(`重複ID検出: "${id}" (${elements.length}個)`, elements);
    }
  });
}
```

#### 9.1.4 ベストプラクティス

1. **要素の存在確認**
   ```javascript
   const element = document.querySelector('適切なセレクター');
   if (!element) {
     console.error('要素が見つかりません');
     return;
   }
   ```

2. **コンテキストを意識した要素取得**
   ```javascript
   // 特定のページ内でのみ要素を検索
   const pageContainer = document.getElementById('specific-page');
   const targetElement = pageContainer?.querySelector('.target-class');
   ```

3. **イベントリスナーの適切な管理**
   ```javascript
   // ページ遷移時に古いリスナーを削除
   element.removeEventListener('change', oldHandler);
   element.addEventListener('change', newHandler);
   ```

### 9.2 ページ間の状態管理

#### 9.2.1 グローバル変数の適切な使用
- ページ固有の状態は、ページプレフィックスを付けて管理
- 共通状態は専用のオブジェクトで管理

```javascript
// ページ固有の状態
window.urlGeneratorState = {
  selectedLocation: '',
  generatedUrl: ''
};

// 共通状態
window.appState = {
  currentUser: '',
  debugMode: false
};
```

#### 9.2.2 イベントの伝播制御
- ページ間でのイベント干渉を防ぐ
- 必要に応じてevent.stopPropagation()を使用

### 9.3 パフォーマンス最適化

#### 9.3.1 セレクターの最適化
- 可能な限り具体的なセレクターを使用
- querySelectorAllの使用を最小限に抑える
- キャッシュ可能な要素参照はキャッシュする

#### 9.3.2 DOM操作の最適化
- バッチ更新の実装
- DocumentFragmentの活用
- 不要な再描画の回避

## 10. QRコード生成アーキテクチャ

### 10.1 QRコード生成APIの選定

#### 10.1.1 API選定基準

1. **セキュリティ**
   - HTTPS通信の必須
   - SSL/TLS証明書の正当性検証
   - 悪意のあるコード注入の防止

2. **信頼性**
   - 高い稼働率（99.9%以上）
   - 安定したAPIレスポンス
   - 長期間の運用実績

3. **パフォーマンス**
   - 低レイテンシ（200ms以下）
   - 高速な画像生成
   - 適切なキャッシュメカニズム

#### 10.1.2 採用API一覧

##### 第1優先: GoQR.me API (QR Server)
```
https://api.qrserver.com/v1/create-qr-code/
```
- **提供元**: goqr.me
- **セキュリティ**: 
  - HTTPS対応
  - EU GDPR準拠
  - データ保存なし
  - プライバシー重視
- **特徴**:
  - オープンソース
  - APIキー不要
  - 高い信頼性（稼働率99.95%）
  - QRコード規格に完全準拠
  - 2007年から運用の長い実績

##### 第2優先: QRCode Monkey API
```
https://api.qr-code-generator.com/v1/create
```
- **提供元**: QRCode Monkey
- **セキュリティ**:
  - HTTPS対応
  - SSL/TLS暗号化
- **特徴**:
  - 商用利用可能
  - 高解像度対応
  - カスタマイズ可能

##### 第3優先: QuickChart.io
```
https://quickchart.io/qr
```
- **提供元**: QuickChart
- **セキュリティ**:
  - HTTPS対応
  - データログ保存なし
- **特徴**:
  - 無料利用可能
  - シンプルAPI
  - エラー訂正レベル設定可能
  - 高速レスポンス

#### 10.1.3 Google Charts APIの廃止対応

2025年1月現在、Google Charts APIの一部機能が廃止されているため、代替のAPIへ移行しました。

### 10.2 実装アーキテクチャ

#### 10.2.1 サーバーサイド処理

```javascript
// Google Apps Scriptでの実装
function generateQRCodeImage(url) {
  // 1. QR Server APIを試行
  // 2. 失敗時はQuickChart.ioを試行
  // 3. さらに失敗時はQR Code Generator APIを試行
  // 4. Base64エンコードして返却
}
```

#### 10.2.2 セキュリティ対策

1. **入力検証**
   - URLの正当性確認
   - 特殊文字の適切なエンコード
   - URL長制限（2000文字以下）

2. **通信セキュリティ**
   ```javascript
   UrlFetchApp.fetch(url, {
     muteHttpExceptions: true,
     validateHttpsCertificates: true,  // SSL証明書検証
     headers: {
       'User-Agent': 'Google Apps Script'  // ユーザーエージェント明示
     }
   });
   ```

3. **エラーハンドリング**
   - API失敗時のフォールバック
   - タイムアウト設定
   - リトライ処理

#### 10.2.3 パフォーマンス最適化

1. **サーバーサイド生成**
   - CORS制限の回避
   - クライアントの負荷軽減
   - セキュリティの向上

2. **Base64エンコード**
   - 追加HTTPリクエストの削減
   - クライアントでの即座表示
   - キャッシュ効率の向上

3. **APIフォールバック**
   - 複数APIの並列試行
   - 最速応答の採用
   - 障害時の自動切り替え

### 10.3 運用ガイドライン

#### 10.3.1 監視
- API応答時間のログ記録
- 失敗率の追跡
- プロバイダー別の成功率分析

#### 10.3.2 メンテナンス
- APIの健全性定期確認
- 新しいAPIの評価と追加
- 廃止予定APIの事前対応

#### 10.3.3 セキュリティ更新
- APIプロバイダーのセキュリティ情報確認
- 脆弱性情報のモニタリング
- 定期的なセキュリティ監査
