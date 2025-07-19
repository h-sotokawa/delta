# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

代替機管理システム - Google Apps Script (GAS) ベースの企業向けIT機器管理Webアプリケーション

### 技術スタック
- **バックエンド**: Google Apps Script (JavaScript)
- **フロントエンド**: HTML, CSS, vanilla JavaScript
- **データストア**: Google Sheets
- **フォーム**: Google Forms
- **QRコード**: Firebase Hosting (リダイレクト用)


## アーキテクチャ概要

### ファイル構成
```
バックエンド (.gs):
- Code.gs (2890行) - メインビジネスロジック ※分割予定
- config.gs - システム設定と定数
- error-handling.gs - エラー処理
- date-utils.gs - 日付ユーティリティ
- custody-status-manager.gs - 預り機ステータス管理
- url-generation.gs - URL生成
- qr-*.gs - QRコード関連

フロントエンド (.html):
- Index.html - エントリーポイント
- *-functions.html - JavaScript機能
- *-styles.html - CSSスタイル
- common-ui-*.html - 共通UIコンポーネント
```

### データフロー
1. **ユーザーインターフェース** (HTML/JS) → **Google Apps Script** → **Google Sheets**
2. **Google Forms** → **Google Sheets** → **ステータス管理システム**
3. **QRコード** → **Firebase** → **Google Forms**

### 主要機能モジュール
1. **マスタデータ管理**: 拠点、機種、端末、プリンタ、データタイプ
2. **ステータス管理**: 機器の貸出・返却・預り状態
3. **フォーム生成**: Google Forms連携、QRコード生成
4. **データビューア**: フィルタリング、検索、エクスポート

## コーディング規約

### 言語設定
- **ドキュメント・コメント**: 日本語
- **変数・関数名**: 英語 (camelCase)
- **思考プロセス**: 英語のみ
- **ユーザーへの応答**: 日本語

### 日時処理の標準仕様

#### 基本原則
Google Sheetsへの日時保存は**標準テキスト形式**を使用（アポストロフィなし）

#### 実装方法
```javascript
// ✅ 正しい実装
const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

// ❌ 避けるべき実装
const today = new Date(); // 自動変換される可能性
```

#### 列の書式設定
```javascript
// 日付列のセル書式をテキストに設定
dateRange.setNumberFormat('@'); // '@' = テキスト形式
```

#### 表示用フォーマット関数
```javascript
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  try {
    // 既に文字列の場合はそのまま返す
    if (typeof dateValue === 'string') {
      return dateValue;
    }
    
    // Dateオブジェクトの場合
    if (dateValue instanceof Date) {
      const year = dateValue.getFullYear();
      const month = (dateValue.getMonth() + 1).toString().padStart(2, '0');
      const day = dateValue.getDate().toString().padStart(2, '0');
      return `${year}/${month}/${day}`;
    }
    
    return dateValue.toString();
  } catch (error) {
    console.error('日付フォーマットエラー:', error, dateValue);
    return '';
  }
}
```

### エラーハンドリング
```javascript
// カスタムエラークラスを使用
class AppError extends Error {
  constructor(message, statusCode = 500, details = null) {
    super(message);
    this.name = 'AppError';
    this.statusCode = statusCode;
    this.details = details;
  }
}

// 安全な関数実行
function safeExecute(fn, errorMessage) {
  try {
    return fn();
  } catch (error) {
    handleError(error, errorMessage);
    throw error;
  }
}
```

## 重要な実装詳細

### 拠点管理番号フォーマット (2025年1月更新)
新形式: `拠点_カテゴリ_モデル_製造番号_連番`
例: `OSA_SV_Server_ABC12345_001`

### ステータス管理
- 機器ステータス: 貸出中、返却済み、修理中など
- 預り機ステータス: より詳細な状態管理
- 変更通知: メール通知機能あり

### データ検証
- `config.gs`で定義されたパターンによるバリデーション
- 必須フィールドチェック
- メールアドレス形式検証

## 開発上の注意点

### パフォーマンス
- Google Sheets APIの呼び出し回数を最小限に
- バッチ処理を使用（`getValues()`、`setValues()`）
- SpreadsheetApp.flush()の適切な使用

### セキュリティ
- 権限管理に注意（Google Sheetsアクセス権限）
- デバッグモードは本番環境で必ずOFF
- APIキーやトークンをコードに含めない

### デバッグ
```javascript
// config.gsでDEBUGフラグを管理
if (DEBUG) {
  console.log('デバッグ情報:', data);
}

// Google Sheetsでの確認
// セルを選択 → 数式バーで実際の値を確認
```

## スプレッドシートビューアー既知の問題と対策

### 問題の概要
修正作業後にスプレッドシートビューアーが動作しなくなる現象が過去に発生しています。

### 主な原因

#### 1. 日時フォーマット関連の問題
```javascript
// エラーメッセージ例
'エラーが発生しました : スプレッドシート内の日時の表示形式を"書式なし"に変更してください'
```

**対策**:
- Google Sheetsの日時列の書式を「書式なしテキスト」に設定
- `dateRange.setNumberFormat('@')` で書式設定
- アポストロフィなしの標準テキスト形式に統一

#### 2. データ構造の変更による影響
- **列インデックスの変更**: マスタシートの構造変更時
- **関連性の破綻**: 拠点マスタとの不整合

#### 3. Google Apps Script環境の問題
```javascript
// 環境チェック
if (!google || !google.script || !google.script.run) {
  throw new Error('Google Apps Scriptの実行環境が正しく初期化されていません');
}
```

### 修正後の確認手順

#### 1. 即座に確認すべき項目
```javascript
// デバッグモードの有効化
const DEBUG_MODE = true; // config.gsで設定
```

1. **スプレッドシートの書式設定**
   - 日時列が「書式なしテキスト」になっているか

2. **Google Apps Scriptの権限**
   - スクリプト実行権限の確認
   - 認証状態のチェック

3. **ブラウザコンソールでのエラー確認**
   - F12開発者ツール → Console
   - ネットワークタブでAPI呼び出し失敗の確認

#### 2. 段階的テスト手順
1. **データタイプ選択機能**: プルダウンが正常に表示されるか
2. **拠点選択機能**: 拠点リストが正常に読み込まれるか  
3. **データ読み込み機能**: 実際のデータが表示されるか

#### 3. バックアップと復旧
```javascript
// フォールバック処理の例
if (masterData.length === 0) {
  console.log('マスタデータが空のため、簡易的な処理を実行');
  // 緊急時の代替処理
}
```

### 重要ファイル（修正時要注意）
- `spreadsheet-functions.html` (40,841行) - メイン機能
- `Code.gs` - バックエンドロジック
- `data-type-master.gs` - データタイプ管理
- `error-handling.gs` - エラー処理
- `config.gs` - 設定管理

### 4. HTMLインクルード機能の問題
Google Apps ScriptのHTMLファイル分割システムに起因する問題：

```javascript
// Code.gsのinclude関数（173-188行）
function include(filename) {
  try {
    const fileToInclude = filename.includes('.') ? filename : filename + '.html';
    const content = HtmlService.createHtmlOutputFromFile(fileToInclude).getContent();
    return content;
  } catch (error) {
    return `<!-- Error loading ${filename}: ${error.toString()} -->`;
  }
}
```

**問題点**:
- **31個のHTMLファイル**を動的に読み込み（Index.htmlから）
- 実行時間制限による不完全な読み込みの可能性
- ファイル読み込みエラーがサイレント（HTMLコメントのみ）

**影響**:
- スプレッドシートビューアーの部分的な機能停止
- CSSやJavaScriptの読み込み失敗
- ユーザーには見えないエラー状態

**対策**:
```javascript
// Index.htmlでの使用例
<?!= include('spreadsheet-functions'); ?>
<?!= include('spreadsheet-styles'); ?>
```
- 重要ファイルの存在確認
- ブラウザの開発者ツールでHTMLソースの確認
- エラーコメントの検索：`<!-- Error loading`

### 予防策
1. **修正前**: 該当機能の動作確認
2. **修正中**: デバッグモードでの段階的テスト
3. **修正後**: 全機能の動作確認と権限チェック
4. **HTMLインクルード**: ブラウザでHTMLソースの完全性確認

## 2025年1月: 管轄機能実装

### 実装内容
1. **拠点マスタの拡張**
   - 「管轄」フィールド追加（必須、関西・関東・九州等）
   - 「ステータス変更通知」フィールド追加（Boolean、拠点別通知制御）
   - バリデーションルールを設計書準拠に修正

2. **管轄ベースフィルタリング**
   - 監査データ: 管轄選択 → 拠点選択の2段階フィルタリング
   - サマリーデータ: 管轄別のデータ集計表示
   - `getLocationsByJurisdiction()`, `getJurisdictionList()` 関数追加

3. **関連ファイルの更新**
   - Code.gs: 拠点マスタ管理関数の拡張
   - location-master.html: UI更新
   - location-master-functions.html: フロントエンド機能追加
   - spreadsheet.html: 管轄選択UI追加
   - spreadsheet-functions.html: フィルタリング機能実装

## 2025年1月: 統合ビュー分離・リアルタイム更新システム実装

### 実装内容
1. **統合ビューの分離**
   - 端末系統合ビュー（integrated_view_terminal）：Server、Desktop、Laptop、Tablet
   - プリンタ・その他系統合ビュー（integrated_view_printer_other）：Printer、Router、Hub、Other
   - 旧統合ビュー（integrated_view）：非推奨として存続

2. **リアルタイム更新システム**
   - onFormSubmitトリガー：フォーム送信時の即時更新
   - onChangeトリガー：マスタシート変更時の自動反映
   - timeBasedトリガー：深夜2:00の日次全体再構築

3. **統合ビューの表示制御**
   - 端末系統合ビュー（INTEGRATED_VIEW_TERMINAL）
   - プリンタ・その他系統合ビュー（INTEGRATED_VIEW_PRINTER_OTHER）
   - デバイスタイプ選択の不要化
   - 拠点選択のみでの効率的フィルタリング

4. **パフォーマンス最適化**
   - 部分更新による処理時間短縮
   - 並列処理での端末系・プリンタ系同時更新
   - インデックス活用による高速検索

## 2025年1月: サマリー表示の動的化

### 問題点
- サマリーカード生成で拠点名をハードコード（`'大阪' || '神戸' || '姫路'`）
- 新規拠点追加時に自動反映されない

### 修正内容
1. **動的拠点名検証機能**
   - `isValidLocationName()`: 拠点マスタベースの動的検証
   - ハードコード削除、拠点マスタから動的に取得

2. **拡張性の実現**
   - 新規拠点追加時の自動反映
   - 拠点名変更時の自動対応
   - 管轄フィルタリングとの完全統合

### テスト関数
```javascript
// 管轄機能の統合テスト
testJurisdictionFeatures()

// サマリーデータの動的拠点表示テスト
testDynamicSummaryDisplay()
```

## 2025年1月: サマリー表示の動的拠点フィルタリング修正

### 変更理由
サマリー表示で拠点がハードコードされていた問題を修正し、全管轄表示時に全拠点が表示されない問題を解決

### 変更内容

#### ハードコード削除
1. **opening-screen-functions.html**:
   - フォールバック拠点（大阪・神戸・姫路）を削除
   - 拠点マスタ未読み込み時は空配列を使用

2. **url-generator-functions.html**:
   - 静的拠点データを削除
   - 拠点マスタから動的取得に変更

#### 拠点マスタデータ読み込み修正
**spreadsheet-functions.html**:
1. **サマリー表示開始時の修正**:
   - `loadDataBasedOnDataType()`関数で拠点マスタデータを確実に読み込む
   - 読み込み完了後にサマリーデータを取得する順序に変更

2. **サマリー更新時の修正**:
   - `refreshSummaryData()`関数でも同様の処理を追加
   - 既存データがある場合は再利用

#### 全管轄対応修正
**filterSummaryDataByJurisdiction()関数**:
```javascript
// 変更前: !jurisdictionで早期リターン
if (!data || data.length === 0 || !jurisdiction) {
  return data;
}

// 変更後: 全管轄（空文字）の場合も適切に処理
if (!data || data.length === 0) {
  return Promise.resolve(data);
}
if (!jurisdiction) {
  return Promise.resolve(data); // 全管轄の場合は全データを返す
}
```

### 影響範囲
- **サマリー表示**: 拠点マスタに登録された全拠点が動的に表示される
- **管轄フィルタリング**: 選択した管轄の拠点のみが表示される
- **全管轄表示**: すべての拠点が正しく表示される

## 2025年1月: 列名動的取得システムの全面実装

### 実装理由
Google Sheetsの列順序や列名の変更に対する柔軟性を高め、ハードコードされた列インデックスによる脆弱性を排除

### 実装内容

#### 1. ヘルパー関数の作成（column-helper.gs）
```javascript
// 基本ヘルパー関数
getColumnIndex(headers, columnName)      // 0ベースのインデックス取得
getColumnNumber(headers, columnName)     // 1ベースの列番号取得
getValueByColumnName(row, headers, columnName) // 安全な値取得
getColumnIndexMultiple(headers, columnNames)   // 複数候補から取得
rowToObject(row, headers)               // 行をオブジェクトに変換
getCachedHeaders(sheetName)            // ヘッダーのキャッシュ
```

#### 2. 全GSファイルの更新
以下のファイルで列インデックスのハードコーディングを削除：
- view-sheet-operations.gs
- spreadsheet-setup.gs
- search-index-builder.gs
- spreadsheet-operations.gs
- integrated-view-rebuild.gs
- integrated-view-triggers.gs
- location-master-operations.gs
- model-master-operations.gs
- custody-status-manager.gs

#### 3. 実装パターン

**変更前（ハードコード）**:
```javascript
const status = row[4];  // ステータスは5列目
const managementNumber = row[0];  // 管理番号は1列目
```

**変更後（動的取得）**:
```javascript
const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
const status = getValueByColumnName(row, headers, 'ステータス');
const managementNumber = getValueByColumnName(row, headers, '拠点管理番号');
```

#### 4. 設計ドキュメントの作成
`/workspace/.doc/03_design/09_column_dynamic_access.md`に詳細な設計書を作成：
- 設計原則と命名規則
- ヘルパー関数の仕様
- 実装パターン
- 移行ガイドライン
- パフォーマンス考慮事項
- トラブルシューティング

### 影響範囲
- **全データ処理**: 列順序変更に自動対応
- **保守性向上**: 列名変更時の修正箇所最小化
- **エラー耐性**: 列が見つからない場合の安全な処理
- **可読性向上**: 列名による自己文書化

## 未実装・改善予定

### 優先度：高
1. **Code.gsの分割**: 3000行超の巨大ファイルを機能別に分割
2. **ステータス別表示機能**: 機器状態によるフィルタリング強化
3. **2段階ネストフィルタリング**: より高度な検索機能

### 優先度：中
1. **パフォーマンス最適化**: キャッシュ機能の実装
2. **バッチ処理の改善**: 大量データ処理の効率化
3. **エラーレポート機能**: より詳細なエラー追跡

## トップレベルルール

- このドキュメントは日本語で記述します。コードのコメントも日本語で記述してください。
- 効率を最大化するために、**複数の独立したプロセスを実行する必要がある場合は、ツールを順次ではなく並行して呼び出してください**。
- **思考は英語のみで行ってください**。ただし、**応答は日本語で行う必要があります**。
- ライブラリの使用方法を理解するには、**常にContext MCPを使用して最新情報を取得してください**。

## 関数実行時の注意事項

### 関数指定時のファイル名明示
関数を実行する際は、必ずファイル名を明示してください。複数のファイルに同名の関数が存在する場合があります。

**正しい指定方法**:
```javascript
// debug-sheet-names.gs
debugStatusCollectionData();

// debug-integration-process.gs  
debugIntegrationProcess();

// test-integrated-view-columns.gs
testUpdateIntegratedViews();
```

**避けるべき指定方法**:
```javascript
// ファイル名なし（どのファイルの関数か不明）
debugStatusCollectionData();
```

### 主要なテスト関数とファイル
- **debug-sheet-names.gs**: `debugSheetNames()`, `debugStatusCollectionData()`
- **debug-integration-process.gs**: `debugIntegrationProcess()`, `testSingleDeviceIntegration()`
- **test-integrated-view-columns.gs**: `testIntegratedViewColumns()`, `testUpdateIntegratedViews()`, `testValidateIntegratedViewData()`
- **setup-test-master-data.gs**: `createAllTestData()`, `setupAllTestMasterData()`