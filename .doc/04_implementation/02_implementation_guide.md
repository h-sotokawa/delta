# 実装ガイド

## 1. 実装概要

### 1.1 実装アプローチ
- **段階的実装**: 機能単位での段階的な開発・デプロイ
- **モジュール化設計**: 疎結合な機能モジュールの構築
- **テスト駆動開発**: テストファーストでの品質確保
- **継続的改善**: 実装後のフィードバック反映

### 1.2 実装優先順位
```
Phase 1 (基盤): 
├── 基本画面構成・ナビゲーション
├── Google Sheets連携基盤
├── 認証・セッション管理
└── エラーハンドリング基盤

Phase 2 (マスタ管理):
├── 拠点マスタ管理
├── 機種マスタ管理
├── 基本的なCRUD操作
└── データバリデーション

Phase 3 (コア機能):
├── URL生成機能
├── 拠点管理番号生成
├── マスタ連携・自動保存
└── QRコード機能

Phase 4 (表示機能):
├── スプレッドシートビューアー
├── ダッシュボード機能
├── 検索・フィルタ機能
├── 統合ビュー分離（端末系・プリンタ系）
└── リアルタイム更新システム

Phase 5 (管理機能):
├── システム設定管理
├── ログ・監視機能
├── パフォーマンス最適化
└── セキュリティ強化
```

## 2. 開発環境セットアップ

### 2.1 Google Apps Script プロジェクト作成

#### 2.1.1 新規プロジェクト作成手順
```
1. Google Drive にアクセス
2. 新規 > その他 > Google Apps Script を選択
3. プロジェクト名を「代替機管理システム」に変更
4. マニフェストファイル（appsscript.json）を有効化
```

#### 2.1.2 プロジェクト設定
- タイムゾーン: Asia/Tokyo
- 高度なサービス: Google Sheets API v4を有効化
- Webアプリ: ドメイン内アクセス、実行ユーザー権限
- 実行API: ドメイン内アクセス許可
```

### 2.2 Google Sheets準備

#### 2.2.1 メインスプレッドシート作成
**必要なシート構成**:
- 拠点マスタ: 拠点情報管理
- 機種マスタ: 機種情報管理  
- データタイプマスタ: 表示形式定義
- 端末マスタ: 端末機器管理
- プリンタマスタ: プリンタ機器管理
- その他マスタ: その他機器管理
- integrated_view_terminal: 端末系統合ビュー
- integrated_view_printer_other: プリンタ・その他系統合ビュー
- search_index: 検索インデックス

**初期化処理**:
- Google Apps Scriptで自動シート作成
- 各シートのヘッダー設定
- 書式設定（背景色、フォント等）
- スプレッドシートIDの取得・設定

### 2.3 初期設定

#### 2.3.1 PropertiesService設定
**必須設定項目**:
- SPREADSHEET_ID_MAIN: メインスプレッドシートID
- TERMINAL_COMMON_FORM_URL: 端末用Google FormsのURL
- PRINTER_COMMON_FORM_URL: プリンタ用Google FormsのURL
- QR_REDIRECT_URL: QR中間ページURL
- ERROR_NOTIFICATION_EMAIL: エラー通知メールアドレス
- ALERT_NOTIFICATION_EMAIL: アラート通知メールアドレス
- DEBUG_MODE: デバッグモード（true/false）

**debugMode制御対象設定**:
- debugMode=trueの場合のみ編集可能:
  - 共通フォームURL（端末用・プリンタ用・QR中間ページ）
  - エラー通知メールアドレス
  - アラート通知メールアドレス
- debugMode=falseの場合は読み取り専用表示

**設定方法**:
- Google Apps ScriptのPropertiesServiceを使用
- 環境別設定の管理
- 設定値の永続化

## 3. コアモジュール実装

### 3.1 基盤機能実装

#### 3.1.1 HTMLテンプレート機能
**Include関数**:
- HTMLファイルの動的インクルード機能
- エラーハンドリング付きファイル読み込み
- コンポーネント化によるコード再利用

**Webアプリエントリーポイント**:
- doGet()関数によるメイン処理
- HTMLテンプレートの評価・生成
- メタタグ設定（viewport等）
- エラー時の代替ページ表示

#### 3.1.2 ログ・パフォーマンス管理
**ログ機能**:
- 統一されたログ管理システム
- パフォーマンス測定機能
- エラー情報の詳細記録
- タイムスタンプ付きログ出力
 */
let serverLogs = [];

function addLog(message, data = {}) {
  const logEntry = {
    timestamp: new Date().toISOString(),
    message: message,
    data: data,
    user: Session.getActiveUser().getEmail()
  };
  
  serverLogs.push(logEntry);
  console.log(`[LOG] ${message}`, data);
  
  // ログの上限管理（メモリ対策）
  if (serverLogs.length > 1000) {
    serverLogs = serverLogs.slice(-500);
  }
}

/**
 * パフォーマンス計測
 */
function startPerformanceTimer() {
  return Date.now();
}

function endPerformanceTimer(startTime, operation) {
  const endTime = Date.now();
  const duration = endTime - startTime;
  
  addLog(`パフォーマンス: ${operation}`, {
    duration: duration + 'ms',
    operation: operation
  });
  
  // 長時間処理の警告
  if (duration > 5000) {
    console.warn(`長時間処理検出: ${operation} - ${duration}ms`);
  }
  
  return duration;
}
```

#### 3.1.3 エラーハンドリング
**レスポンス形式**:
- 統一エラーレスポンス（success、error、errorCode、details、timestamp）
- 統一成功レスポンス（success、data、message、timestamp）
- タイムスタンプ付きレスポンス生成

**API実行ラッパー**:
- 安全なAPI実行環境の提供
- パフォーマンス測定の自動化
- 例外処理の統一化
- ログ記録の自動化

### 3.2 データアクセス層実装

#### 3.2.1 Google Sheets基本操作
**スプレッドシート接続**:
- PropertiesServiceからスプレッドシートID取得
- スプレッドシートへの安全な接続
- 接続エラー時の適切な例外処理

**シートアクセス**:
- シート名による安全な取得機能
- 存在チェック付きシートアクセス
- エラーハンドリング機能

**データ取得処理**:
- データ範囲の自動検出
- ヘッダー行とデータ行の分離
- オブジェクト配列への変換
- 空データの適切な処理

#### 3.2.2 マスタデータ操作
```javascript
// Code.gs
/**
 * 拠点マスタ取得
 */
function getLocationMaster() {
  return safeApiCall(() => {
    const data = getSheetData('拠点マスタ');
    return createSuccessResponse(data);
  });
}

/**
 * 拠点追加
 */
function addLocation(locationData) {
  return safeApiCall(() => {
    const sheet = getSheetByName('拠点マスタ');
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    
    // バリデーション
    if (!locationData.locationId || !locationData.locationName) {
      throw new Error('拠点IDと拠点名は必須です');
    }
    
    // 端末マスタの場合は資産管理番号が必須
    if (masterType === '端末マスタ' && !deviceData.assetNumber) {
      throw new Error('端末マスタでは資産管理番号は必須です');
    }
    
    // 重複チェック
    const existingData = getSheetData('拠点マスタ');
    const duplicate = existingData.find(row => row['拠点ID'] === locationData.locationId);
    if (duplicate) {
      throw new Error(`拠点ID '${locationData.locationId}' は既に存在します`);
    }
    
    // データ追加
    const newRow = [
      locationData.locationId,
      locationData.locationCode,
      locationData.locationName,
      locationData.groupEmail || '',
      locationData.status || 'active',
      now,
      now
    ];
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    
    addLog('拠点追加完了', { locationId: locationData.locationId });
    
    return createSuccessResponse({
      addedLocationId: locationData.locationId,
      savedRow: lastRow + 1
    }, '拠点が正常に追加されました');
  });
}

/**
 * 利用可能カテゴリ一覧取得
 */
function getAvailableCategories() {
  return safeApiCall(() => {
    const modelData = getSheetData('機種マスタ');
    
    // カテゴリ別にグループ化して件数をカウント
    const categoryMap = {};
    modelData.forEach(model => {
      const category = model['カテゴリ'];
      if (!categoryMap[category]) {
        categoryMap[category] = {
          categoryId: category,
          categoryName: getCategoryDisplayName(category),
          modelCount: 0
        };
      }
**カテゴリマッピング機能**:
- カテゴリ別グループ化と件数カウント
- 英語カテゴリから日本語表示名への変換
- ソート機能付きカテゴリ一覧返却

### 3.3 URL生成機能実装

#### 3.3.1 拠点管理番号生成
**生成ロジック**:
- 拠点コード、カテゴリ、モデル、製造番号、連番から組み立て
- カテゴリマッピング（Server→SV等）
- モデル名の正規化（空白・記号除去）
- 連番の3桁ゼロパディング
- フォーマット: `拠点_カテゴリ_モデル_製造番号_連番`

**バリデーション機能**:
- 必須項目の入力チェック
- 文字数制限（100文字以内）
- 形式チェック

**重複チェック**:
- 全マスタでの重複確認
- エラー時の適切なハンドリング
- シート不存在時の続行処理

#### 3.3.2 URL生成・保存
```javascript
// url-generation.gs
/**
 * 共通フォームURL生成
 */
function generateCommonFormUrl(locationNumber, deviceCategory, generateQrUrl = false) {
  return safeApiCall(() => {
    // 設定取得
    const settings = getCommonFormsSettings();
    if (!settings.success) {
      throw new Error('設定の取得に失敗しました');
    }
    
    // QRコード用URL生成
    if (generateQrUrl && settings.data.qrRedirectUrl) {
      const qrUrl = `${settings.data.qrRedirectUrl}?id=${encodeURIComponent(locationNumber)}`;
      return createSuccessResponse({
        url: qrUrl,
        baseUrl: settings.data.qrRedirectUrl,
        locationNumber: locationNumber,
        deviceCategory: deviceCategory,
        isQrUrl: true
      });
    }
    
    // 通常のフォームURL生成
    let baseUrl;
    let formType;
    
    if (['desktop', 'laptop', 'server'].includes(deviceCategory)) {
      baseUrl = settings.data.terminalCommonFormUrl;
      formType = '端末';
    } else {
      baseUrl = settings.data.printerCommonFormUrl;
      formType = 'プリンタ・その他';
    }
    
    if (!baseUrl) {
      throw new Error(`${formType}用共通フォームURLが設定されていません`);
    }
    
    // URLパラメータ追加
    const separator = baseUrl.includes('?') ? '&' : '?';
    const generatedUrl = `${baseUrl}${separator}entry.1372464946=${encodeURIComponent(locationNumber)}`;
    
    return createSuccessResponse({
      url: generatedUrl,
      baseUrl: baseUrl,
      locationNumber: locationNumber,
      deviceCategory: deviceCategory,
      isQrUrl: false
    });
  });
}
```

### 3.4 フロントエンド実装

#### 3.4.1 メイン画面構造
```html
<!-- Index.html -->
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <meta charset="UTF-8">
    <title>代替機管理システム</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    
    <!-- 外部リソース -->
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
    
    <!-- スタイルシート -->
    <?!= include('styles'); ?>
    <?!= include('dashboard-styles'); ?>
    <?!= include('url-generator-styles'); ?>
    <!-- 他のスタイルも同様に読み込み -->
  </head>
  <body>
    <!-- ナビゲーション -->
    <?!= include('navigation'); ?>
    
    <!-- メインコンテンツ -->
    <div id="main-content">
      <!-- ダッシュボード -->
      <?!= include('dashboard'); ?>
      
      <!-- URL生成ページ -->
      <div id="url-generator" class="page-content">
        <?!= include('url-generator'); ?>
      </div>
      
      <!-- 他のページも同様に配置 -->
    </div>
    
    <!-- JavaScript -->
    <?!= include('main'); ?>
    <?!= include('url-generator-functions'); ?>
    <!-- 他のスクリプトも同様に読み込み -->
  </body>
</html>
```

#### 3.4.2 ページ遷移制御
```html
<!-- main.html -->
<script>
  // グローバル変数
  const DEBUG_MODE = false;
  
  // 初期化
  window.onload = function() {
    if (DEBUG_MODE) {
      console.log('デバッグモード: 初期化開始');
    }
    // ダッシュボードを直接表示
    showPage('dashboard');
  };
  
  // ページ遷移機能
  function showPage(pageId) {
    // 全ページ非表示
    document.querySelectorAll('.page-content').forEach(page => {
      page.classList.remove('active');
    });
    
    // 全ナビリンク非アクティブ
    document.querySelectorAll('.nav-link').forEach(link => {
      link.classList.remove('active');
    });
    
    // 指定ページ表示
    const targetPage = document.getElementById(pageId);
    if (targetPage) {
      targetPage.classList.add('active');
    }
    
    // 対応ナビリンクアクティブ化
    const targetLink = document.querySelector(`[onclick*="showPage('${pageId}')"]`);
    if (targetLink) {
      targetLink.classList.add('active');
    }
    
    // ページ固有の初期化
    initializePage(pageId);
  }
  
  // ページ固有初期化
  function initializePage(pageId) {
    switch (pageId) {
      case 'url-generator':
        if (typeof onUrlGeneratorPageShow === 'function') {
          onUrlGeneratorPageShow();
        }
        break;
      case 'model-master':
        if (typeof onModelMasterPageShow === 'function') {
          onModelMasterPageShow();
        }
        break;
      // 他のページも同様
    }
  }
  
  // エラー表示
  function showErrorMessage(message) {
    const errorDiv = document.getElementById('error-message') || createErrorDiv();
    errorDiv.querySelector('#error-text').textContent = message;
    errorDiv.style.display = 'block';
    
    setTimeout(() => {
      errorDiv.style.display = 'none';
    }, 5000);
  }
  
  // 成功メッセージ表示
  function showSuccessMessage(message) {
    const successDiv = document.getElementById('success-message') || createSuccessDiv();
    successDiv.querySelector('#success-text').textContent = message;
    successDiv.style.display = 'block';
    
    setTimeout(() => {
      successDiv.style.display = 'none';
    }, 3000);
  }
</script>
```

## 4. テスト実装

### 4.1 単体テスト実装例
```javascript
// Code.gs（テスト関数）
/**
 * 拠点管理番号生成テスト
 */
function testGenerateLocationNumber() {
  console.log('=== 拠点管理番号生成テスト開始 ===');
  
  const testCases = [
    {
      name: '正常ケース1',
      input: {
        locationCode: 'OSAKA',
        category: 'Server',
        model: 'ThinkPad-X1-Carbon',
        serialNumber: 'ABC123456',
        sequence: 1
      },
      expected: 'OSAKA_SV_ThinkPad-X1-Carbon_ABC123456_001'
    },
    {
      name: '正常ケース2（プリンタ）',
      input: {
        locationCode: 'KOBE',
        category: 'Printer',
        model: 'imageClass-MF644Cdw',
        serialNumber: 'XYZ789',
        sequence: 10
      },
      expected: 'KOBE_Printer_imageClass-MF644Cdw_XYZ789_010'
    },
    {
      name: '異常ケース（拠点未入力）',
      input: {
        locationCode: '',
        category: 'Server',
        model: 'ThinkPad',
        serialNumber: 'ABC123',
        sequence: 1
      },
      expected: null  // エラーが期待される
    }
  ];
  
  let passedCount = 0;
  
  testCases.forEach((testCase, index) => {
    try {
      const result = generateLocationNumber(testCase.input);
      
      if (testCase.expected === null) {
        console.log(`❌ Test ${index + 1} (${testCase.name}): 期待されたエラーが発生しませんでした`);
      } else if (result === testCase.expected) {
        console.log(`✅ Test ${index + 1} (${testCase.name}): PASSED`);
        passedCount++;
      } else {
        console.log(`❌ Test ${index + 1} (${testCase.name}): FAILED - Expected: ${testCase.expected}, Got: ${result}`);
      }
    } catch (error) {
      if (testCase.expected === null) {
        console.log(`✅ Test ${index + 1} (${testCase.name}): PASSED - Expected error: ${error.message}`);
        passedCount++;
      } else {
        console.log(`❌ Test ${index + 1} (${testCase.name}): FAILED - Unexpected error: ${error.message}`);
      }
    }
  });
  
  console.log(`=== テスト結果: ${passedCount}/${testCases.length} passed ===`);
}
```

### 4.2 統合テスト実装例
```javascript
// Code.gs（統合テスト）
/**
 * URL生成フロー統合テスト
 */
function integrationTestUrlGenerationFlow() {
  console.log('=== URL生成フロー統合テスト開始 ===');
  
  try {
    // 1. テストデータ準備
    const testLocationNumber = 'TEST_SV_TestModel_ABC123_001';
    
    // 2. 設定確認
    const settings = getCommonFormsSettings();
    if (!settings.success) {
      throw new Error('設定取得に失敗');
    }
    
    // 3. URL生成
    const urlResult = generateCommonFormUrl(testLocationNumber, 'server');
    if (!urlResult.success) {
      throw new Error('URL生成に失敗: ' + urlResult.error);
    }
    
    // 4. 生成URLの検証
    if (!urlResult.data.url.includes(testLocationNumber)) {
      throw new Error('生成URLに拠点管理番号が含まれていません');
    }
    
    console.log('✅ URL生成フロー統合テスト: 全て成功');
    console.log('Generated URL:', urlResult.data.url);
    
  } catch (error) {
    console.log('❌ URL生成フロー統合テスト: 失敗');
    console.log('Error:', error.message);
  }
}
```

## 5. デプロイ・運用

### 5.1 デプロイ手順

#### 5.1.1 テストデプロイ
```
1. Google Apps Script Editor で「デプロイ」→「新しいデプロイ」を選択
2. 種類で「ウェブアプリ」を選択
3. 説明: 「代替機管理システム v1.0.0 (Test)」
4. 実行者: 「自分」
5. アクセスできるユーザー: 「自分のみ」
6. 「デプロイ」をクリック
7. 生成されたURLでテスト実行
```

#### 5.1.2 本番デプロイ
```
1. テストで問題がないことを確認
2. 「デプロイ」→「新しいデプロイ」を選択
3. 説明: 「代替機管理システム v1.0.0 (Production)」
4. 実行者: 「アクセスしているユーザー」
5. アクセスできるユーザー: 「組織内の全員」
6. 「デプロイ」をクリック
7. 本番URLを関係者に共有
```

### 5.2 監視・保守

#### 5.2.1 ログ監視
```javascript
// 定期実行用：システムログ確認
function checkSystemHealth() {
  try {
    // 基本機能の動作確認
    const locationTest = getLocationMaster();
    const settingsTest = getCommonFormsSettings();
    
    if (!locationTest.success || !settingsTest.success) {
      // 管理者に通知（メール等）
      sendAdminAlert('システムヘルスチェック失敗', {
        locationTest: locationTest.success,
        settingsTest: settingsTest.success
      });
    }
    
    addLog('システムヘルスチェック完了', {
      locationTest: locationTest.success,
      settingsTest: settingsTest.success,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('システムヘルスチェック エラー:', error);
    sendAdminAlert('システムヘルスチェック エラー', error.message);
  }
}

// 管理者アラート送信
function sendAdminAlert(subject, details) {
  const adminEmail = 'admin@example.com';  // テスト用管理者メールアドレス
  
  const body = `
代替機管理システムでアラートが発生しました。

件名: ${subject}
詳細: ${JSON.stringify(details, null, 2)}
時刻: ${new Date().toLocaleString('ja-JP')}
URL: ${ScriptApp.getService().getUrl()}

システム管理者にて確認をお願いします。
  `;
  
  try {
    GmailApp.sendEmail(adminEmail, `[システムアラート] ${subject}`, body);
  } catch (error) {
    console.error('アラートメール送信失敗:', error);
  }
}
```

### 5.3 バックアップ・復旧

#### 5.3.1 データバックアップ
```javascript
// 定期実行用：データバックアップ
function createDataBackup() {
  try {
    const mainSpreadsheet = getMainSpreadsheet();
    const backupName = `代替機管理システム_バックアップ_${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss')}`;
    
    // スプレッドシートをコピー
    const backupSpreadsheet = mainSpreadsheet.copy(backupName);
    
    // バックアップフォルダに移動（フォルダIDを設定）
    const backupFolderId = 'YOUR_BACKUP_FOLDER_ID';  // 実際のフォルダIDに置き換え
    if (backupFolderId) {
      const backupFolder = DriveApp.getFolderById(backupFolderId);
      const backupFile = DriveApp.getFileById(backupSpreadsheet.getId());
      backupFolder.addFile(backupFile);
**定期メンテナンス**:
- 日次バックアップの自動実行
- 30日以上古いバックアップの自動削除
- バックアップ状態のログ記録
- エラー時の管理者アラート

## 6. トラブルシューティング

### 6.1 よくある問題と解決方法

| 問題 | 原因 | 解決方法 |
|------|------|----------|
| URLが生成されない | 共通フォームURL未設定 | 設定画面でURLを設定 |
| スプレッドシート読み込み失敗 | ID設定ミス | SPREADSHEET_ID_MAINを確認 |
| 権限エラー | アクセス権限不足 | 組織管理者に権限付与を依頼 |
| 処理が遅い | 大量データ処理 | バッチサイズを調整 |
| エラーメッセージが表示されない | JavaScript無効 | ブラウザ設定を確認 |

### 6.2 緊急時対応手順

**システム緊急停止**:
- PropertiesServiceでEMERGENCY_STOPフラグを設定
- 全API関数での緊急停止チェック
- ユーザーへのメンテナンスメッセージ表示

**システム復旧**:
- 緊急停止フラグの削除
- システム機能の正常化
- 復旧後の動作確認

このガイドに従って実装することで、堅牢で保守しやすい代替機管理システムが構築できます。段階的な実装とテストを重視し、品質の高いシステムを目指してください。

## 7. 実装チェックリスト

### 7.1 基盤機能
- [ ] Google Apps Scriptプロジェクトの作成
- [ ] メインスプレッドシートの作成
- [ ] PropertiesServiceの初期設定
- [ ] include()関数の実装
- [ ] doGet()関数の実装

### 7.2 マスタ管理機能
- [ ] 拠点マスタのCRUD操作
- [ ] 機種マスタのCRUD操作
- [ ] データタイプマスタの管理
- [ ] バリデーション機能

### 7.3 URL生成機能
- [ ] 拠点管理番号生成ロジック
- [ ] 共通フォームURL生成
- [ ] QRコードURL生成
- [ ] マスタデータ保存

### 7.4 統合ビュー・リアルタイム更新
- [ ] 端末系統合ビューシートの作成
- [ ] プリンタ・その他系統合ビューシートの作成
- [ ] 検索インデックスシートの作成
- [ ] onFormSubmitトリガーの設定
- [ ] onChangeトリガーの設定
- [ ] timeBasedトリガーの設定（深夜2:00）
- [ ] updateIntegratedViewOnSubmit関数の実装
- [ ] updateIntegratedViewOnChange関数の実装
- [ ] rebuildAllIntegratedViews関数の実装
- [ ] 部分更新ロジックの実装
- [ ] エラーハンドリング・リトライ機能
- [ ] 更新ログ・監視機能

### 7.5 フロントエンド
- [ ] メインページの構築
- [ ] ページ遷移機能
- [ ] ユーザーインターフェース
- [ ] レスポンシブデザイン

### 7.6 テスト・品質管理
- [ ] ユニットテストの実装
- [ ] 統合テストの実行
- [ ] エラーハンドリングの確認
- [ ] パフォーマンステスト

### 7.7 デプロイ・運用
- [ ] テスト環境でのデプロイ
- [ ] 本番環境でのデプロイ
- [ ] 監視・ログ機能の設置
- [ ] バックアップ体制の構築