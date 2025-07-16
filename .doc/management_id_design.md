# 9999.管理ID生成システム設計書

## 1. 概要

### 1.1 目的
Google Formsから送信されるすべての回答に対して、完全に一意な管理IDを自動的に付与するシステムを設計します。この管理IDは、データの追跡、監査、変更管理において中核的な役割を果たします。

### 1.2 要件
- **一意性**: システム全体で絶対に重複しない
- **予測不可能性**: セキュリティのため、連番だけでない要素を含む
- **可読性**: 人間が識別・分類しやすい形式
- **ソート可能性**: 時系列での並び替えが可能
- **堅牢性**: 同時実行時も重複しない

## 2. ID形式設計

### 2.1 基本形式
```
[PREFIX]_[TIMESTAMP]_[LOCATION]_[RANDOM]_[CHECKSUM]
```

#### 例
```
TRM_20240702143022_OSAKA_A3F2_K7
PRT_20240702143023_KOBE_B5D1_M3
OTH_20240702143024_HIMEJI_C8E9_P2
```

### 2.2 各要素の詳細

#### PREFIX（3文字）
- **TRM**: 端末（Terminal）
- **PRT**: プリンタ（Printer）
- **OTH**: その他（Other）

#### TIMESTAMP（14桁）
- 形式: YYYYMMDDHHmmss
- タイムゾーン: Asia/Tokyo（JST）
- 例: 20240702143022（2024年7月2日14時30分22秒）

#### LOCATION（最大10文字）
- 拠点管理番号から抽出した拠点コード
- 大文字英数字
- 例: OSAKA, KOBE, HIMEJI

#### RANDOM（4文字）
- ランダム生成された英数字
- 大文字のみ使用（混同しやすい文字は除外: 0,O,1,I,L）
- 使用可能文字: A-H,J,K,M,N,P-Z,2-9

#### CHECKSUM（2文字）
- エラー検出用のチェックサム
- CRC-16の下位2桁を36進数で表現

## 3. 実装設計（ハイブリッド方式）

### 3.1 システム構成

```javascript
/**
 * ハイブリッド管理IDシステム
 * - メイン処理: onFormSubmitトリガー（リアルタイム）
 * - バックアップ処理: 定期実行（未処理データ補完）
 * - 監視機能: エラー検知と通知
 */
```

### 3.2 メイン処理（onFormSubmitトリガー）

```javascript
/**
 * フォーム送信時のメイン処理
 * @param {Object} e - フォーム送信イベント
 */
function onFormSubmit(e) {
  const startTime = new Date();
  
  try {
    // 1. イベントデータの取得
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    const sheetName = sheet.getName();
    const namedValues = e.namedValues;
    
    // 2. 基本情報の抽出
    const locationNumber = namedValues['0-0.拠点管理番号'][0];
    const timestamp = new Date(namedValues['タイムスタンプ'][0]);
    
    // 3. 機器種類の判定
    const formType = detectFormType(sheetName, locationNumber);
    
    // 4. 管理IDの生成
    const managementId = generateManagementId(formType, locationNumber, timestamp);
    
    // 5. 管理IDの書き込み（B列）
    sheet.getRange(row, 2).setValue(managementId);
    
    // 6. 成功ログの記録
    logSuccess(managementId, formType, locationNumber, startTime);
    
    // 7. 通知処理（必要に応じて）
    if (isNotificationRequired(formType, locationNumber)) {
      sendStatusChangeNotification(managementId, locationNumber);
    }
    
  } catch (error) {
    // エラーハンドリング
    handleFormSubmitError(e, error, startTime);
  }
}

/**
 * 機器種類の判定（2種類の収集シート対応）
 */
function detectFormType(sheetName, locationNumber) {
  // 方法1: シート名による判定（優先）
  if (sheetName === '端末ステータス収集') {
    // 端末シートの場合、拠点管理番号で詳細判定
    const detectedType = detectFormTypeFromLocationNumber(locationNumber);
    if (detectedType === 'terminal') {
      return 'terminal';
    } else {
      // 端末以外が混入している場合の警告
      console.warn('端末シートに非端末データが記録されました:', locationNumber);
      logDataInconsistency(sheetName, locationNumber, detectedType);
      return 'terminal'; // 強制的に端末として処理
    }
  } else if (sheetName === 'プリンタその他ステータス収集') {
    // プリンタその他シートの場合、拠点管理番号で詳細判定
    const detectedType = detectFormTypeFromLocationNumber(locationNumber);
    if (['printer', 'other'].includes(detectedType)) {
      return detectedType;
    } else {
      // 端末が混入している場合の警告
      console.warn('プリンタその他シートに端末データが記録されました:', locationNumber);
      logDataInconsistency(sheetName, locationNumber, detectedType);
      return 'other'; // 強制的にその他として処理
    }
  }
  
  // 方法2: 拠点管理番号による判定（フォールバック）
  return detectFormTypeFromLocationNumber(locationNumber);
}
```

### 3.3 バックアップ処理（定期実行）

```javascript
/**
 * 定期実行によるバックアップ処理
 * 1日1回実行し、未処理データを補完
 */
function dailyManagementIdCheck() {
  const startTime = new Date();
  console.log('定期管理IDチェック開始:', startTime);
  
  try {
    // 1. 未処理行の検索
    const unprocessedRows = findUnprocessedRows();
    
    if (unprocessedRows.length === 0) {
      console.log('未処理データなし');
      return;
    }
    
    console.log(`未処理データ検出: ${unprocessedRows.length}件`);
    
    // 2. 未処理データの処理
    const results = processUnassignedManagementIds(unprocessedRows);
    
    // 3. 結果の集計
    const summary = {
      total: unprocessedRows.length,
      success: results.filter(r => r.success).length,
      failed: results.filter(r => !r.success).length,
      processingTime: new Date() - startTime
    };
    
    // 4. 管理者への通知
    sendBackupProcessNotification(summary);
    
    // 5. 処理ログの記録
    logBackupProcessing(summary, results);
    
  } catch (error) {
    console.error('定期処理エラー:', error);
    sendErrorNotification('定期管理IDチェック', error);
  }
}

/**
 * 未処理行の検索
 */
function findUnprocessedRows() {
  const sheets = ['端末ステータス収集', 'プリンタその他ステータス収集'];
  const unprocessed = [];
  
  sheets.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return;
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行をスキップして処理
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const managementId = row[1]; // B列
      const timestamp = row[0]; // A列
      const locationNumber = row[2]; // C列
      
      // 管理IDが未設定で、かつ有効なデータの場合
      if ((!managementId || managementId === '') && 
          timestamp && locationNumber) {
        unprocessed.push({
          sheet: sheetName,
          row: i + 1,
          timestamp: new Date(timestamp),
          locationNumber: locationNumber
        });
      }
    }
  });
  
  return unprocessed;
}

/**
 * 未処理データの一括処理
 */
function processUnassignedManagementIds(unprocessedRows) {
  const results = [];
  
  unprocessedRows.forEach(item => {
    try {
      // 機器種類の判定
      const formType = detectFormType(item.sheet, item.locationNumber);
      
      // 管理IDの生成
      const managementId = generateManagementId(
        formType, 
        item.locationNumber, 
        item.timestamp
      );
      
      // スプレッドシートへの書き込み
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(item.sheet);
      sheet.getRange(item.row, 2).setValue(managementId);
      
      results.push({
        success: true,
        managementId: managementId,
        item: item
      });
      
    } catch (error) {
      console.error(`行${item.row}の処理エラー:`, error);
      results.push({
        success: false,
        error: error.toString(),
        item: item
      });
    }
  });
  
  return results;
}
```

### 3.4 管理IDコア生成関数

```javascript
/**
 * 管理IDを生成する
 * @param {string} formType - フォームタイプ (terminal/printer/other)
 * @param {string} locationNumber - 拠点管理番号
 * @param {Date} timestamp - タイムスタンプ（省略時は現在時刻）
 * @returns {string} 生成された管理ID
 */
function generateManagementId(formType, locationNumber, timestamp = new Date()) {
  // 1. プレフィックスの決定
  const prefix = getPrefix(formType);
  
  // 2. タイムスタンプの生成
  const timestampStr = formatTimestamp(timestamp);
  
  // 3. 拠点コードの抽出
  const locationCode = extractLocationCode(locationNumber);
  
  // 4. ランダム文字列の生成
  const randomStr = generateRandomString(4);
  
  // 5. チェックサムの計算
  const baseId = `${prefix}_${timestampStr}_${locationCode}_${randomStr}`;
  const checksum = calculateChecksum(baseId);
  
  // 6. 最終ID の組み立て
  const managementId = `${baseId}_${checksum}`;
  
  // 7. 重複チェック（オプション）
  if (isDuplicateId(managementId)) {
    // 重複の場合は再生成（最大3回まで）
    return regenerateId(formType, locationNumber, timestamp);
  }
  
  return managementId;
}
```

### 3.5 補助関数

#### プレフィックス決定関数
```javascript
function getPrefix(formType) {
  const prefixMap = {
    'terminal': 'TRM',
    'printer': 'PRT',
    'other': 'OTH'
  };
  
  return prefixMap[formType] || 'UNK'; // Unknown
}
```

#### タイムスタンプフォーマット関数
```javascript
function formatTimestamp(date) {
  // Asia/Tokyo タイムゾーンで変換
  const jstDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyyMMddHHmmss');
  return jstDate;
}
```

#### 拠点コード抽出関数
```javascript
function extractLocationCode(locationNumber) {
  // 拠点管理番号の最初のセクションを抽出
  // 例: "OSAKA_SV_ThinkPad_ABC123_001" → "OSAKA"
  const parts = locationNumber.split('_');
  const locationCode = parts[0] || 'UNKNOWN';
  
  // 最大10文字に制限
  return locationCode.substring(0, 10).toUpperCase();
}
```

#### ランダム文字列生成関数
```javascript
function generateRandomString(length) {
  // 混同しやすい文字を除外: 0,O,1,I,L
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let result = '';
  
  for (let i = 0; i < length; i++) {
    const randomIndex = Math.floor(Math.random() * chars.length);
    result += chars[randomIndex];
  }
  
  return result;
}
```

#### チェックサム計算関数
```javascript
function calculateChecksum(baseId) {
  // 簡易的なCRC風のチェックサム
  let sum = 0;
  
  for (let i = 0; i < baseId.length; i++) {
    sum = ((sum << 1) + baseId.charCodeAt(i)) % 1296; // 36^2
  }
  
  // 36進数に変換（0-9, A-Z）
  return sum.toString(36).toUpperCase().padStart(2, '0');
}
```

### 3.6 重複チェック機能

```javascript
/**
 * IDの重複をチェック
 * @param {string} managementId - チェックする管理ID
 * @returns {boolean} 重複している場合true
 */
function isDuplicateId(managementId) {
  // キャッシュを使用した高速チェック
  const cache = CacheService.getScriptCache();
  const cacheKey = `mgmt_id_${managementId}`;
  
  if (cache.get(cacheKey)) {
    return true;
  }
  
  // スプレッドシートでの確認（バッチ処理対応）
  const isDuplicate = checkIdInSpreadsheet(managementId);
  
  if (!isDuplicate) {
    // キャッシュに登録（24時間保持）
    cache.put(cacheKey, '1', 86400);
  }
  
  return isDuplicate;
}

/**
 * スプレッドシートでIDの存在確認
 */
function checkIdInSpreadsheet(managementId) {
  // 各収集シートのB列（管理ID列）を検索
  const sheets = [
    '端末ステータス収集',
    'プリンタステータス収集',
    'その他ステータス収集'
  ];
  
  for (const sheetName of sheets) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) continue;
    
    // 高速検索のためにgetValues()を使用
    const values = sheet.getRange('B:B').getValues();
    for (const row of values) {
      if (row[0] === managementId) {
        return true;
      }
    }
  }
  
  return false;
}
```

## 4. エラーハンドリングと監視

### 4.1 エラーハンドリング関数

```javascript
/**
 * フォーム送信時のエラーハンドリング
 */
function handleFormSubmitError(e, error, startTime) {
  const errorData = {
    timestamp: new Date(),
    error: error.toString(),
    stack: error.stack,
    eventData: {
      sheet: e.range.getSheet().getName(),
      row: e.range.getRow(),
      values: e.values
    },
    processingTime: new Date() - startTime
  };
  
  // エラーログの記録
  logError(errorData);
  
  // フォールバックID生成
  const fallbackId = generateFallbackId();
  
  try {
    // B列にフォールバックIDを設定
    e.range.getSheet().getRange(e.range.getRow(), 2).setValue(fallbackId);
    console.log('フォールバックID設定完了:', fallbackId);
  } catch (fallbackError) {
    console.error('フォールバックID設定失敗:', fallbackError);
  }
  
  // 管理者への緊急通知
  sendErrorNotification('フォーム送信処理', errorData);
}

/**
 * データ不整合の記録
 */
function logDataInconsistency(sheetName, locationNumber, detectedType) {
  const logEntry = {
    timestamp: new Date(),
    type: 'DATA_INCONSISTENCY',
    sheetName: sheetName,
    locationNumber: locationNumber,
    detectedType: detectedType,
    severity: 'WARNING'
  };
  
  // 不整合ログシートに記録
  const logSheet = getOrCreateLogSheet('データ不整合ログ');
  logSheet.appendRow([
    logEntry.timestamp,
    logEntry.type,
    logEntry.sheetName,
    logEntry.locationNumber,
    logEntry.detectedType,
    logEntry.severity
  ]);
}
```

### 4.2 通知システム

```javascript
/**
 * バックアップ処理完了通知
 */
function sendBackupProcessNotification(summary) {
  const subject = '[管理IDシステム] 定期バックアップ処理完了';
  const body = `
定期バックアップ処理が完了しました。

処理結果:
- 総件数: ${summary.total}件
- 成功: ${summary.success}件
- 失敗: ${summary.failed}件
- 処理時間: ${summary.processingTime}ms

${summary.failed > 0 ? '失敗した処理があります。ログを確認してください。' : ''}
  `;
  
  sendNotificationEmail(subject, body);
}

/**
 * エラー通知
 */
function sendErrorNotification(context, errorData) {
  const subject = `[管理IDシステム] エラー発生 - ${context}`;
  const body = `
エラーが発生しました。

コンテキスト: ${context}
発生時刻: ${errorData.timestamp}
エラー内容: ${errorData.error}

詳細な情報はログを確認してください。
  `;
  
  sendNotificationEmail(subject, body);
}
```

### 4.3 監視・メトリクス

```javascript
/**
 * システム状態の監視
 */
function monitorSystemHealth() {
  const metrics = {
    timestamp: new Date(),
    totalIds: getTotalGeneratedIds(),
    duplicateCount: getDuplicateCount(),
    errorRate: getErrorRate(),
    unprocessedCount: findUnprocessedRows().length,
    systemStatus: 'HEALTHY'
  };
  
  // 異常値の検出
  if (metrics.errorRate > 0.05) { // 5%以上のエラー率
    metrics.systemStatus = 'WARNING';
    sendHealthAlert('高いエラー率が検出されました');
  }
  
  if (metrics.unprocessedCount > 100) { // 100件以上の未処理データ
    metrics.systemStatus = 'WARNING';
    sendHealthAlert('大量の未処理データが検出されました');
  }
  
  // メトリクスの記録
  recordMetrics(metrics);
  
  return metrics;
}
```

## 5. トリガー設定とスケジューリング

### 5.1 トリガー設定

```javascript
/**
 * 初期設定：トリガーの設定
 */
function setupManagementIdTriggers() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (['onFormSubmit', 'dailyManagementIdCheck'].includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // onFormSubmitトリガーの設定
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  
  // 定期実行トリガーの設定（毎日深夜2時）
  ScriptApp.newTrigger('dailyManagementIdCheck')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  
  // 週次監視トリガーの設定（毎週日曜日）
  ScriptApp.newTrigger('monitorSystemHealth')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(8)
    .create();
  
  console.log('管理IDシステムのトリガー設定完了');
}
```

## 6. 旧セクションとの統合

### 6.1 設定管理

```javascript
/**
 * ハイブリッドシステムの設定
 */
const MANAGEMENT_ID_CONFIG = {
  // 収集シート設定（2種類対応）
  collectionSheets: [
    '端末ステータス収集',
    'プリンタその他ステータス収集'
  ],
  
  // 機器種類マッピング
  deviceTypeMapping: {
    '端末ステータス収集': {
      allowedTypes: ['terminal'],
      fallbackType: 'terminal'
    },
    'プリンタその他ステータス収集': {
      allowedTypes: ['printer', 'other'],
      fallbackType: 'other'
    }
  },
  
  // 拠点管理番号パターン
  locationNumberPatterns: {
    terminal: ['SV', 'CL', 'Desktop', 'Laptop', 'Server', 'Tablet'],
    printer: ['Printer'],
    other: ['Router', 'Hub', 'Other']
  },
  
  // 実行スケジュール
  schedule: {
    dailyCheck: { hour: 2, minute: 0 },
    weeklyMonitor: { weekday: 'SUNDAY', hour: 8 }
  },
  
  // 通知設定
  notifications: {
    enabled: true,
    recipients: ['admin@example.com'],
    errorThreshold: 0.05, // 5%
    unprocessedThreshold: 100
  }
};
```

### 6.2 従来の拠点管理番号抽出関数（2種類シート対応）

```javascript
/**
 * 拠点管理番号から機器種類を判定（2種類シート対応）
 */
function detectFormTypeFromLocationNumber(locationNumber) {
  const parts = locationNumber.split('_');
  const categorySection = parts[1]; // 2番目のセクション
  
  // 端末系の判定
  if (MANAGEMENT_ID_CONFIG.locationNumberPatterns.terminal.includes(categorySection)) {
    return 'terminal';
  }
  
  // プリンタの判定
  if (MANAGEMENT_ID_CONFIG.locationNumberPatterns.printer.includes(categorySection)) {
    return 'printer';
  }
  
  // その他の判定
  if (MANAGEMENT_ID_CONFIG.locationNumberPatterns.other.includes(categorySection)) {
    return 'other';
  }
  
  return 'unknown';
}
```

## 7. 運用・保守計画

### 7.1 段階的導入計画

1. **フェーズ1: 基盤構築**（1週間）
   - 管理ID生成システムの実装
   - onFormSubmitトリガーの設定
   - 基本的なエラーハンドリング

2. **フェーズ2: バックアップ機能**（1週間）
   - 定期実行システムの実装
   - 未処理データ検出機能
   - 通知システムの構築

3. **フェーズ3: 監視・運用**（1週間）
   - システム監視機能の実装
   - メトリクス収集
   - 運用ドキュメントの作成

4. **フェーズ4: 既存データ移行**（1週間）
   - 既存データへのバックフィル
   - データ整合性チェック
   - 本格運用開始

### 7.2 運用監視項目

- **日次監視**: 未処理データ件数、エラー発生率
- **週次監視**: システム全体の健全性、重複ID発生状況
- **月次監視**: ID生成パフォーマンス、容量使用状況

### 7.3 トラブルシューティング

主要なトラブルパターンと対処法：

1. **onFormSubmitトリガーが動作しない**
   → 定期実行による補完機能が自動的に対応

2. **管理ID生成でエラーが発生**
   → フォールバックIDが自動生成され、管理者に通知

3. **大量の未処理データが発生**
   → 定期実行での一括処理により解決

4. **重複IDの発生**
   → チェックサム機能による検出と再生成

この設計により、リアルタイム処理と確実性を両立したハイブリッド管理IDシステムが実現できます。

## 8. バックフィル機能

### 8.1 既存データへの管理ID付与
```javascript
/**
 * 既存のデータに管理IDを遡及的に付与
 * @param {string} sheetName - 対象シート名
 */
function backfillManagementIds(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行をスキップ
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const existingId = row[1]; // B列
    
    // 既に管理IDがある場合はスキップ
    if (existingId && existingId !== '') continue;
    
    // 必要な情報を抽出
    const timestamp = new Date(row[0]); // A列: タイムスタンプ
    const locationNumber = row[2]; // C列: 拠点管理番号
    const formType = detectFormTypeFromSheet(sheetName);
    
    // 管理IDを生成
    const managementId = generateManagementId(formType, locationNumber, timestamp);
    
    // シートに書き込み
    sheet.getRange(i + 1, 2).setValue(managementId);
    
    // 処理負荷軽減のための待機
    if (i % 100 === 0) {
      Utilities.sleep(1000);
    }
  }
}
```

## 9. パフォーマンス最適化

### 9.1 バッチ処理対応
```javascript
/**
 * 複数の管理IDを一括生成
 * @param {Array} requests - [{formType, locationNumber, timestamp}, ...]
 * @returns {Array} 生成された管理IDの配列
 */
function generateManagementIdsBatch(requests) {
  const results = [];
  const usedIds = new Set();
  
  for (const request of requests) {
    let managementId;
    let attempts = 0;
    
    do {
      managementId = generateManagementId(
        request.formType,
        request.locationNumber,
        request.timestamp
      );
      attempts++;
    } while (usedIds.has(managementId) && attempts < 3);
    
    usedIds.add(managementId);
    results.push(managementId);
  }
  
  return results;
}
```

### 9.2 キャッシュ戦略
```javascript
const ID_GENERATION_CONFIG = {
  // 最近生成されたIDのキャッシュ
  recentIdsCache: CacheService.getScriptCache(),
  
  // バッチサイズ
  batchSize: 50,
  
  // リトライ設定
  maxRetries: 3,
  retryDelay: 100, // ミリ秒
};
```

## 10. エラーハンドリング（従来仕様）

### 10.1 フォールバックID生成
```javascript
function generateFallbackId() {
  // 最低限の一意性を保証するフォールバックID
  const timestamp = new Date().getTime();
  const random = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `FALLBACK_${timestamp}_${random}`;
}
```

### 10.2 エラーログ
```javascript
function logManagementIdGeneration(managementId, formType, locationNumber, timestamp) {
  const logEntry = {
    timestamp: timestamp,
    managementId: managementId,
    formType: formType,
    locationNumber: locationNumber,
    status: 'SUCCESS'
  };
  
  // ログシートに記録
  const logSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('管理IDログ') || createLogSheet();
  
  logSheet.appendRow([
    timestamp,
    managementId,
    formType,
    locationNumber,
    'SUCCESS'
  ]);
}
```

## 11. 検証機能

### 11.1 ID形式検証
```javascript
/**
 * 管理IDの形式が正しいか検証
 * @param {string} managementId - 検証する管理ID
 * @returns {boolean} 形式が正しい場合true
 */
function validateManagementIdFormat(managementId) {
  // 正規表現パターン
  const pattern = /^[A-Z]{3}_\d{14}_[A-Z0-9]{1,10}_[A-Z0-9]{4}_[A-Z0-9]{2}$/;
  
  if (!pattern.test(managementId)) {
    return false;
  }
  
  // チェックサムの検証
  const parts = managementId.split('_');
  const checksum = parts[parts.length - 1];
  const baseId = parts.slice(0, -1).join('_');
  
  return calculateChecksum(baseId) === checksum;
}
```

### 11.2 一意性検証
```javascript
/**
 * システム全体でIDの一意性を検証
 * @returns {Object} 検証結果
 */
function validateSystemUniqueness() {
  const allIds = new Set();
  const duplicates = [];
  
  const sheets = [
    '端末ステータス収集',
    'プリンタステータス収集',
    'その他ステータス収集'
  ];
  
  for (const sheetName of sheets) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) continue;
    
    const ids = sheet.getRange('B2:B').getValues().flat().filter(id => id);
    
    for (const id of ids) {
      if (allIds.has(id)) {
        duplicates.push({
          id: id,
          sheet: sheetName
        });
      }
      allIds.add(id);
    }
  }
  
  return {
    totalIds: allIds.size,
    duplicates: duplicates,
    isValid: duplicates.length === 0
  };
}
```

## 12. 移行計画（従来仕様）

### 12.1 段階的導入
1. **フェーズ1**: 新規回答への管理ID付与開始
2. **フェーズ2**: 直近1ヶ月のデータへのバックフィル
3. **フェーズ3**: 全既存データへのバックフィル
4. **フェーズ4**: ビューシートへの管理ID組み込み

### 12.2 性能目標
- ID生成時間: 10ms以下
- 重複チェック時間: 50ms以下
- バッチ処理: 1000件/分

## 13. セキュリティ考慮事項

### 13.1 予測困難性
- ランダム要素により、次のIDを予測することは困難
- タイムスタンプだけでは特定不可能

### 13.2 改ざん検知
- チェックサムにより、IDの改ざんを検出可能
- 不正なIDは検証関数で排除

### 13.3 アクセス制御
- 管理ID生成関数は限定的なスコープで実行
- ログへのアクセスは管理者のみ