# インデックスシート設計書

## 1. 概要

### 1.1 目的
本設計書は、9999.管理IDを活用した高度な検索・履歴管理を実現するインデックスシートの設計仕様を定義します。管理IDにより、完全なデータトレーサビリティと高速検索を実現します。

### 1.2 基本方針
- **管理IDの活用**: 各データエントリを管理IDで一意に識別
- **高速検索**: インデックスによる効率的なデータアクセス
- **履歴管理**: 全ての変更履歴を管理IDで追跡可能
- **データ整合性**: 管理IDによる確実なデータ関連付け

## 2. インデックスシート構造

### 2.1 拡張検索インデックスシート（enhanced_search_index）

#### 概要
現在の最新状態を高速検索するためのインデックスシート。管理IDを含めることで、元データへの正確なアクセスを実現。

#### 列構造

| 列番号 | 列名         | データ型 | 説明                                      | 例                                                |
| ------ | ------------ | -------- | ----------------------------------------- | ------------------------------------------------- |
| A      | 拠点管理番号 | String   | プライマリキー                            | OSAKA_SV_ThinkPad_ABC123_001                      |
| B      | カテゴリ     | String   | 機器分類（インデックス化）                | Server                                            |
| C      | 機種名       | String   | 機種名（インデックス化）                  | ThinkPad X1 Carbon                                |
| D      | ステータス   | String   | 現在ステータス（インデックス化）          | 1.貸出中                                          |
| E      | 拠点コード   | String   | 拠点識別子（インデックス化）              | OSAKA                                             |
| F      | 管轄         | String   | 管轄エリア（インデックス化）              | 関西                                              |
| G      | 最終更新     | DateTime | ソート用日時                              | 2024/07/02 14:30:00                               |
| H      | 最新管理ID   | String   | 現在の状態を表す管理ID                    | TRM_20240702143022_OSAKA_A3F2_K7                 |
| I      | 検索キー     | String   | 全文検索用結合フィールド                  | OSAKA_SV_ThinkPad_ABC123_001 Server ThinkPad...  |

#### インデックス作成方法

```javascript
/**
 * 拡張検索インデックスの構築
 */
function buildEnhancedSearchIndex() {
  const sheets = ['端末ステータス収集', 'プリンタその他ステータス収集'];
  const indexSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('enhanced_search_index');
  
  const allIndexData = [];
  
  sheets.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return;
    
    // 最新データのみを取得（拠点管理番号でグループ化）
    const latestData = getLatestDataByLocationNumber(sheet);
    
    latestData.forEach(row => {
      const locationNumber = row[2]; // C列: 拠点管理番号
      const managementId = row[1];   // B列: 管理ID
      const timestamp = row[0];      // A列: タイムスタンプ
      const status = row[6];         // G列: ステータス
      
      // 拠点情報の取得
      const locationCode = locationNumber.split('_')[0];
      const category = detectCategoryFromLocationNumber(locationNumber);
      const modelName = extractModelName(locationNumber);
      const jurisdiction = getJurisdictionByLocationCode(locationCode);
      
      // インデックスデータの構築
      allIndexData.push([
        locationNumber,
        category,
        modelName,
        status,
        locationCode,
        jurisdiction,
        timestamp,
        managementId,
        `${locationNumber} ${category} ${modelName} ${status} ${jurisdiction} ${managementId}`
      ]);
    });
  });
  
  // インデックスシートへの書き込み
  indexSheet.clear();
  if (allIndexData.length > 0) {
    indexSheet.getRange(1, 1, allIndexData.length, 9).setValues(allIndexData);
  }
}
```

### 2.2 履歴検索インデックスシート（history_search_index）

#### 概要
全ての変更履歴を管理IDベースで検索可能にするインデックスシート。

#### 列構造

| 列番号 | 列名         | データ型 | 説明                       | 例                                |
| ------ | ------------ | -------- | -------------------------- | --------------------------------- |
| A      | 管理ID       | String   | プライマリキー             | TRM_20240702143022_OSAKA_A3F2_K7 |
| B      | 拠点管理番号 | String   | 機器識別子                 | OSAKA_SV_ThinkPad_ABC123_001      |
| C      | タイムスタンプ | DateTime | 記録日時                   | 2024/07/02 14:30:22               |
| D      | ステータス   | String   | その時点のステータス       | 1.貸出中                          |
| E      | 担当者       | String   | 操作担当者                 | 山田太郎                          |
| F      | 変更種別     | String   | 変更の種類                 | STATUS_CHANGE                     |
| G      | シート名     | String   | データソースシート         | 端末ステータス収集                |
| H      | 行番号       | Number   | 元データの行番号           | 1234                              |

#### インデックス作成方法

```javascript
/**
 * 履歴検索インデックスの構築
 */
function buildHistorySearchIndex() {
  const sheets = ['端末ステータス収集', 'プリンタその他ステータス収集'];
  const indexSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('history_search_index');
  
  const allHistoryData = [];
  
  sheets.forEach(sheetName => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    
    // ヘッダー行をスキップ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const managementId = row[1];  // B列: 管理ID
      
      // 管理IDが存在する行のみ処理
      if (managementId && managementId !== '') {
        const locationNumber = row[2]; // C列: 拠点管理番号
        const timestamp = row[0];      // A列: タイムスタンプ
        const status = row[6];         // G列: ステータス
        const manager = row[3];        // D列: 担当者
        
        allHistoryData.push([
          managementId,
          locationNumber,
          timestamp,
          status,
          manager,
          'STATUS_UPDATE', // 変更種別（今後拡張可能）
          sheetName,
          i + 1 // 行番号（1ベース）
        ]);
      }
    }
  });
  
  // タイムスタンプでソート（新しい順）
  allHistoryData.sort((a, b) => new Date(b[2]) - new Date(a[2]));
  
  // インデックスシートへの書き込み
  indexSheet.clear();
  if (allHistoryData.length > 0) {
    indexSheet.getRange(1, 1, allHistoryData.length, 8).setValues(allHistoryData);
  }
}
```

### 2.3 変更検知インデックスシート（change_detection_index）

#### 概要
機器ごとの最新管理IDと前回管理IDを記録し、変更検知を効率化するインデックスシート。

#### 列構造

| 列番号 | 列名           | データ型 | 説明                              | 例                                |
| ------ | -------------- | -------- | --------------------------------- | --------------------------------- |
| A      | 拠点管理番号   | String   | プライマリキー                    | OSAKA_SV_ThinkPad_ABC123_001      |
| B      | 最新管理ID     | String   | 現在の管理ID                      | TRM_20240702143022_OSAKA_A3F2_K7 |
| C      | 前回管理ID     | String   | 前回の管理ID                      | TRM_20240701120000_OSAKA_B2D1_M3 |
| D      | 変更検知フラグ | Boolean  | 前回から変更があったか            | TRUE                              |
| E      | 変更日時       | DateTime | 最新の変更日時                    | 2024/07/02 14:30:22               |
| F      | 変更内容       | String   | 変更の概要                        | ステータス変更: 3.社内保管→1.貸出中 |

#### 変更検知の実装

```javascript
/**
 * 変更検知インデックスの更新
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} newManagementId - 新しい管理ID
 */
function updateChangeDetectionIndex(locationNumber, newManagementId) {
  const indexSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('change_detection_index');
  
  // 既存のインデックスデータを取得
  const data = indexSheet.getDataRange().getValues();
  let rowIndex = -1;
  
  // 拠点管理番号で検索
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === locationNumber) {
      rowIndex = i;
      break;
    }
  }
  
  if (rowIndex === -1) {
    // 新規エントリ
    indexSheet.appendRow([
      locationNumber,
      newManagementId,
      '', // 前回管理IDなし
      false, // 初回なので変更なし
      new Date(),
      '初回登録'
    ]);
  } else {
    // 既存エントリの更新
    const currentManagementId = data[rowIndex][1];
    
    if (currentManagementId !== newManagementId) {
      // 変更を検知
      const changeDetails = detectChangeDetails(currentManagementId, newManagementId);
      
      indexSheet.getRange(rowIndex + 1, 2, 1, 5).setValues([[
        newManagementId,
        currentManagementId,
        true,
        new Date(),
        changeDetails
      ]]);
    }
  }
}
```

## 3. 管理IDを活用した検索機能

### 3.1 基本検索関数

```javascript
/**
 * 管理IDによる直接検索
 * @param {string} managementId - 検索する管理ID
 * @returns {Object} データオブジェクト
 */
function searchByManagementId(managementId) {
  // 履歴インデックスから検索
  const historyIndex = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('history_search_index');
  
  const data = historyIndex.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === managementId) {
      // 元データシートから詳細を取得
      const sheetName = data[i][6];
      const rowNumber = data[i][7];
      
      const sourceSheet = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(sheetName);
      const sourceData = sourceSheet.getRange(rowNumber, 1, 1, sourceSheet.getLastColumn())
        .getValues()[0];
      
      return {
        managementId: managementId,
        locationNumber: data[i][1],
        timestamp: data[i][2],
        status: data[i][3],
        manager: data[i][4],
        sourceSheet: sheetName,
        sourceRow: rowNumber,
        fullData: sourceData
      };
    }
  }
  
  return null;
}

/**
 * 拠点管理番号による履歴検索
 * @param {string} locationNumber - 拠点管理番号
 * @returns {Array} 履歴データの配列
 */
function searchHistoryByLocationNumber(locationNumber) {
  const historyIndex = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('history_search_index');
  
  const data = historyIndex.getDataRange().getValues();
  const history = [];
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === locationNumber) {
      history.push({
        managementId: data[i][0],
        timestamp: data[i][2],
        status: data[i][3],
        manager: data[i][4],
        changeType: data[i][5]
      });
    }
  }
  
  // タイムスタンプで降順ソート
  history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  
  return history;
}
```

### 3.2 高度な検索機能

```javascript
/**
 * 期間指定による変更履歴検索
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @returns {Array} 期間内の変更履歴
 */
function searchChangesByDateRange(startDate, endDate) {
  const historyIndex = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('history_search_index');
  
  const data = historyIndex.getDataRange().getValues();
  const changes = [];
  
  for (let i = 0; i < data.length; i++) {
    const timestamp = new Date(data[i][2]);
    
    if (timestamp >= startDate && timestamp <= endDate) {
      changes.push({
        managementId: data[i][0],
        locationNumber: data[i][1],
        timestamp: timestamp,
        status: data[i][3],
        manager: data[i][4]
      });
    }
  }
  
  return changes;
}

/**
 * ステータス変更の追跡
 * @param {string} locationNumber - 拠点管理番号
 * @returns {Array} ステータス変更履歴
 */
function trackStatusChanges(locationNumber) {
  const history = searchHistoryByLocationNumber(locationNumber);
  const statusChanges = [];
  
  for (let i = 0; i < history.length - 1; i++) {
    const current = history[i];
    const previous = history[i + 1];
    
    if (current.status !== previous.status) {
      statusChanges.push({
        managementId: current.managementId,
        timestamp: current.timestamp,
        fromStatus: previous.status,
        toStatus: current.status,
        manager: current.manager
      });
    }
  }
  
  return statusChanges;
}
```

## 4. インデックス更新戦略

### 4.1 更新タイミング

#### リアルタイム更新（推奨）
```javascript
/**
 * フォーム送信時のインデックス更新
 */
function onFormSubmit(e) {
  // 既存の管理ID生成処理
  // ...
  
  // インデックスの即時更新
  updateIndexesForNewEntry(managementId, locationNumber, e.range.getRow());
}

function updateIndexesForNewEntry(managementId, locationNumber, rowNumber) {
  // 拡張検索インデックスの更新
  updateEnhancedSearchIndex(locationNumber, managementId);
  
  // 履歴検索インデックスへの追加
  addToHistorySearchIndex(managementId, locationNumber, rowNumber);
  
  // 変更検知インデックスの更新
  updateChangeDetectionIndex(locationNumber, managementId);
}
```

#### バッチ更新（補完）
```javascript
/**
 * 定期的なインデックス再構築
 */
function rebuildAllIndexes() {
  console.log('インデックス再構築開始');
  
  // 各インデックスの再構築
  buildEnhancedSearchIndex();
  buildHistorySearchIndex();
  rebuildChangeDetectionIndex();
  
  console.log('インデックス再構築完了');
}

// トリガー設定（毎日深夜3時）
function setupIndexRebuildTrigger() {
  ScriptApp.newTrigger('rebuildAllIndexes')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}
```

### 4.2 パフォーマンス最適化

```javascript
/**
 * インデックスのインクリメンタル更新
 * 全体再構築ではなく、差分のみを更新
 */
function incrementalIndexUpdate() {
  const lastUpdateTime = getLastIndexUpdateTime();
  const newEntries = getEntriesSince(lastUpdateTime);
  
  if (newEntries.length > 0) {
    // 新規エントリのみをインデックスに追加
    appendToIndexes(newEntries);
    setLastIndexUpdateTime(new Date());
  }
}

/**
 * インデックスのキャッシュ活用
 */
const IndexCache = {
  cache: CacheService.getScriptCache(),
  
  get: function(key) {
    const cached = this.cache.get(key);
    return cached ? JSON.parse(cached) : null;
  },
  
  set: function(key, value, expirationInSeconds = 600) {
    this.cache.put(key, JSON.stringify(value), expirationInSeconds);
  },
  
  searchWithCache: function(managementId) {
    const cacheKey = `mgmt_search_${managementId}`;
    let result = this.get(cacheKey);
    
    if (!result) {
      result = searchByManagementId(managementId);
      if (result) {
        this.set(cacheKey, result);
      }
    }
    
    return result;
  }
};
```

## 5. データ整合性チェック

### 5.1 整合性検証関数

```javascript
/**
 * インデックスの整合性チェック
 */
function validateIndexIntegrity() {
  const issues = [];
  
  // 1. 管理IDの一意性チェック
  const duplicates = checkDuplicateManagementIds();
  if (duplicates.length > 0) {
    issues.push({
      type: 'DUPLICATE_IDS',
      count: duplicates.length,
      items: duplicates
    });
  }
  
  // 2. 欠損データのチェック
  const missingData = checkMissingIndexEntries();
  if (missingData.length > 0) {
    issues.push({
      type: 'MISSING_ENTRIES',
      count: missingData.length,
      items: missingData
    });
  }
  
  // 3. データ不整合のチェック
  const inconsistencies = checkDataConsistency();
  if (inconsistencies.length > 0) {
    issues.push({
      type: 'DATA_INCONSISTENCY',
      count: inconsistencies.length,
      items: inconsistencies
    });
  }
  
  return issues;
}

/**
 * 整合性問題の自動修復
 */
function repairIndexIssues(issues) {
  issues.forEach(issue => {
    switch (issue.type) {
      case 'DUPLICATE_IDS':
        removeDuplicateEntries(issue.items);
        break;
      case 'MISSING_ENTRIES':
        recreateMissingEntries(issue.items);
        break;
      case 'DATA_INCONSISTENCY':
        fixDataInconsistencies(issue.items);
        break;
    }
  });
  
  // 修復後の再検証
  return validateIndexIntegrity();
}
```

## 6. 運用ガイドライン

### 6.1 初期セットアップ

1. **インデックスシートの作成**
   ```javascript
   function createIndexSheets() {
     const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
     
     // 各インデックスシートを作成
     const sheets = [
       'enhanced_search_index',
       'history_search_index', 
       'change_detection_index'
     ];
     
     sheets.forEach(sheetName => {
       if (!spreadsheet.getSheetByName(sheetName)) {
         spreadsheet.insertSheet(sheetName);
       }
     });
   }
   ```

2. **初期インデックス構築**
   ```javascript
   function initialIndexBuild() {
     createIndexSheets();
     rebuildAllIndexes();
     setupIndexRebuildTrigger();
   }
   ```

### 6.2 メンテナンス

- **日次**: インクリメンタル更新の実行
- **週次**: 整合性チェックの実行
- **月次**: 完全再構築とパフォーマンス分析

### 6.3 トラブルシューティング

| 問題 | 原因 | 対処法 |
| ---- | ---- | ------ |
| 検索結果が不正確 | インデックスの不整合 | `rebuildAllIndexes()` を実行 |
| パフォーマンス低下 | インデックスサイズ過大 | 古いデータのアーカイブ |
| 管理ID重複エラー | 同時実行による競合 | 整合性チェックと修復を実行 |

## 7. 将来の拡張性

### 7.1 計画中の機能

1. **AI支援検索**
   - 自然言語による検索
   - 類似パターンの検出

2. **予測分析**
   - ステータス変更パターンの分析
   - 異常検知

3. **可視化ダッシュボード**
   - 変更履歴の可視化
   - トレンド分析

### 7.2 スケーラビリティ

- **分散インデックス**: データ量に応じてインデックスを分割
- **外部データベース連携**: 大規模データ対応
- **リアルタイム同期**: 複数システム間でのインデックス同期

この設計により、管理IDを中心とした高度なデータ管理とアクセス制御が実現できます。