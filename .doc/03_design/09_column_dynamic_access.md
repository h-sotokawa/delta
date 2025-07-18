# 列名動的アクセス設計書

## 1. 概要

Google Sheetsの列名を動的に取得・アクセスする仕組みを提供し、列構造の変更に柔軟に対応可能なシステムを実現する。

## 2. 設計原則

### 2.1 基本原則

1. **列名ハードコーディングの排除**: 列インデックスを直接指定せず、列名から動的に取得
2. **共通ヘルパー関数の使用**: 全システムで統一された方法での列アクセス
3. **エラー耐性**: 列が見つからない場合の安全な処理
4. **パフォーマンス考慮**: 必要最小限のシート読み込み

### 2.2 命名規則

- ヘッダー行は常に1行目に配置
- 列名は一意であることを前提
- 同じ意味を持つ列でも異なるシートでは異なる名前を使用可能

## 3. ヘルパー関数

### 3.1 基本関数

#### 3.1.1 getColumnIndex
```javascript
/**
 * 列インデックスを動的に取得するヘルパー関数
 * @param {Array} headers - ヘッダー行の配列
 * @param {string} columnName - 検索する列名
 * @return {number} 列インデックス（見つからない場合は-1）
 */
function getColumnIndex(headers, columnName) {
  return headers.indexOf(columnName);
}
```

#### 3.1.2 getColumnNumber
```javascript
/**
 * 列番号を動的に取得するヘルパー関数（1ベース）
 * @param {Array} headers - ヘッダー行の配列
 * @param {string} columnName - 検索する列名
 * @return {number} 列番号（見つからない場合は0）
 */
function getColumnNumber(headers, columnName) {
  const index = getColumnIndex(headers, columnName);
  return index >= 0 ? index + 1 : 0;
}
```

#### 3.1.3 getValueByColumnName
```javascript
/**
 * 配列から値を安全に取得するヘルパー関数
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー行の配列
 * @param {string} columnName - 取得する列名
 * @return {*} 値（見つからない場合は空文字）
 */
function getValueByColumnName(row, headers, columnName) {
  const index = getColumnIndex(headers, columnName);
  return index >= 0 ? (row[index] || '') : '';
}
```

### 3.2 拡張関数

#### 3.2.1 getColumnIndexMultiple
```javascript
/**
 * 複数の列名から最初に見つかった列のインデックスを取得
 * @param {Array} headers - ヘッダー行の配列
 * @param {Array<string>} columnNames - 検索する列名の配列（優先順）
 * @return {number} 列インデックス（見つからない場合は-1）
 */
function getColumnIndexMultiple(headers, columnNames) {
  for (const columnName of columnNames) {
    const index = getColumnIndex(headers, columnName);
    if (index >= 0) return index;
  }
  return -1;
}
```

#### 3.2.2 rowToObject
```javascript
/**
 * 行データをオブジェクトに変換
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー行の配列
 * @return {Object} ヘッダーをキーとしたオブジェクト
 */
function rowToObject(row, headers) {
  const obj = {};
  headers.forEach((header, index) => {
    obj[header] = row[index] || '';
  });
  return obj;
}
```

## 4. 実装パターン

### 4.1 基本的な使用方法

```javascript
// シートのヘッダー行を取得
const sheet = SpreadsheetApp.getActiveSheet();
const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

// 特定の列のインデックスを取得
const statusCol = getColumnIndex(headers, 'ステータス');
const dateCol = getColumnIndex(headers, '更新日時');

// データ行から値を取得
const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
const data = dataRange.getValues();

data.forEach(row => {
  const status = getValueByColumnName(row, headers, 'ステータス');
  const updateDate = getValueByColumnName(row, headers, '更新日時');
  // 処理...
});
```

### 4.2 複数の可能性がある列名への対応

```javascript
// 統合ビューとマスタシートで異なる列名を使用している場合
const statusIndex = getColumnIndexMultiple(headers, [
  '現在ステータス',  // 統合ビューでの列名
  '0-4.ステータス',   // 収集シートでの列名
  'ステータス'        // マスタシートでの列名
]);
```

### 4.3 データ検証での使用

```javascript
function validateData(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  const errors = [];
  
  data.forEach((row, index) => {
    // 必須フィールドの検証
    const managementNumber = getValueByColumnName(row, headers, '拠点管理番号');
    if (!managementNumber) {
      errors.push(`行${index + 2}: 拠点管理番号が未入力です`);
    }
    
    // ステータスの検証
    const status = getValueByColumnName(row, headers, 'ステータス');
    if (!isValidStatus(status)) {
      errors.push(`行${index + 2}: 無効なステータス値です`);
    }
  });
  
  return errors;
}
```

## 5. 移行ガイドライン

### 5.1 既存コードの移行

#### Before（ハードコーディング）
```javascript
const statusCol = 5;  // ステータス列は5列目
const status = row[4];  // 0ベースインデックス
```

#### After（動的取得）
```javascript
const statusCol = getColumnNumber(headers, 'ステータス');
const status = getValueByColumnName(row, headers, 'ステータス');
```

### 5.2 チェックリスト

- [ ] ヘッダー行の取得処理を追加
- [ ] ハードコーディングされた列インデックスを特定
- [ ] getColumnIndex/getColumnNumber関数で置き換え
- [ ] 配列アクセスをgetValueByColumnName関数で置き換え
- [ ] エラーハンドリングの追加（列が見つからない場合）

## 6. パフォーマンス考慮事項

### 6.1 ヘッダー行のキャッシュ

```javascript
// グローバル変数でキャッシュ
let headerCache = {};

function getCachedHeaders(sheetName) {
  if (!headerCache[sheetName]) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    headerCache[sheetName] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  return headerCache[sheetName];
}
```

### 6.2 バッチ処理での考慮

```javascript
// 一度の読み込みで複数の列インデックスを取得
const columnIndices = {
  managementNumber: getColumnIndex(headers, '拠点管理番号'),
  category: getColumnIndex(headers, 'カテゴリ'),
  status: getColumnIndex(headers, 'ステータス'),
  updateDate: getColumnIndex(headers, '更新日時')
};

// 後続の処理で使用
data.forEach(row => {
  const managementNumber = row[columnIndices.managementNumber];
  const category = row[columnIndices.category];
  // ...
});
```

## 7. トラブルシューティング

### 7.1 一般的な問題と解決策

| 問題 | 原因 | 解決策 |
|------|------|--------|
| 列が見つからない | 列名の誤字・変更 | getColumnIndexMultiple使用で複数の可能性に対応 |
| パフォーマンス低下 | 毎回ヘッダー取得 | ヘッダーをキャッシュ |
| 空白セルでエラー | null/undefinedアクセス | getValueByColumnNameで安全な取得 |

### 7.2 デバッグ方法

```javascript
// 列名と位置の確認
function debugColumns(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  console.log(`シート: ${sheetName}`);
  headers.forEach((header, index) => {
    console.log(`列${index + 1}: ${header}`);
  });
}
```

## 8. 今後の拡張

### 8.1 検討事項

- 列名のエイリアス機能
- 複数ヘッダー行への対応
- 列グループの概念導入
- 型安全性の向上（TypeScript対応）

### 8.2 ベストプラクティス

1. **列名の標準化**: システム全体で統一された列名規則
2. **ドキュメント化**: 各シートの列構造をドキュメント化
3. **テスト**: 列構造変更時の回帰テスト
4. **監視**: 列構造の変更を検知する仕組み