# ステータス駆動型動的列フィルタリング機能実装ガイド

## 概要

データ設計書（05_data_design.md）に基づいて、スプレッドシートビューアーに**ステータス駆動型動的列フィルタリング機能**を実装しました。

## 実装内容

### 1. 基本仕様

通常データ（NORMAL）表示時に、「0-4.ステータス」の値に応じて表示する列を動的にフィルタリングします。

#### フィルタリングルール

- **ステータスが「1.貸出中」の場合**
  - 「1-」で始まる列のみ表示（例：1-1.貸出先、1-2.貸出日など）
  
- **ステータスが「2.回収後社内保管」の場合**
  - 「2-」で始まる列のみ表示
  
- **ステータスが「3.社内にて保管中」の場合**
  - 「3-」で始まる列を表示
  - さらに「3-0.社内ステータス」の値を確認し、2段階フィルタリングを適用
  - 例：「3-0.社内ステータス」が「1.倉庫保管」なら「3-1-」で始まる列も表示

### 2. 実装詳細

#### 2.1 createTable関数の拡張

```javascript
// ステータス駆動型列フィルタリングのための処理
let statusPrefixFilter = null;
let nestedStatusFilter = null;

// 通常データ（NORMAL）の場合のみステータス駆動型フィルタリングを適用
const dataTypeId = document.getElementById('dataType') ? document.getElementById('dataType').value : '';
if (dataTypeId === 'NORMAL' && displayMode === 'main') {
  // 0-4.ステータス列を探す
  const statusColumnIndex = data[0].findIndex(header => header && header.includes('0-4.ステータス'));
  
  if (statusColumnIndex !== -1 && data.length > 1) {
    // データの中からステータス値を取得（代表的な値を使用）
    const sampleStatusValue = getSampleStatusValue(data, statusColumnIndex);
    
    if (sampleStatusValue) {
      // ステータスの接頭辞数字を抽出
      const match = sampleStatusValue.match(/^(\d+)\./); 
      if (match) {
        statusPrefixFilter = match[1];
        
        // 3.社内にて保管中の場合、ネストしたステータスを確認
        if (statusPrefixFilter === '3') {
          // 3-0.社内ステータスの処理
          // ...
        }
      }
    }
  }
}
```

#### 2.2 補助関数の追加

**getSampleStatusValue関数**
```javascript
function getSampleStatusValue(data, columnIndex) {
  // 最初の非空データを返す
  for (let i = 1; i < data.length && i < 10; i++) {
    if (data[i][columnIndex] && data[i][columnIndex].toString().trim() !== '') {
      return data[i][columnIndex].toString().trim();
    }
  }
  return null;
}
```

**shouldShowColumn関数**
```javascript
function shouldShowColumn(columnHeader, statusPrefix, nestedStatusPrefix) {
  if (!columnHeader) return true;
  
  const headerStr = columnHeader.toString();
  
  // 基本列は常に表示
  const baseColumns = [
    '拠点管理番号',
    '機器種別',
    '機種名',
    '製造番号',
    'タイムスタンプ',
    '0-1.担当者',
    '0-4.ステータス',
    '更新日時',
    '更新日',
    '作成日'
  ];
  
  // 基本列は常に表示
  for (const baseCol of baseColumns) {
    if (headerStr.includes(baseCol)) {
      return true;
    }
  }
  
  // 数字-数字.パターンの列をチェック
  const match = headerStr.match(/^(\d+)-(\d+)\./); 
  if (match) {
    const prefix = match[1];
    const subPrefix = match[2];
    
    // ステータスの接頭辞と一致するかチェック
    if (prefix !== statusPrefix) {
      return false;
    }
    
    // 3.社内にて保管中の場合のネスト処理
    if (statusPrefix === '3' && nestedStatusPrefix && subPrefix !== '0') {
      return subPrefix === nestedStatusPrefix;
    }
    
    return true;
  }
  
  // その他の列は表示
  return true;
}
```

### 3. 適用箇所

#### 3.1 ヘッダー行の処理

```javascript
data[0].forEach((cell, index) => {
  // 空の列は表示しない
  if (isColumnEmpty(index)) {
    return;
  }
  
  // ステータス駆動型フィルタリング
  if (statusPrefixFilter && !shouldShowColumn(cell, statusPrefixFilter, nestedStatusFilter)) {
    return;
  }
  
  table += `<th>${cell}</th>`;
});
```

#### 3.2 データ行の処理

```javascript
data[i].forEach((cell, index) => {
  // 空の列は表示しない
  if (isColumnEmpty(index)) {
    return;
  }
  
  // ステータス駆動型フィルタリング
  if (statusPrefixFilter && !shouldShowColumn(data[0][index], statusPrefixFilter, nestedStatusFilter)) {
    return;
  }
  
  // セルの内容を表示
  // ...
});
```

### 4. デバッグ情報

実装にはデバッグ情報を含めており、以下の情報が確認できます：

- ステータス列のインデックス
- サンプルステータス値
- 適用されたフィルタープレフィックス
- ネストステータスの値（3.社内にて保管中の場合）

### 5. 動作確認方法

1. スプレッドシートビューアーを開く
2. データタイプで「通常データ」を選択
3. 拠点を選択
4. データが表示されたら、以下を確認：
   - 「0-4.ステータス」が「1.貸出中」の行では、「1-」で始まる列のみ表示
   - 「0-4.ステータス」が「3.社内にて保管中」の行では、「3-」で始まる列が表示され、さらに「3-0.社内ステータス」の値に応じて詳細列が表示

### 6. 注意事項

- この機能は**通常データ（NORMAL）**の場合のみ有効です
- 監査データやサマリーデータでは適用されません
- 基本列（拠点管理番号、機種名など）は常に表示されます
- 空の列は従来通り非表示になります

### 7. 今後の拡張可能性

- フィルタリングルールのカスタマイズ
- 複数ステータスの同時表示
- ユーザーによるフィルタリング設定の保存