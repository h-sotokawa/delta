# スプレッドシート操作関連関数の分析

## 関数一覧と位置

### 1. getSpreadsheetData
- **位置**: 222-317行
- **目的**: 指定された拠点とデバイスタイプのスプレッドシートデータを取得
- **パラメータ**: 
  - `location`: 拠点ID
  - `queryType`: クエリタイプ
  - `deviceType`: デバイスタイプ（デフォルト: 'terminal'）
- **依存関数**:
  - `convertLegacyLocation()`: 旧拠点名を新拠点名に変換
  - `getSpreadsheetIdFromProperty()`: スプレッドシートIDを取得
  - `getTargetSheetName()`: デバイスタイプに応じたシート名を決定
  - `formatDateFast()`: 日付フォーマット処理
  - `getLocationNameById()`: 拠点IDから拠点名を取得
  - `startPerformanceTimer()`, `endPerformanceTimer()`: パフォーマンス計測
  - `addLog()`: ログ記録

### 2. getSpreadsheetDataPaginated
- **位置**: 347-440行
- **目的**: ページネーション付きでスプレッドシートデータを取得
- **パラメータ**:
  - `location`: 拠点ID
  - `queryType`: クエリタイプ
  - `deviceType`: デバイスタイプ（デフォルト: 'terminal'）
  - `pageNumber`: ページ番号（デフォルト: 1）
  - `pageSize`: ページサイズ（デフォルト: 100）
- **依存関数**: getSpreadsheetDataと同様の依存関数を使用

### 3. updateMachineStatus
- **位置**: 441-508行
- **目的**: 機器のステータスを更新
- **パラメータ**:
  - `rowIndex`: 行インデックス
  - `newStatus`: 新しいステータス
  - `location`: 拠点
  - `deviceType`: デバイスタイプ（デフォルト: 'terminal'）
- **依存関数**:
  - `getSpreadsheetIdFromProperty()`: スプレッドシートIDを取得
  - `getTargetSheetName()`: デバイスタイプに応じたシート名を決定
  - `formatDateFast()`: 日付フォーマット処理

### 4. updateMultipleStatuses
- **位置**: 509-586行
- **目的**: 複数の機器のステータスを一括更新
- **パラメータ**:
  - `updates`: 更新情報の配列
  - `location`: 拠点
  - `deviceType`: デバイスタイプ（デフォルト: 'terminal'）
- **依存関数**: updateMachineStatusと同様

### 5. checkDataConsistency
- **位置**: 587-686行
- **目的**: データの整合性をチェック
- **パラメータ**:
  - `location`: 拠点
  - `deviceType`: デバイスタイプ（デフォルト: 'terminal'）
- **依存関数**:
  - `getSpreadsheetIdFromProperty()`: スプレッドシートIDを取得
  - `getTargetSheetName()`: デバイスタイプに応じたシート名を決定
  - `formatDateFast()`: 日付フォーマット処理

### 6. getLocationSheetData
- **位置**: 733-826行
- **目的**: 拠点別シートのデータを取得
- **パラメータ**:
  - `location`: 拠点
  - `locationSheetName`: 拠点シート名
  - `queryType`: クエリタイプ
- **依存関数**:
  - `getSpreadsheetIdFromProperty()`: スプレッドシートIDを取得
  - `formatDateFast()`: 日付フォーマット処理
  - `getLocationNameById()`: 拠点IDから拠点名を取得

### 7. updateLocationSheetCell
- **位置**: 827-904行
- **目的**: 拠点別シートのセルを更新
- **パラメータ**:
  - `location`: 拠点
  - `locationSheetName`: 拠点シート名
  - `rowIndex`: 行インデックス
  - `columnIndex`: 列インデックス
  - `newValue`: 新しい値
- **依存関数**:
  - `getSpreadsheetIdFromProperty()`: スプレッドシートIDを取得
  - `formatDateFast()`: 日付フォーマット処理

### 8. getDestinationSheetData
- **位置**: 905-1111行
- **目的**: 統一スプレッドシートのデータを取得
- **パラメータ**:
  - `sheetName`: シート名
  - `queryType`: クエリタイプ
- **依存関数**:
  - `getUnifiedSpreadsheetId()`: 統一スプレッドシートIDを取得
  - `formatDateFast()`: 日付フォーマット処理
  - その他多数のヘルパー関数

### 9. getDestinationSheets
- **位置**: 1127-1161行
- **目的**: 統一スプレッドシートのシート一覧を取得
- **パラメータ**: なし
- **依存関数**:
  - `getUnifiedSpreadsheetId()`: 統一スプレッドシートIDを取得

## 共通依存関数

### ユーティリティ関数（config.gsで定義）
- `startPerformanceTimer()`: パフォーマンス計測開始（224-229行）
- `endPerformanceTimer()`: パフォーマンス計測終了（236-242行）
- `debugLog()`: デバッグログ出力（213-217行）

### ユーティリティ関数（Code.gsで定義）
- `addLog()`: ログ記録（未確認）
- `formatDateFast()`: 高速日付フォーマット（209-220行）

### データアクセス関数（Code.gsで定義）
- `getSpreadsheetIdFromProperty()`: 拠点別スプレッドシートID取得（**注意: 関数定義が見つからない**）
- `getUnifiedSpreadsheetId()`: 統一スプレッドシートID取得（**注意: 関数定義が見つからない**）
- `getTargetSheetName()`: デバイスタイプに応じたシート名取得（38-49行）
- `getTargetSheetNameByCategory()`: カテゴリに応じたシート名取得（51-72行）
- `getLocationNameById()`: 拠点IDから拠点名を取得（6-36行）
- `convertLegacyLocation()`: 旧拠点名を新拠点名に変換（192-206行）

## 機能グループ

### 1. データ取得系
- `getSpreadsheetData`: 基本的なデータ取得
- `getSpreadsheetDataPaginated`: ページネーション付きデータ取得
- `getLocationSheetData`: 拠点別シートデータ取得
- `getDestinationSheetData`: 統一スプレッドシートデータ取得
- `getDestinationSheets`: シート一覧取得

### 2. データ更新系
- `updateMachineStatus`: 単一機器のステータス更新
- `updateMultipleStatuses`: 複数機器のステータス一括更新
- `updateLocationSheetCell`: 拠点別シートのセル更新

### 3. データ検証系
- `checkDataConsistency`: データ整合性チェック

## 注意事項

### 未定義関数の問題
以下の関数が呼び出されているが、定義が見つからない：
1. `getSpreadsheetIdFromProperty()` - 10箇所で使用
2. `getUnifiedSpreadsheetId()` - getDestinationSheetData関数で使用
3. `addLog()` - 複数箇所で使用

これらの関数は以下の可能性がある：
- 別ファイルで定義されている
- PropertiesServiceを直接使用するようにリファクタリングされた
- 実装が未完了

## エラーハンドリング

すべての関数で共通のエラーハンドリングパターンを使用:
```javascript
try {
  // 処理
  return {
    success: true,
    data: data,
    logs: DEBUG ? serverLogs : []
  };
} catch (error) {
  return {
    success: false,
    error: error.toString(),
    errorDetails: {
      message: error.message,
      stack: error.stack,
      name: error.name
    },
    logs: DEBUG ? serverLogs : []
  };
}
```

## パフォーマンス最適化

1. **バッチ処理**: `updateMultipleStatuses`では複数の更新を一括処理
2. **キャッシュ利用**: スプレッドシートIDなどはプロパティから取得
3. **高速日付処理**: `formatDateFast`関数による最適化
4. **ページネーション**: 大量データの分割取得対応

## 関数間の依存関係

### 直接的な依存関係
1. `getSpreadsheetDataPaginated` → `getSpreadsheetData` を内部で使用している可能性
2. `updateMultipleStatuses` → 単一更新のロジックを再利用
3. `getDestinationSheetData` → 統一スプレッドシートにアクセス

### 共通依存
- すべての関数が `formatDateFast` を使用して日付フォーマットを統一
- パフォーマンス計測のための共通フレームワーク
- エラーハンドリングの一貫性

## 今後の改善提案

1. **未定義関数の実装**
   - `getSpreadsheetIdFromProperty()`
   - `getUnifiedSpreadsheetId()`
   - `addLog()`

2. **関数の分割**
   - 特に`getDestinationSheetData`（206行）は大きすぎるため、機能別に分割を推奨

3. **テストコードの追加**
   - 各関数の単体テスト
   - 統合テスト

4. **ドキュメントの充実**
   - JSDocコメントの追加
   - 使用例の明記