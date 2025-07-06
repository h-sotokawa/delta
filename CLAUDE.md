# ガイドライン

このドキュメントは、プロジェクトのルール、目標、進捗管理方法を定義しています。以下の内容に従ってプロジェクトを進めてください。

## トップレベルルール

- このドキュメントは日本語で記述します。コードのコメントも日本語で記述してください。
- 効率を最大化するために、**複数の独立したプロセスを実行する必要がある場合は、ツールを順次ではなく並行して呼び出してください**。
- **思考は英語のみで行ってください**。ただし、**応答は日本語で行う必要があります**。
- ライブラリの使用方法を理解するには、**常にContex7 MCPを使用して最新情報を取得してください**。

## 日時処理の標準仕様

### 基本原則
このシステムでは、Google Sheetsへの日時保存において**書式なしテキスト形式**を採用しています。これにより、Google Sheetsの自動日付変換による問題を回避します。
回答は必ず日本語で行います。

### 実装方法

#### 1. 日時の保存形式
```javascript
// ✅ 正しい実装
const today = "'" + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

// ❌ 避けるべき実装
const today = new Date(); // 自動変換される可能性
const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'); // アポストロフィなし
```

#### 2. 列の書式設定
```javascript
// 日付列のセル書式をテキストに設定
dateRange.setNumberFormat('@'); // '@' = テキスト形式
```

#### 3. 表示用の日時フォーマット関数
```javascript
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  try {
    // 既に文字列の場合（アポストロフィ付きの場合を含む）
    if (typeof dateValue === 'string') {
      return dateValue.startsWith("'") ? dateValue.substring(1) : dateValue;
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

### 適用箇所

#### 拠点マスタ (location-master)
- **新規追加**: `addLocation()` - 作成日・更新日、グループメールアドレス
- **データ更新**: `updateLocation()` - 更新日、グループメールアドレス
- **初期データ**: `getLocationMasterSheet()` - 既存拠点の初期化
- **表示**: `formatDate()` - yyyy/MM/dd形式で表示
- **グループメールアドレス**: 拠点ごとのグループメールアドレスを管理可能

#### 機種マスタ (model-master) 
- **新規追加**: `addModelMasterData()` - 作成日・更新日
- **データ更新**: `updateModelMasterData()` - 更新日
- **データ修正**: `fixExistingDateFormats()` - 既存データの書式修正
- **表示**: `formatDate()` - yyyy/MM/dd形式で表示

### 技術的詳細

#### なぜアポストロフィ（'）を使用するのか
1. **自動変換防止**: Google Sheetsは `'2024/07/02` を文字列として認識し、日付形式に自動変換しない
2. **データ整合性**: 入力された値がそのまま保持される
3. **表示制御**: フロントエンドで適切にフォーマットして表示可能

#### タイムゾーン設定
```javascript
// 日本時間（JST）で統一
'Asia/Tokyo'
```

#### 日付形式
```javascript
// 標準フォーマット
'yyyy/MM/dd' // 例: 2024/07/02
```

### 新機能実装時の注意点

1. **日時を扱う新しいマスタデータを作成する場合**
   - 必ずアポストロフィ付きテキスト形式で保存
   - 列の書式を `@` (テキスト) に設定
   - 表示用の `formatDate` 関数を実装

2. **既存データの移行が必要な場合**
   - `fixExistingDateFormats` のような修正関数を作成
   - バックアップを取ってから実行

3. **フロントエンド表示**
   - 常に `formatDate` 関数を通して表示
   - アポストロフィの除去処理を含める

### ファイル構成

#### バックエンド (Google Apps Script)
- `Code.gs` - メインロジック、日時保存処理

#### フロントエンド
- `*-functions.html` - 各マスタの機能、`formatDate`関数
- `*-styles.html` - スタイル定義
- `*.html` - UI定義

### デバッグ・確認方法

#### Google Sheetsでの確認
1. セルを選択して数式バーを確認
2. 正しく保存されていれば `'2024/07/02` のように表示される
3. セルの書式が「書式なしテキスト」になっていることを確認

#### ブラウザでの確認
```javascript
// コンソールで実行
console.log('日付データ:', location.createdAt);
// 期待値: "2024/07/02" または "'2024/07/02"
```

この仕様により、システム全体で一貫した日時処理が実現され、Google Sheetsの自動変換による問題を回避できます。