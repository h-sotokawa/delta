# Google SitesでのQRコード中間ページ実装ガイド

## Google Sitesの制限事項

1. **静的コンテンツのみ**: Google Sitesは基本的に静的なウェブサイトホスティング
2. **URLパラメータ処理不可**: 直接的なURLパラメータの読み取りは不可能
3. **サーバーサイド処理なし**: PHPやNode.jsなどのサーバーサイド言語は使用不可
4. **JavaScriptの制限**: カスタムJavaScriptの実行に制限あり

## 実装可能な解決策

### 方法1: Google Apps Script Web Appの埋め込み（推奨）

Google SitesにGoogle Apps ScriptのWeb Appを埋め込むことで、動的な機能を実現できます。

#### 実装手順

1. **Google Apps Script Web Appの準備**
   - 既存の`gas-qr-redirect-optimized.gs`を使用
   - デプロイ時に「アクセスできるユーザー: 全員」に設定

2. **Google Sitesでページを作成**
   - 新しいページを作成（例：「QRコードリダイレクト」）
   - URLは`/qr-redirect`などに設定

3. **埋め込みコードの追加**
   - ページ編集モードで「埋め込み」→「埋め込みコード」を選択
   - 以下のコードを貼り付け：

```html
<iframe 
  src="https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec" 
  width="100%" 
  height="600" 
  frameborder="0"
  style="border: none;">
</iframe>
```

### 方法2: 直接リンク方式

QRコードに直接Google FormsのURLを埋め込む方法です。

#### メリット
- 最速（中間ページなし）
- 実装が簡単

#### デメリット
- エラーハンドリング不可
- ログ記録不可
- 拠点管理番号の検証不可

## 推奨実装: ハイブリッドアプローチ

### 1. メインシステム（Google Apps Script）での設定追加

```javascript
// QRコード生成時のオプション
function generateQRCodeUrl(locationNumber, options = {}) {
  const settings = getCommonFormsSettings();
  
  // 直接フォームURLか中間ページURLかを選択
  if (options.directLink) {
    // 直接Google Formsへのリンク
    const category = getCategoryFromLocationNumber(locationNumber);
    const formType = (category === 'printer' || category === 'other') ? 'printer' : 'terminal';
    const form = forms[formType];
    return `${form.formUrl}?${form.entryField}=${encodeURIComponent(locationNumber)}`;
  } else {
    // 中間ページ経由
    return `${settings.qrRedirectUrl}?id=${encodeURIComponent(locationNumber)}`;
  }
}
```

### 2. Google Sitesでの説明ページ

Google Sitesに以下の内容を含む説明ページを作成：

1. **QRコード読み取り手順**
2. **トラブルシューティング**
3. **よくある質問**
4. **問い合わせ先**

### 3. エラー時のフォールバック

QRコードラベルに以下を印刷：
- QRコード
- 拠点管理番号（テキスト）
- 手動入力用URL（短縮URL）

## 実装例

### Google Apps Script（改良版）

```javascript
// より高速なリダイレクト実装
function doGet(e) {
  const locationNumber = e.parameter.id || '';
  
  if (!locationNumber) {
    // エラーページを表示
    return HtmlService.createHtmlOutput(getErrorPage())
      .setTitle('エラー - 代替機管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // フォームURLを構築
  const redirectUrl = buildFormUrl(locationNumber);
  
  // 302リダイレクトを試みる
  return HtmlService.createHtmlOutput()
    .addMetaTag('http-equiv', 'refresh', '0; url=' + redirectUrl)
    .setContent('<script>window.location.href="' + redirectUrl + '";</script>');
}
```

### Google Sitesページ構成

```
代替機管理システム
├── ホーム
├── QRコード読み取り
│   └── [埋め込み: Gas Web App]
├── 使い方ガイド
├── トラブルシューティング
└── お問い合わせ
```

## 最適なソリューション

1. **プライマリ**: Google Apps Script Web Appを使用（現在の実装）
2. **セカンダリ**: Google Sitesに説明ページを作成
3. **バックアップ**: QRコードラベルに手動入力情報を記載

この組み合わせにより、最も信頼性の高いシステムを構築できます。