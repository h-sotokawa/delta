# ナビゲーション・ダッシュボード設計書

## 1. 概要

代替機管理システムのナビゲーションとダッシュボードの設計仕様を定義する。

## 2. ナビゲーション設計

### 2.1 ナビゲーションバー構成

#### 2.1.1 構造
```
navbar
├── navbar-brand（システム名・ロゴ）
│   └── 代替機管理システム
└── navbar-nav（メニュー項目）
    ├── ダッシュボード
    ├── スプレッドシート
    ├── URL生成
    ├── 機種マスタ
    ├── 拠点マスタ
    ├── レポート
    └── 設定
```

#### 2.1.2 メニュー項目詳細

| メニュー項目 | アイコン | 説明 | ページID |
|------------|---------|------|----------|
| ダッシュボード | fa-home | システムのメインページ | dashboard |
| スプレッドシート | fa-table | 機器データの閲覧・管理 | spreadsheet |
| URL生成 | fa-link | 拠点管理番号からURL生成 | url-generator |
| 機種マスタ | fa-database | 機種情報の管理 | model-master |
| 拠点マスタ | fa-map-marker-alt | 拠点情報の管理 | location-master |
| レポート | fa-chart-bar | 各種レポート表示 | reports |
| 設定 | fa-cog | システム設定管理 | settings |

### 2.2 ナビゲーション機能

#### 2.2.1 基本動作
- シングルページアプリケーション（SPA）形式
- `showPage(pageId)` 関数でページ切り替え
- アクティブなメニュー項目のハイライト表示
- レスポンシブ対応（モバイル表示時の折りたたみ）

#### 2.2.2 状態管理
```javascript
// ページ切り替え処理
function showPage(pageId) {
  // 全ページを非表示
  document.querySelectorAll('.page-content').forEach(page => {
    page.classList.remove('active');
  });
  
  // 選択ページを表示
  const targetPage = document.getElementById(pageId);
  if (targetPage) {
    targetPage.classList.add('active');
  }
  
  // ナビゲーションのアクティブ状態を更新
  updateNavigation(pageId);
}
```

## 3. ダッシュボード設計

### 3.1 ダッシュボード構成

#### 3.1.1 レイアウト
```
dashboard
├── dashboard-header
│   ├── dashboard-title（代替機管理システムダッシュボード）
│   └── dashboard-subtitle（関西フィールドサービス課）
└── dashboard-cards
    ├── スプレッドシートビューアー カード
    └── URL生成ツール カード
```

### 3.2 ダッシュボードカード

#### 3.2.1 カードコンポーネント構造
```
dashboard-card
├── card-icon（機能アイコン）
├── card-title（機能名）
├── card-description（機能説明）
└── card-button（機能へのリンクボタン）
```

#### 3.2.2 カード詳細

**スプレッドシートビューアー**
- アイコン: fa-table
- 説明: 各拠点のスプレッドシートデータを閲覧・編集できます。リアルタイムでステータスの更新も可能です。
- リンク先: spreadsheet

**URL生成ツール**
- アイコン: fa-link
- 説明: 拠点管理番号を共通フォームURLに追加して個別URLを生成します。生成したURLは自動的にマスタデータに保存されます。
- リンク先: url-generator

### 3.3 拡張性

将来的に以下のカードを追加可能：
- 統計情報カード（機器数、貸出状況など）
- 通知カード（期限切れアラートなど）
- クイックアクションカード（よく使う機能へのショートカット）

## 4. スタイリング

### 4.1 ナビゲーションバースタイル

```css
.navbar {
  background-color: #1a237e;
  color: white;
  padding: 0;
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  z-index: 1000;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.nav-link {
  color: rgba(255, 255, 255, 0.8);
  padding: 1rem 1.5rem;
  transition: all 0.3s ease;
}

.nav-link.active {
  background-color: rgba(255, 255, 255, 0.1);
  color: white;
}
```

### 4.2 ダッシュボードスタイル

```css
.dashboard-cards {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
  gap: 2rem;
  margin-top: 2rem;
}

.dashboard-card {
  background: white;
  border-radius: 12px;
  padding: 2rem;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
  transition: transform 0.3s ease, box-shadow 0.3s ease;
}

.dashboard-card:hover {
  transform: translateY(-4px);
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.15);
}
```

## 5. アクセシビリティ

### 5.1 キーボードナビゲーション
- Tabキーでメニュー項目を移動
- Enterキーで選択
- Escapeキーでモバイルメニューを閉じる

### 5.2 スクリーンリーダー対応
- 適切なaria-label属性の設定
- ランドマークロールの使用（nav, main）
- フォーカス管理

## 6. パフォーマンス最適化

### 6.1 遅延読み込み
- 非表示ページのコンテンツは初回アクセス時に読み込み
- 画像やアイコンの遅延読み込み

### 6.2 キャッシュ戦略
- ナビゲーション状態のセッションストレージ保存
- 頻繁にアクセスされるデータのキャッシュ

## 7. 今後の拡張計画

### 7.1 ナビゲーション拡張
- ユーザー権限に基づくメニュー表示制御
- お気に入り機能
- 最近使った機能の表示

### 7.2 ダッシュボード拡張
- カスタマイズ可能なダッシュボード
- ウィジェットの追加・削除・並び替え
- リアルタイムデータの表示（WebSocket使用）

## 8. 実装ファイル

### 8.1 関連ファイル一覧

| ファイル名 | 説明 | 役割 |
|-----------|------|------|
| navigation.html | ナビゲーションバーコンポーネント | UI構造 |
| dashboard.html | ダッシュボードページコンポーネント | UI構造 |
| dashboard-styles.html | ダッシュボード専用スタイル | スタイリング |
| main.html | メインJavaScript（ページ切り替えロジック） | 動作制御 |
| Index.html | エントリーポイント | 統合 |

### 8.2 削除されたコンポーネント
- データタイプマスタ関連のメニュー項目とページ（2025年1月に廃止）