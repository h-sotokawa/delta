# 代替機管理システム - システム概要

## 1. システム目的

代替機（端末・プリンタ・その他機器）の効率的な管理を目的とした Web アプリケーション。

### 1.1 背景・課題

- 複数拠点に分散する代替機の在庫管理の複雑化
- 手動での機器情報管理によるミス・効率性の低下
- 機器の貸出・返却時における情報共有の煩雑さ
- フォーム入力時の拠点管理番号手入力による負荷

### 1.2 解決アプローチ

- Google Sheets を活用した統合データ管理
- 拠点管理番号の自動生成とフォーム事前入力
- リアルタイムでの在庫状況可視化
- QR コードによる効率的なアクセス方法の提供

## 2. システム対象ユーザー

### 2.1 メインユーザー

- **関西フィールドサービス課スタッフ**: 日常的な機器管理業務
- **拠点管理者**: 各拠点での在庫確認・管理

### 2.2 管理ユーザー

- **システム管理者**: マスタデータ管理・設定変更
- **課長・責任者**: 全体状況の把握・レポート確認

## 3. 主要機能概要

### 3.1 ダッシュボード機能

- システム全体の概要表示
- 各機能への入口提供
- 重要な通知・アラートの表示

### 3.2 スプレッドシートビューアー

- 拠点別機器データの表示・検索(ステータス別に最適なデータを取得し表示させる)
- 「預り機のステータス」の表示・編集機能
- ステータス変更履歴の記録・管理
- リアルタイムデータ更新
- データのフィルタリング・ソート機能

### 3.3 URL 生成機能

- 拠点、カテゴリ、機種、製造番号より拠点管理番号の自動生成
- 登録する代替機が端末の場合、属性を入力できるようにする(OS、ソフト、資産管理番号)
- Google Forms 事前入力 URL 作成
- QR コード用 URL 生成(拠点管理番号を含めた Google Forms 事前入力 URL へ転送する web アプリ.site/gas-qr-redirect-final.gs を使用)
- マスタデータへの自動登録

### 3.4 マスタデータ管理

- 拠点マスタ管理
- 機種マスタ管理
- 機器マスタ管理（端末・プリンタ・その他）

### 3.5 システム設定

- 共通フォーム URL 設定
- 通知メールアドレス設定（エラー通知・アラート通知）
- debugMode による設定編集制御
- 管理者権限による設定変更

## 4. 技術アーキテクチャ概要

### 4.1 プラットフォーム

- **フロントエンド**: HTML5 + CSS3 + JavaScript
- **バックエンド**: Google Apps Script
- **データストレージ**: Google Sheets
- **デプロイ**: Google Apps Script Web App

### 4.2 選定理由

- **Google Workspace 環境**: 既存インフラとの親和性
- **メンテナンス性**: コード管理・更新の簡便性
- **コスト効率**: 追加ライセンス費用なし
- **セキュリティ**: Google 認証による統合セキュリティ

## 5. システム境界

### 5.1 対象範囲

- 代替機在庫管理
- 拠点管理番号生成・管理
- フォーム連携・事前入力
- 基本的なレポート機能

### 5.2 対象外

- 購買・発注機能
- 複雑な承認ワークフロー
- 外部システム連携
- 高度な分析・BI 機能

## 6. 成功指標

### 6.1 効率性指標

- フォーム入力時間の短縮（手入力 → 自動入力）
- 機器検索時間の短縮
- データ更新・同期の即時反映

### 6.2 品質指標

- データ入力ミスの削減
- 在庫状況の正確性向上
- 拠点間情報共有の改善

### 6.3 利用性指標

- ユーザー操作習得時間の短縮
- システム稼働率の維持
- モバイル対応による利便性向上
