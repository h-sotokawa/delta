# ドキュメント構造

代替機管理システムのドキュメントを種類別に整理しています。

## ディレクトリ構造

```
.doc/
├── README.md                      # このファイル
├── CLAUDE.md                      # Claude Code用プロジェクト指針
├── 01_requirements/               # 要件定義
│   ├── 01_system_overview.md      # システム概要
│   ├── 02_functional_requirements.md # 機能要件
│   └── 03_non_functional_requirements.md # 非機能要件
├── 02_architecture/               # アーキテクチャ
│   └── 01_technical_architecture.md # 技術アーキテクチャ
├── 03_design/                     # 設計書
│   ├── 01_data_design.md          # データ設計
│   ├── 02_ui_ux_design.md         # UI/UX設計
│   ├── 03_api_design.md           # API設計
│   ├── 04_sheet_structures.md     # シート構造設計
│   ├── 05_management_id_design.md # 管理ID設計
│   ├── 06_index_sheet_design.md   # インデックスシート設計
│   └── 07_spreadsheet_viewer_design.md # スプレッドシートビューア設計
├── 04_implementation/             # 実装ガイド
│   ├── 01_test_strategy.md        # テスト戦略
│   ├── 02_implementation_guide.md # 実装ガイド
│   ├── 03_view_sheet_readme.md    # ビューシート実装
│   ├── 04_spreadsheet_functions_analysis.md # 関数分析
│   └── 05_status_driven_filtering_implementation.md # ステータス駆動フィルタリング
└── 05_refactoring/                # リファクタリング記録
    ├── 01_refactoring_summary.md  # リファクタリング概要
    ├── 02_spreadsheet_viewer_refactoring.md # ビューアリファクタリング
    ├── 03_url_generation_refactoring_summary.md # URL生成リファクタリング
    └── 04_documentation_updates_summary.md # ドキュメント更新記録
```

## カテゴリ説明

### 01_requirements/ - 要件定義
システムの基本要件、機能要件、非機能要件を定義

### 02_architecture/ - アーキテクチャ
システム全体の技術的なアーキテクチャ設計

### 03_design/ - 設計書
詳細な設計書類（データ設計、UI設計、API設計など）

### 04_implementation/ - 実装ガイド
実装時の具体的なガイド、テスト戦略、分析結果

### 05_refactoring/ - リファクタリング記録
過去のリファクタリング作業の記録と経緯

## 命名規則

- フォルダ: `01_カテゴリ名/`
- ファイル: `01_ファイル名.md`
- 数字は連番で、追加時は末尾に続ける

## 最終更新

2025年7月17日 - ドキュメント構造の整理とカテゴリ分類