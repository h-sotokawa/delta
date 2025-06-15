# Dev Container セットアップ

このプロジェクトでは、Node.js 22 Alpine ベースのDev Container環境を提供しています。

## 含まれる機能

- Node.js v22.x (Alpine Linux ベース)
- npm v10.x
- Google Apps Script 開発用の `clasp` CLI
- Git, Bash, その他開発ツール
- VS Code 拡張機能の自動インストール

## 使用方法

1. VS Codeで「Remote-Containers」拡張機能がインストールされていることを確認
2. プロジェクトを開く
3. コマンドパレット（Ctrl+Shift+P）から「Remote-Containers: Reopen in Container」を選択
4. コンテナの構築と起動を待つ

## インストールされる拡張機能

### JavaScript/TypeScript 開発
- TypeScript サポート
- Prettier (コードフォーマッター)
- ESLint (リンター)
- Tailwind CSS サポート

### Node.js 開発
- Node.js デバッガー
- npm IntelliSense
- npm スクリプト実行

### Google Apps Script 開発
- Google CLASP

### その他の便利ツール
- GitLens
- GitHub Copilot
- Material Icon Theme
- Docker サポート

## ポート転送

以下のポートが自動的に転送されます：
- 3000, 3001: 開発サーバー用
- 5000: 追加の開発サーバー用
- 8000, 8080: Web サーバー用

## 開発環境の確認

コンテナ内で以下のコマンドを実行して環境を確認できます：

```bash
# Node.js バージョン確認
node -v  # v22.16.0 が表示される

# npm バージョン確認
npm -v   # 10.9.2 が表示される

# clasp の確認
clasp --version
```

## トラブルシューティング

### 権限の問題
コンテナ内では `node` ユーザーとして実行されます。sudo権限も利用可能です。

### Git 設定
ホストマシンの `.gitconfig` と `.ssh` 設定が自動的にマウントされます。 