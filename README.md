# Google Apps Script プロジェクトテンプレート

このテンプレートは、Google Apps Scriptプロジェクトを効率的に開発するための基盤を提供します。

GitHub Actionsによる自動デプロイ機能を備えており、`main`ブランチへのプッシュで自動的にGoogle Apps Scriptへデプロイされます。

---

## 📁 プロジェクト構造

```
.
├── .github/
│   └── workflows/
│       ├── deploy-gas.yml      # GitHub Actionsワークフロー
│       └── README.md           # ワークフローの詳細説明
├── src/
│   ├── appsscript.json         # GAS設定ファイル
│   ├── core/
│   │   ├── Code.gs             # メインロジック
│   │   └── Config.gs           # プロジェクト設定
│   ├── integrations/
│   │   └── ExternalService.gs  # 外部サービス連携
│   └── ui/
│       ├── Sidebar.html        # サイドバーUI
│       └── dialogs/
│           ├── SettingsDialog.html  # 設定ダイアログ
│           └── HelpDialog.html      # ヘルプダイアログ
├── docs/
│   └── context.md              # プロジェクト要件定義書
├── .clasp.json                 # clasp設定（スクリプトID）
├── .claspignore                # claspデプロイ除外設定
└── README.md                   # このファイル
```

---

## 🚀 セットアップ手順

### 1. このテンプレートを使用

1. GitHubで「Use this template」ボタンをクリック
2. 新しいリポジトリ名を入力して作成
3. ローカルにクローン

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPO.git
cd YOUR_REPO
```

### 2. Google Apps Scriptプロジェクトを作成

#### オプションA: 新規プロジェクトを作成

```bash
# claspをインストール
npm install -g @google/clasp

# Googleアカウントでログイン
clasp login

# 新規GASプロジェクトを作成
clasp create --type standalone --title "My GAS Project" --rootDir ./src
```

#### オプションB: 既存プロジェクトに接続

1. Google Apps Scriptの既存プロジェクトを開く
2. プロジェクト設定からスクリプトIDをコピー
3. `.clasp.json`を編集してスクリプトIDを設定

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "./src"
}
```

### 3. プロジェクトをカスタマイズ

#### 基本設定

- `src/core/Config.gs` - プロジェクトの設定を記述
- `src/core/Code.gs` - メインロジックを実装
- `docs/context.md` - プロジェクトの要件定義を記述

#### UI（必要な場合）

- `src/ui/dialogs/SettingsDialog.html` - 設定画面
- `src/ui/dialogs/HelpDialog.html` - ヘルプ画面
- `src/ui/Sidebar.html` - サイドバー

#### 外部連携（必要な場合）

- `src/integrations/ExternalService.gs` - 外部API連携

### 4. ローカルでテスト

```bash
# GASにプッシュ
clasp push

# ブラウザで開く
clasp open
```

---

## 🤖 GitHub Actionsによる自動デプロイ

### セットアップ手順

#### 1. Apps Script APIを有効化

https://script.google.com/home/usersettings にアクセスして「Google Apps Script API」を有効化

#### 2. clasp認証情報を取得

ローカル環境で以下のコマンドを実行:

```bash
# claspでログイン（未実施の場合）
clasp login

# 認証情報をbase64エンコード
cat ~/.clasprc.json | base64 | tr -d '\n'
```

出力された文字列をコピー

#### 3. GitHub Secretsを設定

1. GitHubリポジトリの「Settings」→「Secrets and variables」→「Actions」を開く
2. 「New repository secret」をクリック
3. 以下のシークレットを追加:

| Name | Value |
|------|-------|
| `CLASPRC_JSON_BASE64` | 手順2でコピーしたbase64文字列 |

#### 4. `.clasp.json`の扱い

**オプションA: リポジトリに含める（推奨）**

`.clasp.json`をリポジトリにコミット（既に含まれています）

```bash
git add .clasp.json
git commit -m "Add clasp config"
git push
```

**オプションB: GitHub Secretsで管理**

1. `.clasp.json`をGitで管理しない場合は`.gitignore`に追加
2. GitHub Secretsに`CLASP_JSON`を追加:

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "./src"
}
```

### 自動デプロイの仕組み

- **トリガー**: `main`ブランチへのプッシュ
- **処理**: `src/`配下のファイルがGoogle Apps Scriptに自動デプロイ
- **確認**: GitHubの「Actions」タブでデプロイ状況を確認

#### 手動デプロイ

GitHub Actionsから手動で実行することも可能です:

1. GitHubリポジトリの「Actions」タブを開く
2. 「Deploy to Google Apps Script」ワークフローを選択
3. 「Run workflow」ボタンをクリック

---

## 📝 開発ワークフロー

### 基本的な流れ

```bash
# 1. 機能開発
# src/配下のファイルを編集

# 2. ローカルでテスト
clasp push
clasp open  # ブラウザでGASエディタを開く

# 3. コミット & プッシュ
git add .
git commit -m "Add new feature"
git push origin main

# 4. 自動デプロイ
# GitHub Actionsが自動的にGASへデプロイ
```

### ブランチ戦略（推奨）

```bash
# 開発用ブランチで作業
git checkout -b feature/new-feature

# 開発・テスト
clasp push
# ... 動作確認 ...

# mainブランチにマージ
git checkout main
git merge feature/new-feature
git push origin main  # ← 自動デプロイ実行
```

---

## 🔧 clasp コマンド一覧

### よく使うコマンド

```bash
# GASにプッシュ（ローカル→GAS）
clasp push

# GASから取得（GAS→ローカル）
clasp pull

# ブラウザでGASエディタを開く
clasp open

# デプロイの作成
clasp deploy

# ログの表示
clasp logs

# バージョン一覧
clasp versions
```

### プロジェクト管理

```bash
# 新規プロジェクト作成
clasp create --type standalone --title "プロジェクト名" --rootDir ./src

# 既存プロジェクトのクローン
clasp clone SCRIPT_ID --rootDir ./src

# プロジェクト情報の表示
clasp status
```

---

## 📚 ドキュメント

### プロジェクト固有のドキュメント

- `docs/context.md` - プロジェクトの要件定義書

### GitHub Actions関連

- `.github/workflows/README.md` - ワークフローの詳細説明
- `.github/workflows/deploy-gas.yml` - デプロイワークフロー

---

## 🔐 セキュリティのベストプラクティス

### 機密情報の管理

1. **GitHub Secrets を使用**
   - 認証情報は必ずGitHub Secretsで管理
   - `.clasprc.json`は絶対にコミットしない

2. **Apps Script のスクリプトプロパティを使用**
   - API キーやトークンはスクリプトプロパティで管理
   - コード内にハードコードしない

```javascript
// Good
const apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');

// Bad
const apiKey = 'sk-1234567890abcdef';  // ❌ ハードコード禁止
```

3. **.gitignore の確認**
   - `.clasprc.json`が除外されていることを確認
   - その他の機密情報も除外

---

## 🐛 トラブルシューティング

### clasp コマンドが使えない

```bash
# claspを再インストール
npm install -g @google/clasp

# ログイン状態を確認
clasp login --status
```

### GitHub Actionsでデプロイが失敗する

#### エラー: "User has not enabled the Apps Script API"

https://script.google.com/home/usersettings にアクセスして「Google Apps Script API」を有効化

#### エラー: "Could not find .clasp.json"

以下のいずれかを実施:
1. リポジトリに`.clasp.json`をコミット
2. GitHub Secretsに`CLASP_JSON`を設定

#### エラー: "base64: invalid input"

`CLASPRC_JSON_BASE64`に改行が含まれています。以下で再取得:

```bash
cat ~/.clasprc.json | base64 | tr -d '\n'
```

### clasp push でエラーが出る

```bash
# プロジェクト設定を確認
cat .clasp.json

# rootDirが正しいか確認
# 正しい例: "rootDir": "./src"

# 再度プッシュ
clasp push --force
```

---

## 🔗 参考リンク

### 公式ドキュメント

- [Google Apps Script 公式ドキュメント](https://developers.google.com/apps-script)
- [clasp (Google Apps Script CLI)](https://github.com/google/clasp)
- [GitHub Actions 公式ドキュメント](https://docs.github.com/en/actions)

### ガイド

- [Apps Script API 有効化](https://script.google.com/home/usersettings)
- [clasp を使った開発ワークフロー](https://github.com/google/clasp/blob/master/docs/README.md)

---

## 📝 ライセンス

このテンプレートは自由に使用・改変できます。

---

## 🙋 サポート

質問や問題がある場合は、GitHubのIssuesで報告してください。
