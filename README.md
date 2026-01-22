# ScanSnap to Notion

Google Driveのスキャンファイルをスプレッドシートで管理し、Notionデータベースに送信するGoogle Apps Scriptアプリケーション。

## 機能

- **Google Driveフォルダの読み込み**: 指定フォルダ内のファイル一覧をスプレッドシートに展開
- **ファイル名の同期**: スプレッドシートでファイル名を変更すると、Google Driveの実ファイル名も自動で変更
- **Notion連携**: 選択したファイルをNotionデータベースに送信（PDF埋め込み対応）
- **毎日自動送信**: 毎日9時に未送信ファイルを自動でNotionに送信
- **Discord通知**: Notionへの送信結果をDiscordに通知
- **サイドバーUI**: 各種操作をサイドバーから簡単に実行
- **初期設定ウィザード**: ステップバイステップで初期設定（設定済み項目はスキップ可能）
- **送信履歴管理**: 「送信履歴」シートで送信ログを管理

## スプレッドシートのカラム構成

| 選択 | ID | ファイル名 | ファイル形式 | 容量 | MIME Type | リンク先 | 作成日時 | 更新日時 | Notion送信済み |
|------|-----|-----------|-------------|------|-----------|---------|---------|---------|--------------|
| チェックボックス | ファイルID | 編集可能 | 拡張子 | サイズ | MIMEタイプ | リンク | 作成日 | 更新日 | ステータス |

## セットアップ

### 1. デプロイ

```bash
# claspでGASにデプロイ
clasp push
```

### 2. 初期設定

1. スプレッドシートを開く
2. 初期設定ウィザードが自動で表示される
3. 以下を設定:
   - **Google DriveフォルダID**: スキャンファイルが保存されているフォルダのID
   - **Notion Integration Key**: [Notion Integrations](https://www.notion.so/my-integrations)で作成
   - **Notion Parent ID**: データベースを作成するNotionページのID

### 3. ファイル名同期トリガーの有効化

スプレッドシートでファイル名を変更した際にGoogle Driveのファイル名も自動で変更するには、インストール可能なトリガーを有効化する必要があります。

**サイドバーから設定する方法:**
1. メニュー「ScanSnap to Notion」→「サイドバーを開く」
2. 「トリガー設定」セクションでトグルをONにする

**メニューから設定する方法:**
1. メニュー「ScanSnap to Notion」→「トリガー設定」→「ファイル名同期を有効化」

## 使い方

### サイドバーから操作（推奨）

1. メニュー「ScanSnap to Notion」→「サイドバーを開く」
2. サイドバーから各種操作を実行:
   - **ファイル一覧を更新**: 新規ファイルのみ追加（差分更新）
   - **ファイル一覧を再読込**: 全件再取得
   - **全選択/選択解除**: チェックボックスの一括操作
   - **選択ファイルをNotionに送信**: チェックしたファイルをNotionに送信
   - **トリガー設定**: ファイル名同期・毎日自動送信の有効/無効切り替え
   - **Discord通知設定**: Webhook URLの設定・テスト送信

### メニューから操作

メニュー「ScanSnap to Notion」から各機能を実行できます。

## プロジェクト構造

```
src/
├── core/
│   ├── Code.gs           # メイン処理・メニュー・トリガー
│   └── Config.gs         # 設定管理（PropertiesService）
├── integrations/
│   ├── DriveService.gs   # Google Drive連携
│   └── NotionService.gs  # Notion API連携
└── ui/
    ├── Sidebar.html      # サイドバーUI（幅300px）
    └── dialogs/
        └── SetupWizard.html  # 初期設定ウィザード
```

## 主要な関数

### メニュー・UI

| 関数名 | 説明 |
|-------|------|
| `onOpen()` | メニュー追加・初期設定チェック |
| `showSidebar()` | サイドバーを表示 |
| `showSetupWizard()` | 初期設定ウィザードを表示 |

### ファイル操作

| 関数名 | 説明 |
|-------|------|
| `loadDriveFilesToSheet()` | フォルダ内ファイルをシートに展開 |
| `refreshDriveFiles()` | ファイル一覧を差分更新 |
| `getSelectedRowsData()` | チェックされた行のデータを取得 |
| `selectAll()` | 全て選択 |
| `clearAllSelections()` | 選択解除 |

### Notion連携

| 関数名 | 説明 |
|-------|------|
| `createNotionDatabase()` | Notionにデータベースを作成 |
| `sendFilesToNotion()` | ファイルをNotionに送信 |
| `sendSelectedFilesToNotion()` | 選択ファイルをNotionに送信 |

### トリガー管理

| 関数名 | 説明 |
|-------|------|
| `isTriggerEnabled()` | ファイル名同期トリガーが有効か確認 |
| `enableEditTrigger()` | ファイル名同期トリガーを有効化 |
| `disableEditTrigger()` | ファイル名同期トリガーを無効化 |
| `onEditInstallable()` | 編集時のトリガーハンドラー |
| `isDailyTriggerEnabled()` | 毎日自動送信トリガーが有効か確認 |
| `enableDailyTrigger()` | 毎日自動送信トリガーを有効化 |
| `disableDailyTrigger()` | 毎日自動送信トリガーを無効化 |
| `dailyAutoSend()` | 毎日自動送信の実行（トリガーから呼び出し） |

### Discord通知

| 関数名 | 説明 |
|-------|------|
| `saveDiscordWebhookUrl()` | Discord Webhook URLを保存 |
| `isDiscordEnabled()` | Discord通知が有効か確認 |
| `sendDiscordNotification()` | Discordに通知を送信 |
| `testDiscordNotification()` | Discord通知のテスト送信 |

## 注意事項

### ファイル名同期について

- シンプルトリガー（onEdit）ではDriveAppの権限が不足するため、**インストール可能なトリガー**を使用
- サイドバーまたはメニューから「ファイル名同期を有効化」を実行する必要がある
- トリガーの状態はサイドバーで確認可能（有効/無効バッジ表示）

### Notion API制限

- Notion APIには レート制限があります
- 大量のファイルを一度に送信する場合は注意してください

### 必要な権限

- Google Drive: ファイルの読み取り・名前変更
- Google Sheets: スプレッドシートの読み書き
- External Services: Notion APIへのアクセス

## トラブルシューティング

### 「DriveApp.getFileById を呼び出すことができません」エラー

**原因**: シンプルトリガーでは権限が不足している

**解決方法**:
1. サイドバーを開く
2. 「トリガー設定」セクションでトグルをONにする
3. 権限の承認ダイアログが表示されたら承認する

### Notionにデータベースが作成されない

**確認事項**:
1. Integration KeyがNotionページに接続されているか
2. Parent IDが正しいか（ページURLの末尾のID）
3. IntegrationがページにInviteされているか

## ライセンス

MIT License
