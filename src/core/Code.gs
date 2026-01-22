/**
 * ScanSnap to Notion - メインコード
 *
 * Google Driveのファイルをスプレッドシートで管理し、
 * Notionに送信する機能を提供します。
 */

// ========================================
// メニュー・UI
// ========================================

/**
 * スプレッドシート起動時にカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('ScanSnap to Notion')
    .addItem('初期設定', 'showSetupWizard')
    .addSeparator()
    .addItem('ファイル一覧を更新', 'menuRefreshFiles')
    .addItem('ファイル一覧を再読込', 'menuReloadFiles')
    .addSeparator()
    .addItem('選択行をNotionに送信', 'menuSendToNotion')
    .addSeparator()
    .addItem('設定を確認', 'menuShowSettings')
    .addItem('設定をリセット', 'menuResetSettings')
    .addToUi();

  // 初期設定が完了していない場合は自動でウィザードを表示
  if (!isSetupComplete()) {
    showSetupWizard();
  }
}

/**
 * 初期設定ウィザードを表示
 */
function showSetupWizard() {
  const html = HtmlService.createHtmlOutputFromFile('ui/dialogs/SetupWizard')
    .setWidth(550)
    .setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, '初期設定ウィザード');
}

/**
 * 初期設定を保存（ウィザードから呼び出し）
 * @param {Object} settings - 設定オブジェクト
 * @returns {Object} - 保存結果
 */
function saveSetupSettings(settings) {
  try {
    setConfig(PROPERTY_KEYS.DRIVE_FOLDER_ID, settings.driveFolderId);
    setConfig(PROPERTY_KEYS.NOTION_INTEGRATION_KEY, settings.notionIntegrationKey);
    setConfig(PROPERTY_KEYS.NOTION_PARENT_ID, settings.notionParentId);

    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ========================================
// メニューアクション
// ========================================

/**
 * ファイル一覧を更新（差分更新）
 */
function menuRefreshFiles() {
  if (!checkSetup()) return;

  const ui = SpreadsheetApp.getUi();

  try {
    const result = refreshDriveFiles();

    if (result.success) {
      ui.alert(
        '更新完了',
        `ファイル一覧を更新しました。\n\n` +
        `総ファイル数: ${result.totalFiles}\n` +
        `新規追加: ${result.addedFiles}`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('エラー', '更新に失敗しました: ' + result.error, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('エラー', error.message, ui.ButtonSet.OK);
  }
}

/**
 * ファイル一覧を再読込（全件再取得）
 */
function menuReloadFiles() {
  if (!checkSetup()) return;

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    'ファイル一覧を再読込しますか？\n既存のデータは上書きされます。',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    const result = loadDriveFilesToSheet();

    if (result.success) {
      ui.alert(
        '再読込完了',
        `${result.fileCount}件のファイルを読み込みました。`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('エラー', '再読込に失敗しました: ' + result.error, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('エラー', error.message, ui.ButtonSet.OK);
  }
}

/**
 * 選択行をNotionに送信
 */
function menuSendToNotion() {
  if (!checkSetup()) return;

  const ui = SpreadsheetApp.getUi();
  const selectedFiles = getSelectedRowsData();

  if (selectedFiles.length === 0) {
    ui.alert('注意', '送信するファイルを選択してください。\n（行を選択してから実行）', ui.ButtonSet.OK);
    return;
  }

  // 既に送信済みのファイルを除外
  const filesToSend = selectedFiles.filter(f => !f.notionSent);

  if (filesToSend.length === 0) {
    ui.alert('注意', '選択されたファイルは全て送信済みです。', ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert(
    '確認',
    `${filesToSend.length}件のファイルをNotionに送信しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    const results = sendFilesToNotion(filesToSend);

    // 送信成功したファイルのステータスを更新
    for (const file of filesToSend) {
      updateNotionStatus(file.row, '送信済み');
    }

    ui.alert(
      '送信完了',
      `Notionへの送信が完了しました。\n\n` +
      `成功: ${results.success}件\n` +
      `失敗: ${results.failed}件`,
      ui.ButtonSet.OK
    );

    if (results.errors.length > 0) {
      console.log('送信エラー:', results.errors);
    }
  } catch (error) {
    ui.alert('エラー', '送信に失敗しました: ' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * 現在の設定を表示
 */
function menuShowSettings() {
  const ui = SpreadsheetApp.getUi();
  const settings = getCurrentSettings();

  ui.alert(
    '現在の設定',
    `Google DriveフォルダID: ${settings.driveFolderId || '未設定'}\n` +
    `Notion Integration Key: ${settings.notionIntegrationKey}\n` +
    `Notion Parent ID: ${settings.notionParentId || '未設定'}\n` +
    `Notion Database ID: ${settings.notionDatabaseId || '未作成'}\n` +
    `セットアップ完了: ${settings.isSetupComplete ? 'はい' : 'いいえ'}`,
    ui.ButtonSet.OK
  );
}

/**
 * 設定をリセット
 */
function menuResetSettings() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '警告',
    '全ての設定をリセットしますか？\nこの操作は元に戻せません。',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    clearAllConfigs();
    ui.alert('完了', '設定をリセットしました。', ui.ButtonSet.OK);
  }
}

/**
 * 初期設定が完了しているかチェック
 * @returns {boolean}
 */
function checkSetup() {
  if (!isSetupComplete()) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '初期設定が必要です',
      '初期設定ウィザードを開きますか？',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      showSetupWizard();
    }
    return false;
  }
  return true;
}

// ========================================
// トリガー（ファイル名変更の監視）
// ========================================

/**
 * セル編集時のトリガー
 * ファイル名が変更されたらGoogle Driveのファイル名も変更
 * @param {Object} e - イベントオブジェクト
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();

  // ファイル一覧シート以外は無視
  if (sheet.getName() !== 'ファイル一覧') return;

  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // ヘッダー行は無視
  if (row === 1) return;

  // ファイル名列（2列目）の変更を検知
  const fileNameCol = SHEET_HEADERS.indexOf('ファイル名') + 1;
  if (col !== fileNameCol) return;

  const newFileName = e.value;
  const oldFileName = e.oldValue;

  // 値が変更されていない場合は無視
  if (newFileName === oldFileName) return;

  // ファイルIDを取得
  const fileIdCol = SHEET_HEADERS.indexOf('ID') + 1;
  const fileId = sheet.getRange(row, fileIdCol).getValue();

  if (!fileId) return;

  // Google Driveのファイル名を変更
  const result = renameFile(fileId, newFileName);

  if (result.success) {
    // 更新日時を更新
    const updatedCol = SHEET_HEADERS.indexOf('更新日時') + 1;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    sheet.getRange(row, updatedCol).setValue(now);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `ファイル名を「${newFileName}」に変更しました`,
      '完了',
      3
    );
  } else {
    // エラー時は元に戻す
    range.setValue(oldFileName);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `ファイル名の変更に失敗しました: ${result.error}`,
      'エラー',
      5
    );
  }
}

/**
 * インストール可能なonEditトリガーを設定
 * （シンプルトリガーでは権限が足りない場合に使用）
 */
function createEditTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 新しいトリガーを作成
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

/**
 * インストール可能なonEditハンドラー
 * @param {Object} e - イベントオブジェクト
 */
function onEditInstallable(e) {
  onEdit(e);
}

// ========================================
// ユーティリティ関数
// ========================================

/**
 * テスト用：全機能の動作確認
 */
function testAllFunctions() {
  console.log('=== 設定確認 ===');
  console.log(validateSetup());

  console.log('=== Notion接続テスト ===');
  console.log(testNotionConnection());

  console.log('=== Google Driveファイル取得テスト ===');
  try {
    const files = getFilesFromFolder();
    console.log('取得ファイル数:', files.length);
    if (files.length > 0) {
      console.log('最初のファイル:', files[0]);
    }
  } catch (e) {
    console.log('エラー:', e.message);
  }
}
