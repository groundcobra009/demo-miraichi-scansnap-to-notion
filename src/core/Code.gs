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
    .addItem('サイドバーを開く', 'showSidebar')
    .addSeparator()
    .addItem('初期設定', 'showSetupWizard')
    .addSeparator()
    .addItem('ファイル一覧を更新', 'menuRefreshFiles')
    .addItem('ファイル一覧を再読込', 'menuReloadFiles')
    .addSeparator()
    .addItem('選択ファイルをNotionに送信', 'menuSendToNotion')
    .addSeparator()
    .addSubMenu(ui.createMenu('トリガー設定')
      .addItem('ファイル名同期を有効化', 'menuEnableTrigger')
      .addItem('ファイル名同期を無効化', 'menuDisableTrigger'))
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
 * サイドバーを表示（幅最大300px）
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ui/Sidebar')
    .setTitle('ScanSnap to Notion')
    .setWidth(300);

  SpreadsheetApp.getUi().showSidebar(html);
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
 * 選択ファイルをNotionに送信（メニュー用）
 */
function menuSendToNotion() {
  if (!checkSetup()) return;

  const ui = SpreadsheetApp.getUi();
  const selectedFiles = getSelectedRowsData();

  if (selectedFiles.length === 0) {
    ui.alert('注意', '送信するファイルを選択してください。\n（チェックボックスで選択）', ui.ButtonSet.OK);
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
    // 送信済みステータスはsendFilesToNotion内でログシートに記録され、
    // syncNotionStatusFromLogで自動更新されます

    ui.alert(
      '送信完了',
      `Notionへの送信が完了しました。\n\n` +
      `成功: ${results.success}件\n` +
      `失敗: ${results.failed}件\n\n` +
      `送信履歴は「送信履歴」シートで確認できます。`,
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
 * 選択ファイルをNotionに送信（サイドバー用）
 * @returns {Object} - 送信結果
 */
function sendSelectedFilesToNotion() {
  const selectedFiles = getSelectedRowsData();

  if (selectedFiles.length === 0) {
    return { success: 0, failed: 0, errors: [] };
  }

  // 既に送信済みのファイルを除外
  const filesToSend = selectedFiles.filter(f => !f.notionSent);

  if (filesToSend.length === 0) {
    return { success: 0, failed: 0, errors: [] };
  }

  // 送信済みステータスはsendFilesToNotion内でログシートに記録され、
  // syncNotionStatusFromLogで自動更新されます
  return sendFilesToNotion(filesToSend);
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
    `セットアップ完了: ${settings.isSetupComplete ? 'はい' : 'いいえ'}\n` +
    `ファイル名同期トリガー: ${isTriggerEnabled() ? '有効' : '無効'}`,
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
    disableEditTrigger();
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
// トリガー管理
// ========================================

const TRIGGER_FUNCTION_NAME = 'onEditInstallable';

/**
 * トリガーが有効かどうかを確認
 * @returns {boolean}
 */
function isTriggerEnabled() {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.some(trigger => trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME);
}

/**
 * ファイル名同期トリガーを有効化
 * @returns {Object} - 結果
 */
function enableEditTrigger() {
  try {
    // 既に有効な場合は何もしない
    if (isTriggerEnabled()) {
      return { success: true, message: 'トリガーは既に有効です' };
    }

    // 新しいトリガーを作成
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * ファイル名同期トリガーを無効化
 * @returns {Object} - 結果
 */
function disableEditTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
        ScriptApp.deleteTrigger(trigger);
      }
    }

    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * トリガー有効化（メニュー用）
 */
function menuEnableTrigger() {
  const ui = SpreadsheetApp.getUi();
  const result = enableEditTrigger();

  if (result.success) {
    ui.alert('完了', 'ファイル名同期トリガーを有効化しました。\n\nスプレッドシートでファイル名を変更すると、\nGoogle Driveのファイル名も自動で変更されます。', ui.ButtonSet.OK);
  } else {
    ui.alert('エラー', 'トリガーの有効化に失敗しました: ' + result.error, ui.ButtonSet.OK);
  }
}

/**
 * トリガー無効化（メニュー用）
 */
function menuDisableTrigger() {
  const ui = SpreadsheetApp.getUi();
  const result = disableEditTrigger();

  if (result.success) {
    ui.alert('完了', 'ファイル名同期トリガーを無効化しました。', ui.ButtonSet.OK);
  } else {
    ui.alert('エラー', 'トリガーの無効化に失敗しました: ' + result.error, ui.ButtonSet.OK);
  }
}

// ========================================
// トリガーハンドラー（ファイル名変更の監視）
// ========================================

/**
 * インストール可能なonEditハンドラー
 * ファイル名が変更されたらGoogle Driveのファイル名も変更
 * @param {Object} e - イベントオブジェクト
 */
function onEditInstallable(e) {
  try {
    const sheet = e.source.getActiveSheet();

    // ファイル一覧シート以外は無視
    if (sheet.getName() !== 'ファイル一覧') return;

    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();

    // ヘッダー行は無視
    if (row === 1) return;

    // ファイル名列の変更を検知
    const fileNameCol = SHEET_HEADERS.indexOf('ファイル名') + 1;
    if (col !== fileNameCol) return;

    const newFileName = e.value;
    const oldFileName = e.oldValue;

    // 値が変更されていない場合は無視
    if (newFileName === oldFileName || !newFileName) return;

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
      if (oldFileName) {
        range.setValue(oldFileName);
      }
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `ファイル名の変更に失敗しました: ${result.error}`,
        'エラー',
        5
      );
    }
  } catch (error) {
    console.error('onEditInstallable error:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `エラーが発生しました: ${error.message}`,
      'エラー',
      5
    );
  }
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

  console.log('=== トリガー状態 ===');
  console.log('トリガー有効:', isTriggerEnabled());

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

  console.log('=== ファイル統計 ===');
  console.log(getFileListStats());
}
