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
    if (settings.driveFolderId) {
      setConfig(PROPERTY_KEYS.DRIVE_FOLDER_ID, settings.driveFolderId);
    }

    // Integration Keyは既存値を保持するフラグがない場合のみ更新
    if (!settings.keepExistingNotionKey && settings.notionIntegrationKey) {
      setConfig(PROPERTY_KEYS.NOTION_INTEGRATION_KEY, settings.notionIntegrationKey);
    }

    if (settings.notionParentId) {
      setConfig(PROPERTY_KEYS.NOTION_PARENT_ID, settings.notionParentId);
    }

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
    `Discord Webhook: ${settings.discordWebhookUrl}\n` +
    `セットアップ完了: ${settings.isSetupComplete ? 'はい' : 'いいえ'}\n` +
    `ファイル名同期トリガー: ${isTriggerEnabled() ? '有効' : '無効'}\n` +
    `毎日自動送信トリガー: ${isDailyTriggerEnabled() ? '有効' : '無効'}`,
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
const DAILY_TRIGGER_FUNCTION_NAME = 'dailyAutoSend';

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
// 毎日自動送信トリガー
// ========================================

/**
 * 毎日自動送信トリガーが有効かどうかを確認
 * @returns {boolean}
 */
function isDailyTriggerEnabled() {
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.some(trigger => trigger.getHandlerFunction() === DAILY_TRIGGER_FUNCTION_NAME);
}

/**
 * 毎日自動送信トリガーを有効化
 * @returns {Object} - 結果
 */
function enableDailyTrigger() {
  try {
    // 既に有効な場合は何もしない
    if (isDailyTriggerEnabled()) {
      return { success: true, message: 'トリガーは既に有効です' };
    }

    // 毎日午前9時に実行するトリガーを作成
    ScriptApp.newTrigger(DAILY_TRIGGER_FUNCTION_NAME)
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();

    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 毎日自動送信トリガーを無効化
 * @returns {Object} - 結果
 */
function disableDailyTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === DAILY_TRIGGER_FUNCTION_NAME) {
        ScriptApp.deleteTrigger(trigger);
      }
    }

    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 毎日自動送信の実行（トリガーから呼び出し）
 */
function dailyAutoSend() {
  try {
    // セットアップが完了していない場合は終了
    if (!isSetupComplete()) {
      console.log('dailyAutoSend: セットアップが完了していません');
      return;
    }

    // ファイル一覧を更新
    const refreshResult = refreshDriveFiles();
    if (!refreshResult.success) {
      console.error('dailyAutoSend: ファイル更新エラー', refreshResult.error);
      sendDiscordNotification('❌ 自動送信エラー', 'ファイル一覧の更新に失敗しました: ' + refreshResult.error);
      return;
    }

    // 未送信ファイルを取得
    const unsentFiles = getUnsentFiles();

    if (unsentFiles.length === 0) {
      console.log('dailyAutoSend: 送信するファイルがありません');
      return;
    }

    // Notionに送信
    const sendResult = sendFilesToNotion(unsentFiles);

    // Discord通知を送信
    const message = `📄 **送信完了**\n` +
      `成功: ${sendResult.success}件\n` +
      `失敗: ${sendResult.failed}件`;

    if (sendResult.errors.length > 0) {
      const errorDetails = sendResult.errors.map(e => `• ${e.fileName}: ${e.error}`).join('\n');
      sendDiscordNotification('📤 Notion自動送信結果', message + '\n\n**エラー詳細:**\n' + errorDetails);
    } else {
      sendDiscordNotification('📤 Notion自動送信結果', message);
    }

    console.log('dailyAutoSend: 完了', sendResult);
  } catch (error) {
    console.error('dailyAutoSend error:', error);
    sendDiscordNotification('❌ 自動送信エラー', 'エラーが発生しました: ' + error.message);
  }
}

/**
 * 未送信ファイルを取得
 * @returns {Array} - 未送信ファイルの配列
 */
function getUnsentFiles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ファイル一覧');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const sentFileIds = getSentFileIds();
  const dataRange = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.length);
  const data = dataRange.getValues();

  const unsentFiles = [];

  for (let i = 0; i < data.length; i++) {
    const rowData = data[i];
    const fileId = rowData[1]; // ID列

    if (!fileId || sentFileIds.has(fileId)) continue;

    unsentFiles.push({
      row: i + 2,
      fileId: fileId,
      fileName: rowData[2], // ファイル名列
      mimeType: rowData[5], // MIME Type列
      size: rowData[4], // 容量列
      url: rowData[6], // リンク先列
      createdTime: rowData[7] ? Utilities.formatDate(new Date(rowData[7]), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'+09:00'") : null,
      updatedTime: rowData[8] ? Utilities.formatDate(new Date(rowData[8]), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'+09:00'") : null,
      notionSent: false
    });
  }

  return unsentFiles;
}

// ========================================
// Discord通知
// ========================================

/**
 * Discord Webhook URLを保存
 * @param {string} url - Webhook URL
 * @returns {Object} - 保存結果
 */
function saveDiscordWebhookUrl(url) {
  try {
    setConfig(PROPERTY_KEYS.DISCORD_WEBHOOK_URL, url);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Discord Webhook URLを取得
 * @returns {string|null}
 */
function getDiscordWebhookUrl() {
  return getConfig(PROPERTY_KEYS.DISCORD_WEBHOOK_URL);
}

/**
 * Discord Webhook URLが設定されているか確認
 * @returns {boolean}
 */
function isDiscordEnabled() {
  const url = getDiscordWebhookUrl();
  return url && url.length > 0;
}

/**
 * Discordに通知を送信
 * @param {string} title - 通知タイトル
 * @param {string} message - 通知メッセージ
 * @returns {Object} - 送信結果
 */
function sendDiscordNotification(title, message) {
  const webhookUrl = getDiscordWebhookUrl();

  if (!webhookUrl) {
    console.log('Discord Webhook URLが設定されていません');
    return { success: false, error: 'Discord Webhook URLが設定されていません' };
  }

  try {
    const payload = {
      embeds: [{
        title: title,
        description: message,
        color: 5814783, // Notionの紫色
        timestamp: new Date().toISOString(),
        footer: {
          text: 'ScanSnap to Notion'
        }
      }]
    };

    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(webhookUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode >= 200 && responseCode < 300) {
      return { success: true };
    } else {
      return { success: false, error: 'Discord API Error: ' + responseCode };
    }
  } catch (error) {
    console.error('Discord notification error:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Discord通知テスト
 * @returns {Object} - テスト結果
 */
function testDiscordNotification() {
  return sendDiscordNotification('🔔 テスト通知', 'ScanSnap to Notionからのテスト通知です。');
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
