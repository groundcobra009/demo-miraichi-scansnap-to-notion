/**
 * プロジェクト設定ファイル
 * PropertiesServiceを使用して設定を管理
 */

// ========================================
// 定数
// ========================================

const PROPERTY_KEYS = {
  DRIVE_FOLDER_ID: 'driveFolderId',
  NOTION_INTEGRATION_KEY: 'notionIntegrationKey',
  NOTION_PARENT_ID: 'notionParentId',
  NOTION_DATABASE_ID: 'notionDatabaseId',
  IS_SETUP_COMPLETE: 'isSetupComplete'
};

const SHEET_HEADERS = [
  '選択',
  'ID',
  'ファイル名',
  'ファイル形式',
  '容量',
  'MIME Type',
  'リンク先',
  '作成日時',
  '更新日時',
  'Notion送信済み'
];

// 送信履歴シートのヘッダー
const LOG_SHEET_NAME = '送信履歴';
const LOG_HEADERS = [
  '送信日時',
  'ファイルID',
  'ファイル名',
  'Notion Page ID',
  'Notion URL',
  'ステータス',
  'エラー詳細'
];

// ========================================
// 設定の取得・保存
// ========================================

/**
 * 設定を取得
 * @param {string} key - 設定キー
 * @returns {string|null} - 設定値
 */
function getConfig(key) {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(key);
}

/**
 * 設定を保存
 * @param {string} key - 設定キー
 * @param {string} value - 設定値
 */
function setConfig(key, value) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(key, value);
}

/**
 * 複数の設定を一括保存
 * @param {Object} configs - 設定オブジェクト
 */
function setConfigs(configs) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties(configs);
}

/**
 * すべての設定を取得
 * @returns {Object} - 全設定
 */
function getAllConfigs() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperties();
}

/**
 * 設定を削除
 * @param {string} key - 設定キー
 */
function deleteConfig(key) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(key);
}

/**
 * すべての設定をクリア
 */
function clearAllConfigs() {
  const props = PropertiesService.getScriptProperties();
  props.deleteAllProperties();
}

// ========================================
// セットアップ状態の確認
// ========================================

/**
 * 初期設定が完了しているか確認
 * @returns {boolean}
 */
function isSetupComplete() {
  const isComplete = getConfig(PROPERTY_KEYS.IS_SETUP_COMPLETE);
  return isComplete === 'true';
}

/**
 * 必要な設定がすべて揃っているか確認
 * @returns {Object} - 検証結果
 */
function validateSetup() {
  const driveFolderId = getConfig(PROPERTY_KEYS.DRIVE_FOLDER_ID);
  const notionKey = getConfig(PROPERTY_KEYS.NOTION_INTEGRATION_KEY);
  const notionParentId = getConfig(PROPERTY_KEYS.NOTION_PARENT_ID);

  const errors = [];

  if (!driveFolderId) {
    errors.push('Google DriveフォルダIDが設定されていません');
  }
  if (!notionKey) {
    errors.push('Notion Integration Keyが設定されていません');
  }
  if (!notionParentId) {
    errors.push('Notion Parent IDが設定されていません');
  }

  return {
    isValid: errors.length === 0,
    errors: errors,
    config: {
      driveFolderId: driveFolderId || '',
      notionIntegrationKey: notionKey ? '********' : '',
      notionParentId: notionParentId || '',
      notionDatabaseId: getConfig(PROPERTY_KEYS.NOTION_DATABASE_ID) || ''
    }
  };
}

/**
 * 現在の設定を取得（UI表示用）
 * @returns {Object}
 */
function getCurrentSettings() {
  return {
    driveFolderId: getConfig(PROPERTY_KEYS.DRIVE_FOLDER_ID) || '',
    notionIntegrationKey: getConfig(PROPERTY_KEYS.NOTION_INTEGRATION_KEY) ? '設定済み' : '未設定',
    notionParentId: getConfig(PROPERTY_KEYS.NOTION_PARENT_ID) || '',
    notionDatabaseId: getConfig(PROPERTY_KEYS.NOTION_DATABASE_ID) || '',
    isSetupComplete: isSetupComplete()
  };
}
