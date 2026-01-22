/**
 * Notion API連携サービス
 */

const NOTION_API_VERSION = '2022-06-28';
const NOTION_API_BASE = 'https://api.notion.com/v1';

// ========================================
// API通信の基本関数
// ========================================

/**
 * Notion APIにリクエストを送信
 * @param {string} endpoint - APIエンドポイント
 * @param {string} method - HTTPメソッド
 * @param {Object} payload - リクエストボディ
 * @returns {Object} - レスポンス
 */
function notionRequest(endpoint, method, payload) {
  const notionKey = getConfig(PROPERTY_KEYS.NOTION_INTEGRATION_KEY);

  if (!notionKey) {
    throw new Error('Notion Integration Keyが設定されていません');
  }

  const options = {
    method: method,
    headers: {
      'Authorization': 'Bearer ' + notionKey,
      'Notion-Version': NOTION_API_VERSION,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload && (method === 'POST' || method === 'PATCH')) {
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(NOTION_API_BASE + endpoint, options);
  const responseCode = response.getResponseCode();
  const responseBody = JSON.parse(response.getContentText());

  if (responseCode >= 400) {
    throw new Error('Notion API Error: ' + (responseBody.message || JSON.stringify(responseBody)));
  }

  return responseBody;
}

// ========================================
// データベース操作
// ========================================

/**
 * ファイル管理用のNotionデータベースを作成
 * @returns {Object} - 作成結果
 */
function createNotionDatabase() {
  const parentId = getConfig(PROPERTY_KEYS.NOTION_PARENT_ID);

  if (!parentId) {
    return { success: false, error: 'Notion Parent IDが設定されていません' };
  }

  try {
    const payload = {
      parent: {
        type: 'page_id',
        page_id: parentId
      },
      title: [
        {
          type: 'text',
          text: {
            content: 'ScanSnap ファイル管理'
          }
        }
      ],
      properties: {
        'ファイル名': {
          title: {}
        },
        'ファイルID': {
          rich_text: {}
        },
        'ファイル形式': {
          select: {
            options: [
              { name: 'PDF', color: 'red' },
              { name: 'JPEG', color: 'blue' },
              { name: 'PNG', color: 'green' },
              { name: 'その他', color: 'gray' }
            ]
          }
        },
        '容量': {
          rich_text: {}
        },
        'Google Driveリンク': {
          url: {}
        },
        '作成日時': {
          date: {}
        },
        '更新日時': {
          date: {}
        },
        'ステータス': {
          select: {
            options: [
              { name: '未処理', color: 'gray' },
              { name: '処理中', color: 'yellow' },
              { name: '完了', color: 'green' }
            ]
          }
        }
      }
    };

    const result = notionRequest('/databases', 'POST', payload);

    // データベースIDを保存
    setConfig(PROPERTY_KEYS.NOTION_DATABASE_ID, result.id);
    setConfig(PROPERTY_KEYS.IS_SETUP_COMPLETE, 'true');

    return { success: true, databaseId: result.id };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ========================================
// ページ作成（PDF埋め込み対応）
// ========================================

/**
 * Google DriveファイルのプレビューURLを取得
 * @param {string} fileId - ファイルID
 * @returns {string} - プレビューURL
 */
function getGoogleDrivePreviewUrl(fileId) {
  return 'https://drive.google.com/file/d/' + fileId + '/preview';
}

/**
 * Google Driveファイルの埋め込みURLを取得
 * @param {string} fileId - ファイルID
 * @returns {string} - 埋め込みURL
 */
function getGoogleDriveEmbedUrl(fileId) {
  return 'https://drive.google.com/file/d/' + fileId + '/view';
}

/**
 * Notionデータベースにページを追加（PDF埋め込み対応）
 * @param {Object} fileData - ファイルデータ
 * @returns {Object} - 作成結果
 */
function addPageToNotion(fileData) {
  const databaseId = getConfig(PROPERTY_KEYS.NOTION_DATABASE_ID);

  if (!databaseId) {
    return { success: false, error: 'Notionデータベースが設定されていません' };
  }

  try {
    const fileType = getFileTypeCategory(fileData.mimeType);
    const embedUrl = getGoogleDriveEmbedUrl(fileData.fileId);

    // ページ本体のコンテンツブロックを作成（PDF埋め込み）
    const children = [];

    // ファイル情報のヘッダー
    children.push({
      object: 'block',
      type: 'heading_2',
      heading_2: {
        rich_text: [{
          type: 'text',
          text: { content: 'ファイルプレビュー' }
        }]
      }
    });

    // PDF/画像の埋め込み
    children.push({
      object: 'block',
      type: 'embed',
      embed: {
        url: embedUrl
      }
    });

    // ファイル情報セクション
    children.push({
      object: 'block',
      type: 'heading_3',
      heading_3: {
        rich_text: [{
          type: 'text',
          text: { content: 'ファイル情報' }
        }]
      }
    });

    // ファイル詳細をテーブルで表示
    children.push({
      object: 'block',
      type: 'bulleted_list_item',
      bulleted_list_item: {
        rich_text: [{
          type: 'text',
          text: { content: 'ファイルID: ' + fileData.fileId }
        }]
      }
    });

    children.push({
      object: 'block',
      type: 'bulleted_list_item',
      bulleted_list_item: {
        rich_text: [{
          type: 'text',
          text: { content: 'ファイル形式: ' + fileType }
        }]
      }
    });

    children.push({
      object: 'block',
      type: 'bulleted_list_item',
      bulleted_list_item: {
        rich_text: [{
          type: 'text',
          text: { content: '容量: ' + (fileData.size || '不明') }
        }]
      }
    });

    // Google Driveへのリンク
    children.push({
      object: 'block',
      type: 'divider',
      divider: {}
    });

    children.push({
      object: 'block',
      type: 'paragraph',
      paragraph: {
        rich_text: [{
          type: 'text',
          text: {
            content: 'Google Driveで開く',
            link: { url: fileData.url }
          },
          annotations: {
            bold: true,
            color: 'blue'
          }
        }]
      }
    });

    // ページを作成
    const payload = {
      parent: {
        database_id: databaseId
      },
      properties: {
        'ファイル名': {
          title: [
            {
              text: {
                content: fileData.fileName || ''
              }
            }
          ]
        },
        'ファイルID': {
          rich_text: [
            {
              text: {
                content: fileData.fileId || ''
              }
            }
          ]
        },
        'ファイル形式': {
          select: {
            name: fileType
          }
        },
        '容量': {
          rich_text: [
            {
              text: {
                content: fileData.size || ''
              }
            }
          ]
        },
        'Google Driveリンク': {
          url: fileData.url || null
        },
        '作成日時': {
          date: fileData.createdTime ? { start: fileData.createdTime } : null
        },
        '更新日時': {
          date: fileData.updatedTime ? { start: fileData.updatedTime } : null
        },
        'ステータス': {
          select: {
            name: '未処理'
          }
        }
      },
      children: children
    };

    const result = notionRequest('/pages', 'POST', payload);

    // Notion URLを生成
    const notionUrl = 'https://www.notion.so/' + result.id.replace(/-/g, '');

    return {
      success: true,
      pageId: result.id,
      notionUrl: notionUrl
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 複数のファイルをNotionに送信（ログ記録対応）
 * @param {Array} filesData - ファイルデータの配列
 * @returns {Object} - 送信結果
 */
function sendFilesToNotion(filesData) {
  const results = {
    success: 0,
    failed: 0,
    errors: []
  };

  for (const fileData of filesData) {
    const result = addPageToNotion(fileData);

    // ログシートに記録
    const logEntry = {
      fileId: fileData.fileId,
      fileName: fileData.fileName,
      notionPageId: result.pageId || '',
      notionUrl: result.notionUrl || '',
      status: result.success ? '成功' : '失敗',
      error: result.error || ''
    };
    addSendLog(logEntry);

    if (result.success) {
      results.success++;
    } else {
      results.failed++;
      results.errors.push({
        fileName: fileData.fileName,
        error: result.error
      });
    }
  }

  // ファイル一覧シートの送信済みステータスを更新
  syncNotionStatusFromLog();

  return results;
}

/**
 * MIMEタイプからファイル形式カテゴリを取得
 * @param {string} mimeType - MIMEタイプ
 * @returns {string} - カテゴリ名
 */
function getFileTypeCategory(mimeType) {
  if (!mimeType) return 'その他';

  if (mimeType.includes('pdf')) return 'PDF';
  if (mimeType.includes('jpeg') || mimeType.includes('jpg')) return 'JPEG';
  if (mimeType.includes('png')) return 'PNG';
  return 'その他';
}

/**
 * Notion APIの接続テスト
 * @returns {Object} - テスト結果
 */
function testNotionConnection() {
  try {
    const result = notionRequest('/users/me', 'GET');
    return {
      success: true,
      botName: result.name,
      message: 'Notion APIに接続しました'
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}
