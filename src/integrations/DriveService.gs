/**
 * Google Drive連携サービス
 */

// ========================================
// フォルダ操作
// ========================================

/**
 * 指定フォルダ内のファイル一覧を取得
 * @param {string} folderId - フォルダID（省略時は設定から取得）
 * @returns {Array} - ファイル情報の配列
 */
function getFilesFromFolder(folderId) {
  const targetFolderId = folderId || getConfig(PROPERTY_KEYS.DRIVE_FOLDER_ID);

  if (!targetFolderId) {
    throw new Error('Google DriveフォルダIDが設定されていません');
  }

  try {
    const folder = DriveApp.getFolderById(targetFolderId);
    const files = folder.getFiles();
    const fileList = [];

    while (files.hasNext()) {
      const file = files.next();
      fileList.push(getFileInfo(file));
    }

    // 作成日時の降順でソート
    fileList.sort((a, b) => new Date(b.createdTime) - new Date(a.createdTime));

    return fileList;
  } catch (error) {
    throw new Error('フォルダの読み込みに失敗しました: ' + error.message);
  }
}

/**
 * ファイルの詳細情報を取得
 * @param {File} file - Driveファイルオブジェクト
 * @returns {Object} - ファイル情報
 */
function getFileInfo(file) {
  return {
    fileId: file.getId(),
    fileName: file.getName(),
    mimeType: file.getMimeType(),
    fileType: getFileExtension(file.getName()),
    size: formatFileSize(file.getSize()),
    sizeBytes: file.getSize(),
    url: file.getUrl(),
    createdTime: formatDateTime(file.getDateCreated()),
    updatedTime: formatDateTime(file.getLastUpdated()),
    createdTimeISO: file.getDateCreated().toISOString(),
    updatedTimeISO: file.getLastUpdated().toISOString()
  };
}

/**
 * ファイル名から拡張子を取得
 * @param {string} fileName - ファイル名
 * @returns {string} - 拡張子
 */
function getFileExtension(fileName) {
  const parts = fileName.split('.');
  if (parts.length > 1) {
    return parts.pop().toUpperCase();
  }
  return '';
}

/**
 * ファイルサイズをフォーマット
 * @param {number} bytes - バイト数
 * @returns {string} - フォーマット済みサイズ
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 B';

  const units = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  const size = (bytes / Math.pow(1024, i)).toFixed(1);

  return size + ' ' + units[i];
}

/**
 * 日時をフォーマット
 * @param {Date} date - 日時
 * @returns {string} - フォーマット済み日時
 */
function formatDateTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

// ========================================
// ファイル操作
// ========================================

/**
 * ファイル名を変更
 * @param {string} fileId - ファイルID
 * @param {string} newName - 新しいファイル名
 * @returns {Object} - 変更結果
 */
function renameFile(fileId, newName) {
  try {
    const file = DriveApp.getFileById(fileId);
    file.setName(newName);

    return {
      success: true,
      fileId: fileId,
      newName: newName
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ファイルIDからファイル情報を取得
 * @param {string} fileId - ファイルID
 * @returns {Object} - ファイル情報
 */
function getFileById(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    return {
      success: true,
      file: getFileInfo(file)
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// ========================================
// スプレッドシート連携
// ========================================

/**
 * Google Driveフォルダの内容をスプレッドシートに展開
 * @returns {Object} - 処理結果
 */
function loadDriveFilesToSheet() {
  try {
    const files = getFilesFromFolder();
    const sheet = getOrCreateDataSheet();

    // ヘッダー行を設定
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

    if (files.length === 0) {
      return { success: true, fileCount: 0 };
    }

    // データを配列に変換
    const data = files.map(file => [
      file.fileId,
      file.fileName,
      file.fileType,
      file.size,
      file.mimeType,
      file.url,
      file.createdTime,
      file.updatedTime,
      '' // Notion送信済みフラグ
    ]);

    // データを書き込み
    sheet.getRange(2, 1, data.length, SHEET_HEADERS.length).setValues(data);

    // 列幅を調整
    sheet.autoResizeColumns(1, SHEET_HEADERS.length);

    // リンク列をハイパーリンクに設定
    const linkColumn = SHEET_HEADERS.indexOf('リンク先') + 1;
    for (let i = 0; i < files.length; i++) {
      const cell = sheet.getRange(i + 2, linkColumn);
      const url = files[i].url;
      if (url) {
        cell.setFormula('=HYPERLINK("' + url + '", "開く")');
      }
    }

    return { success: true, fileCount: files.length };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * データシートを取得または作成
 * @returns {Sheet} - シートオブジェクト
 */
function getOrCreateDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('ファイル一覧');

  if (!sheet) {
    sheet = ss.insertSheet('ファイル一覧');
  } else {
    // 既存データをクリア
    sheet.clear();
  }

  return sheet;
}

/**
 * スプレッドシートのデータを更新（差分更新）
 * @returns {Object} - 更新結果
 */
function refreshDriveFiles() {
  try {
    const files = getFilesFromFolder();
    const sheet = getOrCreateDataSheet();

    // 既存のファイルIDを取得
    const lastRow = sheet.getLastRow();
    const existingIds = new Set();

    if (lastRow > 1) {
      const idRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      idRange.forEach(row => existingIds.add(row[0]));
    }

    // 新規ファイルを追加
    let addedCount = 0;
    for (const file of files) {
      if (!existingIds.has(file.fileId)) {
        const newRow = [
          file.fileId,
          file.fileName,
          file.fileType,
          file.size,
          file.mimeType,
          file.url,
          file.createdTime,
          file.updatedTime,
          ''
        ];
        sheet.appendRow(newRow);
        addedCount++;
      }
    }

    return {
      success: true,
      totalFiles: files.length,
      addedFiles: addedCount
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * 選択行のファイル情報を取得
 * @returns {Array} - 選択されたファイル情報の配列
 */
function getSelectedRowsData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ファイル一覧');
  if (!sheet) {
    return [];
  }

  const selection = sheet.getActiveRange();
  if (!selection) {
    return [];
  }

  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  // ヘッダー行は除外
  if (startRow === 1 && numRows === 1) {
    return [];
  }

  const filesData = [];
  const actualStartRow = startRow === 1 ? 2 : startRow;
  const actualNumRows = startRow === 1 ? numRows - 1 : numRows;

  for (let i = 0; i < actualNumRows; i++) {
    const row = actualStartRow + i;
    const rowData = sheet.getRange(row, 1, 1, SHEET_HEADERS.length).getValues()[0];

    filesData.push({
      row: row,
      fileId: rowData[0],
      fileName: rowData[1],
      fileType: rowData[2],
      size: rowData[3],
      mimeType: rowData[4],
      url: rowData[5],
      createdTime: formatToISO(rowData[6]),
      updatedTime: formatToISO(rowData[7]),
      notionSent: rowData[8]
    });
  }

  return filesData;
}

/**
 * 日時文字列をISO形式に変換
 * @param {string} dateStr - 日時文字列
 * @returns {string} - ISO形式の日時
 */
function formatToISO(dateStr) {
  if (!dateStr) return null;
  try {
    const date = new Date(dateStr.replace(/\//g, '-'));
    return date.toISOString().split('T')[0];
  } catch (e) {
    return null;
  }
}

/**
 * 行のNotion送信済みステータスを更新
 * @param {number} row - 行番号
 * @param {string} status - ステータス（例: '送信済み'）
 */
function updateNotionStatus(row, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ファイル一覧');
  if (sheet) {
    const notionColumn = SHEET_HEADERS.indexOf('Notion送信済み') + 1;
    sheet.getRange(row, notionColumn).setValue(status);
  }
}
