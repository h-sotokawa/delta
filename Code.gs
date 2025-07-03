// 拠点名の日本語表示用マッピング（3拠点に統合）
const LOCATION_NAMES = {
  'osaka': '大阪',
  'kobe': '神戸', 
  'himeji': '姫路'
};

// 端末・プリンタ・機種マスタシート名
const MASTER_SHEET_NAMES = {
  terminal: '端末マスタ',
  printer: 'プリンタマスタ',
  other: 'その他マスタ',
  model: '機種マスタ'
};

// 旧拠点から新拠点へのマッピング（互換性維持用）
const LEGACY_LOCATION_MAPPING = {
  'osaka_desktop': 'osaka',
  'osaka_notebook': 'osaka', 
  'osaka_printer': 'osaka',
  'kobe': 'kobe',
  'himeji': 'himeji',
  'hyogo_printer': 'kobe'
};

// デバイスタイプに応じてマスタシートを決定する関数（既存の互換性維持）
function getTargetSheetName(deviceType) {
  if (deviceType === 'terminal' || deviceType === 'desktop' || deviceType === 'notebook') {
    return MASTER_SHEET_NAMES.terminal;
  } else if (deviceType === 'printer') {
    return MASTER_SHEET_NAMES.printer;
  } else if (deviceType === 'model') {
    return MASTER_SHEET_NAMES.model;
  }
  // デフォルトは端末マスタ
  return MASTER_SHEET_NAMES.terminal;
}

// 詳細カテゴリに応じてマスタシートを決定する関数（フォーム作成用）
function getTargetSheetNameByCategory(deviceCategory) {
  if (deviceCategory === 'SV' || deviceCategory === 'CL') {
    return MASTER_SHEET_NAMES.terminal;
  } else if (deviceCategory === 'プリンタ') {
    return MASTER_SHEET_NAMES.printer;
  } else if (deviceCategory === 'その他') {
    return MASTER_SHEET_NAMES.other;
  } else if (deviceCategory === 'model') {
    return MASTER_SHEET_NAMES.model;
  }
  // デフォルトは端末マスタ
  return MASTER_SHEET_NAMES.terminal;
}

// デバッグモード（本番環境ではfalseに設定）
const DEBUG = false;

// パフォーマンス監視用
let performanceMetrics = {
  totalRequests: 0,
  averageResponseTime: 0,
  lastResetTime: new Date()
};

// ログを保持する配列（本番環境では軽量化）
let serverLogs = [];

// 軽量化されたログ関数
function addLog(message, data = null) {
  // 本番環境では詳細ログを無効化してパフォーマンスを向上
  if (DEBUG) {
    const timestamp = new Date().toISOString();
    const log = {
      timestamp: timestamp,
      message: message,
      data: data
    };
    serverLogs.push(log);
    console.log(`[${timestamp}] ${message}`, data);
    return log;
  }
  return null;
}

// パフォーマンス測定開始
function startPerformanceTimer() {
  return Date.now();
}

// パフォーマンス測定終了
function endPerformanceTimer(startTime, operation) {
  const duration = Date.now() - startTime;
  performanceMetrics.totalRequests++;
  performanceMetrics.averageResponseTime = 
    (performanceMetrics.averageResponseTime * (performanceMetrics.totalRequests - 1) + duration) / performanceMetrics.totalRequests;
  
  addLog(`パフォーマンス: ${operation}`, `${duration}ms`);
  return duration;
}

// パフォーマンス統計を取得
function getPerformanceStats() {
  return {
    success: true,
    stats: performanceMetrics
  };
}

// パフォーマンス統計をリセット
function resetPerformanceStats() {
  performanceMetrics = {
    totalRequests: 0,
    averageResponseTime: 0,
    lastResetTime: new Date()
  };
  return { success: true, message: 'パフォーマンス統計をリセットしました' };
}

// ========================================
// 拠点マスタ管理機能
// ========================================

// 拠点マスタを格納するスプレッドシートのシート名
const LOCATION_MASTER_SHEET_NAME = '拠点マスタ';

// 拠点マスタ一覧を取得
function getLocationMaster() {
  const startTime = startPerformanceTimer();
  addLog('拠点マスタ一覧取得を開始');
  
  try {
    const sheet = getLocationMasterSheet();
    addLog('拠点マスタシート取得完了', { sheetName: sheet.getName() });
    
    const data = sheet.getDataRange().getValues();
    addLog('シートデータ取得完了', { rowCount: data.length, data: data });
    
    if (data.length <= 1) {
      // ヘッダー行のみまたはデータなし
      addLog('データなし: ヘッダー行のみ');
      return [];
    }
    
    const locations = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      addLog('行データ処理中', { rowIndex: i, row: row });
      
      if (row[0]) { // 拠点IDが存在する行のみ
        const today = "'" + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
        const location = {
          locationId: row[0],
          locationName: row[1],
          locationCode: row[2],
          status: row[3] || 'active',
          createdAt: row[4] || today,
          updatedAt: row[5] || today
        };
        locations.push(location);
        addLog('拠点データ追加', location);
      }
    }
    
    endPerformanceTimer(startTime, '拠点マスタ一覧取得');
    addLog('拠点マスタ一覧取得完了', { count: locations.length, locations: locations });
    
    return locations;
  } catch (error) {
    addLog('拠点マスタ一覧取得エラー', error.toString());
    throw new Error('拠点マスタ一覧の取得に失敗しました: ' + error.message);
  }
}

// 拠点マスタシートを取得（存在しない場合は作成）
function getLocationMasterSheet() {
  const spreadsheetId = getSpreadsheetIdFromProperty('destination');
  if (!spreadsheetId) {
    throw new Error('拠点マスタ用スプレッドシートIDが設定されていません');
  }
  
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let sheet = spreadsheet.getSheetByName(LOCATION_MASTER_SHEET_NAME);
  
  if (!sheet) {
    // シートが存在しない場合は作成
    sheet = spreadsheet.insertSheet(LOCATION_MASTER_SHEET_NAME);
    
    // ヘッダー行を作成
    const headers = ['拠点ID', '拠点名', '拠点コード', 'ステータス', '作成日', '更新日'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のスタイルを設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f0f0f0');
    
    // 初期データを追加（既存の拠点）
    const today = "'" + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    const initialData = [
      ['osaka', '大阪', 'OSA', 'active', today, today],
      ['kobe', '神戸', 'KOB', 'active', today, today],
      ['himeji', '姫路', 'HIM', 'active', today, today]
    ];
    
    if (initialData.length > 0) {
      sheet.getRange(2, 1, initialData.length, 6).setValues(initialData);
    }
    
    addLog('拠点マスタシートを作成しました');
  }
  
  return sheet;
}

// 拠点IDで拠点情報を取得
function getLocationById(locationId) {
  const startTime = startPerformanceTimer();
  addLog('拠点情報取得を開始', { locationId });
  
  try {
    const sheet = getLocationMasterSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === locationId) {
        const today = "'" + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
        const location = {
          locationId: row[0],
          locationName: row[1],
          locationCode: row[2],
          status: row[3] || 'active',
          createdAt: row[4] || today,
          updatedAt: row[5] || today
        };
        
        endPerformanceTimer(startTime, '拠点情報取得');
        addLog('拠点情報取得完了', location);
        
        return location;
      }
    }
    
    endPerformanceTimer(startTime, '拠点情報取得');
    addLog('拠点が見つかりませんでした', { locationId });
    
    return null;
  } catch (error) {
    addLog('拠点情報取得エラー', error.toString());
    throw new Error('拠点情報の取得に失敗しました: ' + error.message);
  }
}

// 新規拠点を追加
function addLocation(locationData) {
  const startTime = startPerformanceTimer();
  addLog('拠点追加を開始', locationData);
  
  try {
    // バリデーション
    if (!locationData.locationId || !locationData.locationName || !locationData.locationCode) {
      throw new Error('必須項目が不足しています');
    }
    
    // 拠点IDの重複チェック
    const existingLocation = getLocationById(locationData.locationId);
    if (existingLocation) {
      throw new Error('同じ拠点IDが既に存在します: ' + locationData.locationId);
    }
    
    const sheet = getLocationMasterSheet();
    const today = "'" + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
    
    // 新しい行を追加
    const newRow = [
      locationData.locationId,
      locationData.locationName,
      locationData.locationCode,
      locationData.status || 'active',
      today,
      today
    ];
    
    sheet.appendRow(newRow);
    
    endPerformanceTimer(startTime, '拠点追加');
    addLog('拠点追加完了', locationData);
    
    return { success: true, message: '拠点が正常に追加されました' };
  } catch (error) {
    addLog('拠点追加エラー', error.toString());
    throw new Error('拠点の追加に失敗しました: ' + error.message);
  }
}

// 拠点情報を更新
function updateLocation(locationData) {
  const startTime = startPerformanceTimer();
  addLog('拠点更新を開始', locationData);
  
  try {
    // バリデーション
    if (!locationData.locationId || !locationData.locationName || !locationData.locationCode) {
      throw new Error('必須項目が不足しています');
    }
    
    const sheet = getLocationMasterSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === locationData.locationId) {
        // 該当行を更新
        const range = sheet.getRange(i + 1, 1, 1, 6);
        const today = "'" + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
        const updatedRow = [
          locationData.locationId,
          locationData.locationName,
          locationData.locationCode,
          locationData.status || 'active',
          row[4], // 作成日は保持
          today // 更新日を現在日付に設定
        ];
        
        range.setValues([updatedRow]);
        
        endPerformanceTimer(startTime, '拠点更新');
        addLog('拠点更新完了', locationData);
        
        return { success: true, message: '拠点が正常に更新されました' };
      }
    }
    
    throw new Error('更新対象の拠点が見つかりません: ' + locationData.locationId);
  } catch (error) {
    addLog('拠点更新エラー', error.toString());
    throw new Error('拠点の更新に失敗しました: ' + error.message);
  }
}

// 拠点を削除
function deleteLocation(locationId) {
  const startTime = startPerformanceTimer();
  addLog('拠点削除を開始', { locationId });
  
  try {
    const sheet = getLocationMasterSheet();
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === locationId) {
        // 該当行を削除
        sheet.deleteRow(i + 1);
        
        endPerformanceTimer(startTime, '拠点削除');
        addLog('拠点削除完了', { locationId });
        
        return { success: true, message: '拠点が正常に削除されました' };
      }
    }
    
    throw new Error('削除対象の拠点が見つかりません: ' + locationId);
  } catch (error) {
    addLog('拠点削除エラー', error.toString());
    throw new Error('拠点の削除に失敗しました: ' + error.message);
  }
}

function doGet(e) {
  serverLogs = []; // ログをリセット
  addLog('doGet関数が呼び出されました', e);
  addLog('現在のユーザー', Session.getActiveUser().getEmail());
  addLog('実行環境', Session.getEffectiveUser().getEmail());
  
  try {
    addLog('テンプレート作成開始');
    const template = HtmlService.createTemplateFromFile('Index');
    addLog('テンプレート作成成功');
    
    addLog('HTML評価開始');
    const html = template.evaluate()
        .setTitle('スプレッドシートビューアー')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setFaviconUrl('https://www.google.com/favicon.ico');
    
    addLog('HTMLの評価成功', {
      title: html.getTitle(),
      content: html.getContent().substring(0, 100) + '...' // 最初の100文字のみ表示
    });
    
    return html;
  } catch (error) {
    addLog('doGetでエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack,
      name: error.name
    });
    
    return HtmlService.createHtmlOutput(`
      <html>
        <body>
          <h1>エラーが発生しました</h1>
          <pre>${error.toString()}</pre>
          <h2>デバッグ情報</h2>
          <pre>${JSON.stringify(serverLogs, null, 2)}</pre>
        </body>
      </html>
    `);
  }
}

function include(filename) {
  addLog('include関数が呼び出されました', filename);
  try {
    // ファイル名に拡張子がない場合は.htmlを追加
    const fileToInclude = filename.includes('.') ? filename : filename + '.html';
    const content = HtmlService.createHtmlOutputFromFile(fileToInclude).getContent();
    addLog('ファイルの読み込み成功', fileToInclude);
    return content;
  } catch (error) {
    addLog('includeでエラーが発生', {
      filename: filename,
      error: error.toString()
    });
    // エラーが発生した場合は、エラーメッセージを含むHTMLを返す
    return `<!-- Error loading ${filename}: ${error.toString()} -->`;
  }
}

// 旧拠点から新拠点への変換関数
function convertLegacyLocation(location) {
  return LEGACY_LOCATION_MAPPING[location] || location;
}

// 統一スプレッドシートIDをスクリプトプロパティから取得
function getSpreadsheetIdFromProperty(location) {
  // 旧拠点名を新拠点名に変換
  const normalizedLocation = convertLegacyLocation(location);
  
  // 全拠点で統一スプレッドシートを使用
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
}

// フロントエンド用の拠点一覧を取得
function getLocations() {
  try {
    return {
      success: true,
      locations: LOCATION_NAMES
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// 高速化された日付フォーマット関数
function formatDateFast(date) {
  if (!(date instanceof Date)) return date;
  
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const seconds = date.getSeconds();
  
  return `${year}/${month < 10 ? '0' + month : month}/${day < 10 ? '0' + day : day} ${hours < 10 ? '0' + hours : hours}:${minutes < 10 ? '0' + minutes : minutes}:${seconds < 10 ? '0' + seconds : seconds}`;
}

function getSpreadsheetData(location, queryType, deviceType = 'terminal') {
  const startTime = startPerformanceTimer();
  addLog('getSpreadsheetData関数が呼び出されました', { location, queryType, deviceType });
  
  try {
    // 旧拠点名を新拠点名に変換
    const normalizedLocation = convertLegacyLocation(location);
    
    // 統一スプレッドシートIDを取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません: ' + location);
    }

    addLog('使用するスプレッドシートID', spreadsheetId);
    
    // デバイスタイプに応じてマスタシートを決定
    const targetSheetName = getTargetSheetName(deviceType);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(targetSheetName);
    
    if (!sheet) {
      throw new Error('シート「' + targetSheetName + '」が見つかりません。');
    }

    // データを取得
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow === 0) {
      throw new Error('シートにデータがありません。');
    }
    
    const range = sheet.getRange(1, 1, lastRow, lastColumn);
    const data = range.getValues();
    
    // 日付データの処理
    const dateProcessingStart = Date.now();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        if (row[j] instanceof Date) {
          row[j] = formatDateFast(row[j]);
        }
      }
    }
    const dateProcessingTime = Date.now() - dateProcessingStart;
    addLog('日付処理時間', dateProcessingTime + 'ms');
    
    const responseTime = endPerformanceTimer(startTime, 'スプレッドシート取得');
    
    const response = {
      success: true,
      data: data,
      logs: DEBUG ? serverLogs : [],
      metadata: {
        location: LOCATION_NAMES[location],
        queryType: queryType,
        spreadsheetName: spreadsheet.getName(),
        sheetName: sheet.getName(),
        lastRow: lastRow,
        lastColumn: lastColumn,
        responseTime: responseTime,
        dateProcessingTime: dateProcessingTime,
        dataSize: data.length
      }
    };
    
    addLog('返却するレスポンス準備完了');
    return response;
    
  } catch (error) {
    endPerformanceTimer(startTime, 'エラー処理');
    addLog('エラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack,
      name: error.name
    });
    
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack,
        name: error.name
      },
      logs: DEBUG ? serverLogs : []
    };
  }
}

// 拠点一覧を取得する関数
function getLocations() {
  return {
    success: true,
    locations: LOCATION_NAMES
  };
}

// スプレッドシートIDが設定されているかチェックする関数
function checkSpreadsheetIdExists(location) {
  try {
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    return {
      success: true,
      exists: !!spreadsheetId,
      location: location,
      locationName: LOCATION_NAMES[location] || location,
      spreadsheetId: spreadsheetId
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      location: location,
      locationName: LOCATION_NAMES[location] || location
    };
  }
}

// 部分データ取得機能（ページネーション対応）
function getSpreadsheetDataPaginated(location, queryType, startRow = 1, maxRows = 100, deviceType = 'terminal') {
  addLog('getSpreadsheetDataPaginated関数が呼び出されました', { location, queryType, startRow, maxRows, deviceType });
  
  try {
    // 旧拠点名を新拠点名に変換
    const normalizedLocation = convertLegacyLocation(location);
    
    // 統一スプレッドシートIDを取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません: ' + location);
    }

    // デバイスタイプに応じてマスタシートを決定
    const targetSheetName = getTargetSheetName(deviceType);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(targetSheetName);
    
    if (!sheet) {
      throw new Error('シート「' + targetSheetName + '」が見つかりません。');
    }

    const totalRows = sheet.getLastRow();
    const totalColumns = sheet.getLastColumn();
    
    if (totalRows === 0) {
      throw new Error('シートにデータがありません。');
    }
    
    // ヘッダー行を取得
    const headerRange = sheet.getRange(1, 1, 1, totalColumns);
    const headers = headerRange.getValues()[0];
    
    // 指定された範囲のデータを取得
    const endRow = Math.min(startRow + maxRows - 1, totalRows);
    const actualStartRow = Math.max(startRow, 2); // ヘッダー行をスキップ
    
    let data = [headers]; // ヘッダーを最初に追加
    
    if (actualStartRow <= totalRows) {
      const dataRange = sheet.getRange(actualStartRow, 1, endRow - actualStartRow + 1, totalColumns);
      const rowData = dataRange.getValues();
      
      // 日付データの処理
      for (let i = 0; i < rowData.length; i++) {
        const row = rowData[i];
        for (let j = 0; j < row.length; j++) {
          if (row[j] instanceof Date) {
            row[j] = formatDateFast(row[j]);
          }
        }
      }
      
      data = data.concat(rowData);
    }
    
    const response = {
      success: true,
      data: data,
      pagination: {
        startRow: startRow,
        endRow: endRow,
        maxRows: maxRows,
        totalRows: totalRows,
        hasMore: endRow < totalRows
      },
      metadata: {
        location: LOCATION_NAMES[location],
        queryType: queryType,
        spreadsheetName: spreadsheet.getName(),
        sheetName: sheet.getName(),
        totalRows: totalRows,
        totalColumns: totalColumns
      }
    };
    
    return response;
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack,
        name: error.name
      }
    };
  }
}

// ステータス更新関数
function updateMachineStatus(rowIndex, newStatus, location, deviceType = 'terminal') {
  const startTime = startPerformanceTimer();
  
  try {
    // 旧拠点名を新拠点名に変換
    const normalizedLocation = convertLegacyLocation(location);
    
    // 統一スプレッドシートIDを取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    addLog('updateMachineStatus: ステータスを変更するスプレッドシートID', spreadsheetId);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません: ' + location);
    }
    
    // デバイスタイプに応じてマスタシートを決定
    const targetSheetName = getTargetSheetName(deviceType);
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(targetSheetName);
    
    // 行インデックスを1から始まる形式に変換（ヘッダー行を考慮）
    const actualRow = rowIndex + 1;
    
    // 更新前のA列の値を取得（変更履歴用）
    const oldStatusRange = sheet.getRange(actualRow, 1);
    const oldStatus = oldStatusRange.getValue();
    
    // マシン情報を取得（通知用）
    const machineNameRange = sheet.getRange(actualRow, 2);
    const machineTypeRange = sheet.getRange(actualRow, 3);
    const machineName = machineNameRange.getValue() || 'Unknown';
    const machineType = machineTypeRange.getValue() || 'Unknown';
    
    // A列：ステータスのみ更新
    const statusRange = sheet.getRange(actualRow, 1);
    statusRange.setValue(newStatus);
    
    const updateTime = endPerformanceTimer(startTime, 'ステータス更新');
    
    return {
      success: true,
      message: 'ステータスを更新しました',
      data: {
        rowIndex: rowIndex,
        oldStatus: oldStatus,
        newStatus: newStatus,
        spreadsheetId: spreadsheetId,
        updateTime: updateTime,
        machineInfo: {
          name: machineName,
          type: machineType
        }
      }
    };
  } catch (error) {
    endPerformanceTimer(startTime, 'ステータス更新エラー');
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack
      }
    };
  }
}

// 複数行のステータスを一括更新する関数
function updateMultipleStatuses(updates, location, deviceType = 'terminal') {
  const startTime = startPerformanceTimer();
  
  try {
    // 旧拠点名を新拠点名に変換
    const normalizedLocation = convertLegacyLocation(location);
    
    // 統一スプレッドシートIDを取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません: ' + location);
    }
    
    // デバイスタイプに応じてマスタシートを決定
    const targetSheetName = getTargetSheetName(deviceType);
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(targetSheetName);
    
    const results = [];
    
    // 一括更新のためのデータを準備
    for (const update of updates) {
      const { rowIndex, newStatus } = update;
      const actualRow = rowIndex + 1;
      
      // 更新前のA列の値を取得
      const oldStatusRange = sheet.getRange(actualRow, 1);
      const oldStatus = oldStatusRange.getValue();
      
      // マシン情報を取得
      const machineNameRange = sheet.getRange(actualRow, 2);
      const machineTypeRange = sheet.getRange(actualRow, 3);
      const machineName = machineNameRange.getValue() || 'Unknown';
      const machineType = machineTypeRange.getValue() || 'Unknown';
      
      // A列：ステータスのみ更新
      const statusRange = sheet.getRange(actualRow, 1);
      statusRange.setValue(newStatus);
      
      results.push({
        rowIndex: rowIndex,
        oldStatus: oldStatus,
        newStatus: newStatus,
        success: true,
        machineInfo: {
          name: machineName,
          type: machineType
        }
      });
    }
    
    const updateTime = endPerformanceTimer(startTime, '一括ステータス更新');
    
    return {
      success: true,
      message: `${updates.length}件のステータスを更新しました`,
      data: {
        updateCount: updates.length,
        results: results,
        updateTime: updateTime
      }
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '一括ステータス更新エラー');
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack
      }
    };
  }
}

// データ整合性チェック機能
function checkDataConsistency(location, deviceType = 'terminal') {
  try {
    // 旧拠点名を新拠点名に変換
    const normalizedLocation = convertLegacyLocation(location);
    
    // 統一スプレッドシートIDを取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません: ' + location);
    }
    
    // デバイスタイプに応じてマスタシートを決定
    const targetSheetName = getTargetSheetName(deviceType);
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(targetSheetName);
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow === 0) {
      return {
        success: true,
        message: 'シートにデータがありません',
        stats: { totalRows: 0, validRows: 0, invalidRows: 0 }
      };
    }
    
    const data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    const headers = data[0];
    
    let validRows = 0;
    let invalidRows = 0;
    const issues = [];
    
    // データ行をチェック（ヘッダー行をスキップ）
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let isValid = true;
      const rowIssues = [];
      
      // 基本的な検証
      if (!row[0] || row[0].toString().trim() === '') {
        rowIssues.push('ステータスが空です');
        isValid = false;
      }
      
      if (!row[1] || row[1].toString().trim() === '') {
        rowIssues.push('マシン名が空です');
        isValid = false;
      }
      
      // 日付の検証
      if (row[15] && !(row[15] instanceof Date) && isNaN(Date.parse(row[15]))) {
        rowIssues.push('更新日時の形式が不正です');
        isValid = false;
      }
      
      if (isValid) {
        validRows++;
      } else {
        invalidRows++;
        issues.push({
          row: i + 1,
          issues: rowIssues,
          data: row.slice(0, 3) // 最初の3列のみ表示
        });
      }
    }
    
    return {
      success: true,
      stats: {
        totalRows: lastRow - 1, // ヘッダー行を除く
        validRows: validRows,
        invalidRows: invalidRows,
        validityRate: ((validRows / (lastRow - 1)) * 100).toFixed(2) + '%'
      },
      issues: issues.slice(0, 10), // 最初の10件のみ表示
      metadata: {
        location: LOCATION_NAMES[location],
        spreadsheetName: ss.getName(),
        sheetName: sheet.getName(),
        checkTime: new Date()
      }
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack
      }
    };
  }
}

// システム全体の健全性チェック
function systemHealthCheck() {
  const startTime = Date.now();
  const results = {
    success: true,
    timestamp: new Date(),
    checks: {}
  };
  
  try {
    // 1. プロパティ設定チェック
    const locations = Object.keys(LOCATION_NAMES);
    const propertyCheck = {
      total: locations.length,
      configured: 0,
      missing: []
    };
    
    for (const location of locations) {
      const spreadsheetId = getSpreadsheetIdFromProperty(location);
      if (spreadsheetId) {
        propertyCheck.configured++;
      } else {
        propertyCheck.missing.push(location);
      }
    }
    
    results.checks.properties = propertyCheck;
    
    // 2. パフォーマンス統計
    results.checks.performance = performanceMetrics;
    
    // 3. 実行時間
    results.executionTime = Date.now() - startTime + 'ms';
    
    return results;
    
  } catch (error) {
    results.success = false;
    results.error = error.toString();
    results.executionTime = Date.now() - startTime + 'ms';
    return results;
  }
}

// 代替機フォーム回答シートのデータを取得する関数（新機能）
function getLocationSheetData(location, locationSheetName, queryType) {
  const startTime = startPerformanceTimer();
  addLog('getLocationSheetData関数が呼び出されました', { location, locationSheetName, queryType });
  
  try {
    // 旧拠点名を新拠点名に変換
    const normalizedLocation = convertLegacyLocation(location);
    
    // 統一スプレッドシートIDを取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません: ' + location);
    }

    addLog('使用するスプレッドシートID', spreadsheetId);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(locationSheetName);
    
    if (!sheet) {
      throw new Error('代替機フォーム回答シート「' + locationSheetName + '」が見つかりません。');
    }

    // データを取得
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow === 0) {
      throw new Error('代替機フォーム回答シートにデータがありません。');
    }
    
    const range = sheet.getRange(1, 1, lastRow, lastColumn);
    const data = range.getValues();
    
    // 日付データの処理
    const dateProcessingStart = Date.now();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        if (row[j] instanceof Date) {
          row[j] = formatDateFast(row[j]);
        }
      }
    }
    const dateProcessingTime = Date.now() - dateProcessingStart;
    addLog('日付処理時間', dateProcessingTime + 'ms');
    
    const responseTime = endPerformanceTimer(startTime, '代替機フォーム回答シート取得');
    
    const response = {
      success: true,
      data: data,
      logs: DEBUG ? serverLogs : [],
      metadata: {
        location: LOCATION_NAMES[location],
        locationSheetName: locationSheetName,
        queryType: queryType,
        spreadsheetName: spreadsheet.getName(),
        sheetName: sheet.getName(),
        lastRow: lastRow,
        lastColumn: lastColumn,
        responseTime: responseTime,
        dateProcessingTime: dateProcessingTime,
        dataSize: data.length
      }
    };
    
    addLog('返却するレスポンス準備完了');
    return response;
    
  } catch (error) {
    endPerformanceTimer(startTime, 'エラー処理');
    addLog('エラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack,
      name: error.name
    });
    
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack,
        name: error.name
      },
      logs: DEBUG ? serverLogs : []
    };
  }
}

// 代替機フォーム回答シートの特定セルを更新する関数（新機能）
function updateLocationSheetCell(location, locationSheetName, rowIndex, columnIndex, newValue) {
  const startTime = startPerformanceTimer();
  addLog('updateLocationSheetCell関数が呼び出されました', { location, locationSheetName, rowIndex, columnIndex, newValue });
  
  try {
    // 選択された拠点のスプレッドシートIDをスクリプトプロパティから取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにスプレッドシートIDが設定されていません: ' + location);
    }

    addLog('使用するスプレッドシートID', spreadsheetId);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(locationSheetName);
    
    if (!sheet) {
      throw new Error('代替機フォーム回答シート「' + locationSheetName + '」が見つかりません。');
    }

    // 行インデックスを1から始まる形式に変換（ヘッダー行を考慮）
    const actualRow = rowIndex + 1;
    const actualColumn = columnIndex + 1;
    
    // 更新前の値を取得（変更履歴用）
    const targetRange = sheet.getRange(actualRow, actualColumn);
    const oldValue = targetRange.getValue();
    
    // セルを更新
    targetRange.setValue(newValue);
    
    const updateTime = endPerformanceTimer(startTime, 'セル更新');
    
    addLog('セル更新成功', {
      cell: `${String.fromCharCode(64 + actualColumn)}${actualRow}`,
      oldValue: oldValue,
      newValue: newValue
    });
    
    return {
      success: true,
      message: 'セルを更新しました',
      data: {
        location: location,
        locationSheetName: locationSheetName,
        rowIndex: rowIndex,
        columnIndex: columnIndex,
        cell: `${String.fromCharCode(64 + actualColumn)}${actualRow}`,
        oldValue: oldValue,
        newValue: newValue,
        spreadsheetId: spreadsheetId,
        updateTime: updateTime
      }
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'セル更新エラー');
    addLog('セル更新でエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack,
      name: error.name
    });
    
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack,
        name: error.name
      }
    };
  }
}

// 監査・サマリー用スプレッドシートのデータを取得する関数
function getDestinationSheetData(sheetName, queryType) {
  const startTime = startPerformanceTimer();
  addLog('getDestinationSheetData関数が呼び出されました', { sheetName, queryType });
  
  // 関数の開始を明確にマーク
  addLog('=== getDestinationSheetData開始 ===', {
    sheetName: sheetName,
    queryType: queryType,
    timestamp: new Date().toISOString()
  });
  
  try {
    // SPREADSHEET_ID_DESTINATIONをスクリプトプロパティから取得
    addLog('スクリプトプロパティ取得開始');
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID_DESTINATION');
    
    addLog('スクリプトプロパティ取得結果', {
      spreadsheetId: spreadsheetId,
      spreadsheetIdType: typeof spreadsheetId,
      spreadsheetIdNull: spreadsheetId === null,
      spreadsheetIdUndefined: spreadsheetId === undefined
    });
    
    if (!spreadsheetId) {
      addLog('エラー: スプレッドシートIDが未設定');
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません');
    }

    addLog('使用するスプレッドシートID', spreadsheetId);
    
    // スプレッドシートを開く
    addLog('スプレッドシートを開いています...');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    addLog('スプレッドシート取得成功', {
      spreadsheetName: spreadsheet.getName(),
      spreadsheetId: spreadsheet.getId()
    });
    
    // 利用可能なシート一覧を取得
    const allSheets = spreadsheet.getSheets();
    const sheetNames = allSheets.map(s => s.getName());
    addLog('利用可能なシート一覧', {
      sheetNames: sheetNames,
      totalSheets: allSheets.length,
      targetSheet: sheetName
    });
    
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      addLog('エラー: 指定されたシートが見つかりません', {
        requestedSheet: sheetName,
        availableSheets: sheetNames
      });
      throw new Error('シート「' + sheetName + '」が見つかりません。利用可能なシート: ' + sheetNames.join(', '));
    }

    addLog('シート取得成功', {
      sheetName: sheet.getName(),
      sheetId: sheet.getSheetId()
    });

    // データを取得
    addLog('データ範囲情報を取得中...');
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    addLog('データ範囲情報', {
      lastRow: lastRow,
      lastColumn: lastColumn,
      isEmpty: lastRow === 0
    });
    
    if (lastRow === 0) {
      addLog('エラー: シートにデータがありません');
      throw new Error('シートにデータがありません。');
    }
    
    // 監査データの場合は3行目をヘッダーとして扱う
    const isAuditData = sheetName === '大阪' || sheetName === '神戸' || sheetName === '姫路';
    let dataStartRow = 1;
    let headerRowIndex = 0;
    
    if (isAuditData) {
      // 監査データの場合、ヘッダー行は3行目
      headerRowIndex = 2; // 0-based index
      dataStartRow = 1; // 1行目から取得して、フロントエンド側でヘッダー行を調整
    }
    
    addLog('データ読み込み開始', {
      range: `A${dataStartRow}:${String.fromCharCode(64 + lastColumn)}${lastRow}`,
      totalCells: (lastRow - dataStartRow + 1) * lastColumn,
      isAuditData: isAuditData,
      headerRowIndex: headerRowIndex
    });
    
    const range = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastColumn);
    const data = range.getValues();
    
    addLog('データ読み込み完了', {
      dataLength: data.length,
      firstRowLength: data[0] ? data[0].length : 0,
      sampleData: data.length > 0 ? data[0].slice(0, 3) : []
    });
    
    // データの構造をチェック
    let dateColumnCount = 0;
    let errorCellCount = 0;
    
    // 日付データの処理とエラーチェック
    const dateProcessingStart = Date.now();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        try {
          if (row[j] instanceof Date) {
            row[j] = formatDateFast(row[j]);
            dateColumnCount++;
          } else if (row[j] && typeof row[j] === 'object' && row[j].toString().startsWith('#')) {
            // エラー値を検出
            errorCellCount++;
            addLog('エラーセル検出', {
              row: i + 1,
              column: j + 1,
              cellValue: row[j].toString(),
              cellType: typeof row[j]
            });
          }
        } catch (cellError) {
          errorCellCount++;
          addLog('セル処理エラー', {
            row: i + 1,
            column: j + 1,
            cellValue: String(row[j]),
            error: cellError.toString()
          });
        }
      }
    }
    const dateProcessingTime = Date.now() - dateProcessingStart;
    
    addLog('データ処理完了', {
      dateProcessingTime: dateProcessingTime + 'ms',
      dateColumnCount: dateColumnCount,
      errorCellCount: errorCellCount
    });
    
    const responseTime = endPerformanceTimer(startTime, '監査・サマリーシート取得');
    
    const response = {
      success: true,
      data: data,
      logs: DEBUG ? serverLogs : [],
      metadata: {
        sheetType: getSheetTypeFromName(sheetName),
        sheetName: sheetName,
        queryType: queryType,
        spreadsheetName: spreadsheet.getName(),
        lastRow: lastRow,
        lastColumn: lastColumn,
        responseTime: responseTime,
        dateProcessingTime: dateProcessingTime,
        dataSize: data.length,
        dateColumnCount: dateColumnCount,
        errorCellCount: errorCellCount,
        availableSheets: sheetNames,
        isAuditData: isAuditData,
        headerRowIndex: headerRowIndex
      }
    };
    
    addLog('返却するレスポンス準備完了', {
      responseSuccess: response.success,
      responseDataLength: response.data.length,
      metadataKeys: Object.keys(response.metadata)
    });
    
    return response;
    
  } catch (error) {
    endPerformanceTimer(startTime, 'エラー処理');
    addLog('エラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack,
      name: error.name,
      sheetName: sheetName,
      queryType: queryType
    });
    
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack,
        name: error.name,
        sheetName: sheetName,
        queryType: queryType
      },
      logs: DEBUG ? serverLogs : []
    };
  }
}

// シート名からシートタイプを判定するヘルパー関数
function getSheetTypeFromName(sheetName) {
  if (sheetName === 'サマリー' || sheetName === 'summary') {
    return 'サマリーデータ';
  } else if (sheetName === '大阪' || sheetName === 'osaka') {
    return '大阪監査データ';
  } else if (sheetName === '神戸' || sheetName === 'kobe') {
    return '神戸監査データ';
  } else if (sheetName === '姫路' || sheetName === 'himeji') {
    return '姫路監査データ';
  } else {
    return '監査・サマリーデータ';
  }
}

// 監査・サマリー用のシート一覧を取得する関数
function getDestinationSheets() {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      return {
        success: false,
        error: 'SPREADSHEET_ID_DESTINATIONが設定されていません'
      };
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    
    const sheetList = {};
    sheets.forEach(sheet => {
      const name = sheet.getName();
      sheetList[name] = getSheetTypeFromName(name);
    });
    
    return {
      success: true,
      sheets: sheetList,
      spreadsheetName: spreadsheet.getName()
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// サマリーシートの詳細診断を行う関数
// 機種ID自動生成機能
function generateNextModelId() {
  const startTime = startPerformanceTimer();
  addLog('generateNextModelId関数が呼び出されました');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      // シートが存在しない場合は M00001 から開始
      addLog('機種マスタシートが存在しないため、M00001から開始');
      return 'M00001';
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      // データがない場合は M00001 から開始
      addLog('データが存在しないため、M00001から開始');
      return 'M00001';
    }
    
    // 既存の機種IDを取得して最大値を見つける
    const idRange = sheet.getRange(2, 1, lastRow - 1, 1); // A列の2行目以降
    const idValues = idRange.getValues().flat();
    
    let maxNumber = 0;
    const modelIdPattern = /^M(\d{5})$/;
    
    for (const id of idValues) {
      if (id && typeof id === 'string') {
        const match = id.match(modelIdPattern);
        if (match) {
          const number = parseInt(match[1], 10);
          if (number > maxNumber) {
            maxNumber = number;
          }
        }
      }
    }
    
    // 次の番号を生成
    const nextNumber = maxNumber + 1;
    const nextId = 'M' + nextNumber.toString().padStart(5, '0');
    
    addLog('機種ID生成完了', {
      maxNumber: maxNumber,
      nextNumber: nextNumber,
      nextId: nextId,
      existingIds: idValues.length
    });
    
    const responseTime = endPerformanceTimer(startTime, '機種ID生成');
    
    return nextId;
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種ID生成エラー');
    addLog('機種ID生成でエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    // エラーの場合もデフォルトIDを返す
    return 'M001';
  }
}

// 機種マスタ管理機能
function getModelMasterData() {
  const startTime = startPerformanceTimer();
  addLog('getModelMasterData関数が呼び出されました');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      addLog('機種マスタシートが存在しないため新規作成します');
      sheet = spreadsheet.insertSheet(MASTER_SHEET_NAMES.model);
      const headers = ['機種ID', '機種名', 'メーカー', 'カテゴリ', '作成日', '更新日', '備考'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // 列フォーマットを設定
      const dateColumn1 = sheet.getRange('E:E'); // 作成日列（機種ID追加により1列ずれる）
      const dateColumn2 = sheet.getRange('F:F'); // 更新日列
      dateColumn1.setNumberFormat('@'); // @は文字列フォーマット
      dateColumn2.setNumberFormat('@');
      
      addLog('新規シート作成時にフォーマット設定完了');
      
      return {
        success: true,
        data: [headers],
        message: '機種マスタシートを新規作成しました',
        metadata: {
          spreadsheetName: spreadsheet.getName(),
          sheetName: sheet.getName(),
          lastRow: 1,
          lastColumn: headers.length,
          responseTime: endPerformanceTimer(startTime, '機種マスタ作成')
        }
      };
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow === 0 || lastColumn === 0) {
      const headers = ['機種ID', '機種名', 'メーカー', 'カテゴリ', '作成日', '更新日', '備考'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // 列フォーマットを設定
      const dateColumn1 = sheet.getRange('E:E'); // 作成日列（機種ID追加により1列ずれる）
      const dateColumn2 = sheet.getRange('F:F'); // 更新日列
      dateColumn1.setNumberFormat('@'); // @は文字列フォーマット
      dateColumn2.setNumberFormat('@');
      
      addLog('ヘッダー追加時にフォーマット設定完了');
      
      return {
        success: true,
        data: [headers],
        message: '機種マスタにヘッダーを追加しました',
        metadata: {
          spreadsheetName: spreadsheet.getName(),
          sheetName: sheet.getName(),
          lastRow: 1,
          lastColumn: headers.length,
          responseTime: endPerformanceTimer(startTime, '機種マスタ初期化')
        }
      };
    }
    
    // 既存シートの日付列フォーマットをチェック・修正
    const dateColumn1 = sheet.getRange('E:E'); // 作成日列（機種ID追加により1列ずれる）
    const dateColumn2 = sheet.getRange('F:F'); // 更新日列
    const currentFormat1 = dateColumn1.getNumberFormat();
    const currentFormat2 = dateColumn2.getNumberFormat();
    
    if (currentFormat1 !== '@' || currentFormat2 !== '@') {
      addLog('既存シートの日付フォーマットを文字列に修正します', {
        currentFormat1: currentFormat1,
        currentFormat2: currentFormat2
      });
      dateColumn1.setNumberFormat('@'); // @は文字列フォーマット
      dateColumn2.setNumberFormat('@');
    }
    
    const range = sheet.getRange(1, 1, lastRow, lastColumn);
    const data = range.getValues();
    
    const responseTime = endPerformanceTimer(startTime, '機種マスタ取得');
    
    return {
      success: true,
      data: data,
      metadata: {
        spreadsheetName: spreadsheet.getName(),
        sheetName: sheet.getName(),
        lastRow: lastRow,
        lastColumn: lastColumn,
        responseTime: responseTime
      }
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタ取得エラー');
    addLog('機種マスタ取得でエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack
      }
    };
  }
}

function addModelMasterData(modelData) {
  const startTime = startPerformanceTimer();
  addLog('addModelMasterData関数が呼び出されました', modelData);
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(MASTER_SHEET_NAMES.model);
      const headers = ['機種ID', '機種名', 'メーカー', 'カテゴリ', '作成日', '更新日', '備考'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    const lastRow = sheet.getLastRow();
    
    // 機種名の重複チェック
    if (lastRow > 1) {
      const existingData = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
      for (let i = 0; i < existingData.length; i++) {
        const existingModelName = existingData[i][1]; // 機種名は2列目（インデックス1）
        const existingManufacturer = existingData[i][2]; // メーカーは3列目（インデックス2）
        
        if (existingModelName === modelData.modelName && existingManufacturer === modelData.manufacturer) {
          addLog('機種名重複エラー', {
            existingModelName: existingModelName,
            existingManufacturer: existingManufacturer,
            newModelName: modelData.modelName,
            newManufacturer: modelData.manufacturer
          });
          
          return {
            success: false,
            error: `機種名「${modelData.modelName}」（メーカー：${modelData.manufacturer}）は既に登録されています。`,
            errorType: 'DUPLICATE_MODEL'
          };
        }
      }
    }
    
    const nextRow = lastRow + 1;
    const currentDate = new Date();
    const dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    // 機種IDを自動生成
    const modelId = generateNextModelId();
    
    // 文字列として確実に保存するため、先頭にアポストロフィを付ける（Sheetsでは表示されない）
    const formattedDateString = "'" + dateString;
    
    const newRowData = [
      modelId, // 機種ID（自動生成）
      modelData.modelName || '',
      modelData.manufacturer || '',
      modelData.category || '',
      formattedDateString,
      formattedDateString,
      modelData.remarks || ''
    ];
    
    // データを設定
    const range = sheet.getRange(nextRow, 1, 1, newRowData.length);
    range.setValues([newRowData]);
    
    // 日付列を文字列として明示的に設定
    const dateRange1 = sheet.getRange(nextRow, 5); // 作成日列（機種ID追加により1列ずれる）
    const dateRange2 = sheet.getRange(nextRow, 6); // 更新日列
    dateRange1.setNumberFormat('@'); // @は文字列フォーマット
    dateRange2.setNumberFormat('@');
    
    addLog('新規データ追加完了', {
      rowNumber: nextRow,
      modelId: modelId,
      dateString: dateString,
      rawData: newRowData
    });
    
    const responseTime = endPerformanceTimer(startTime, '機種マスタ追加');
    
    return {
      success: true,
      message: '機種マスタに新しいデータを追加しました',
      data: {
        rowIndex: nextRow - 1,
        modelId: modelId,
        modelData: newRowData,
        updateTime: responseTime
      }
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタ追加エラー');
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack
      }
    };
  }
}

function updateModelMasterData(rowIndex, modelData) {
  const startTime = startPerformanceTimer();
  addLog('updateModelMasterData関数が呼び出されました', { rowIndex, modelData });
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      throw new Error('機種マスタシートが見つかりません');
    }

    const actualRow = rowIndex + 1;
    const currentDate = new Date();
    const dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    // 文字列として確実に保存するため、先頭にアポストロフィを付ける
    const formattedDateString = "'" + dateString;
    
    const oldData = sheet.getRange(actualRow, 1, 1, 7).getValues()[0]; // 機種ID列追加により7列に変更
    
    // 機種名の重複チェック（自分自身以外との重複をチェック）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existingData = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
      for (let i = 0; i < existingData.length; i++) {
        const existingRow = i + 2; // スプレッドシートの実際の行番号
        const existingModelName = existingData[i][1]; // 機種名は2列目（インデックス1）
        const existingManufacturer = existingData[i][2]; // メーカーは3列目（インデックス2）
        
        // 自分自身以外で同じ機種名・メーカーの組み合わせがある場合はエラー
        if (existingRow !== actualRow && 
            existingModelName === modelData.modelName && 
            existingManufacturer === modelData.manufacturer) {
          addLog('機種名重複エラー（更新時）', {
            existingRow: existingRow,
            existingModelName: existingModelName,
            existingManufacturer: existingManufacturer,
            updatingRow: actualRow,
            newModelName: modelData.modelName,
            newManufacturer: modelData.manufacturer
          });
          
          return {
            success: false,
            error: `機種名「${modelData.modelName}」（メーカー：${modelData.manufacturer}）は既に登録されています。`,
            errorType: 'DUPLICATE_MODEL'
          };
        }
      }
    }
    
    const updatedRowData = [
      oldData[0], // 機種IDは変更しない
      modelData.modelName || oldData[1],
      modelData.manufacturer || oldData[2],
      modelData.category || oldData[3],
      oldData[4], // 作成日は変更しない
      formattedDateString, // 更新日のみ変更
      modelData.remarks || oldData[6]
    ];
    
    // データを設定
    const range = sheet.getRange(actualRow, 1, 1, updatedRowData.length);
    range.setValues([updatedRowData]);
    
    // 更新日列を文字列として明示的に設定
    const dateRange = sheet.getRange(actualRow, 6); // 更新日列（機種ID追加により1列ずれる）
    dateRange.setNumberFormat('@'); // @は文字列フォーマット
    
    addLog('データ更新完了', {
      rowNumber: actualRow,
      dateString: dateString,
      oldData: oldData,
      newData: updatedRowData
    });
    
    const responseTime = endPerformanceTimer(startTime, '機種マスタ更新');
    
    return {
      success: true,
      message: '機種マスタを更新しました',
      data: {
        rowIndex: rowIndex,
        oldData: oldData,
        newData: updatedRowData,
        updateTime: responseTime
      }
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタ更新エラー');
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack
      }
    };
  }
}

function deleteModelMasterData(modelName, manufacturer) {
  const startTime = startPerformanceTimer();
  addLog('deleteModelMasterData関数が呼び出されました', { modelName, manufacturer });
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      throw new Error('機種マスタシートが見つかりません');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      throw new Error('削除対象のデータがありません');
    }
    
    // 全データを取得してターゲット行を検索
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues(); // ヘッダー行を除く（機種ID列追加により7列に変更）
    let targetRowNumber = -1;
    let deletedData = null;
    
    addLog('削除対象を検索中', {
      searchModelName: modelName,
      searchManufacturer: manufacturer,
      totalRows: data.length
    });
    
    // 機種名とメーカー名で一致する行を検索（機種IDが追加されたため列番号を調整）
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowModelName = row[1] ? row[1].toString().trim() : ''; // 機種名は1列目（0-based）
      const rowManufacturer = row[2] ? row[2].toString().trim() : ''; // メーカーは2列目
      
      addLog(`行${i + 2}をチェック`, {
        rowModelId: row[0],
        rowModelName,
        rowManufacturer,
        searchModelName: modelName.trim(),
        searchManufacturer: manufacturer.trim()
      });
      
      if (rowModelName === modelName.trim() && rowManufacturer === manufacturer.trim()) {
        targetRowNumber = i + 2; // スプレッドシートの実際の行番号（ヘッダー行を考慮）
        deletedData = row;
        addLog('削除対象を発見', {
          targetRowNumber,
          deletedData
        });
        break;
      }
    }
    
    if (targetRowNumber === -1) {
      throw new Error(`指定された機種が見つかりませんでした。機種名: ${modelName}, メーカー: ${manufacturer}`);
    }
    
    // 行を削除
    sheet.deleteRow(targetRowNumber);
    
    const responseTime = endPerformanceTimer(startTime, '機種マスタ削除');
    
    addLog('削除完了', {
      deletedRowNumber: targetRowNumber,
      deletedData: deletedData
    });
    
    return {
      success: true,
      message: '機種マスタからデータを削除しました',
      data: {
        deletedRowNumber: targetRowNumber,
        deletedData: deletedData,
        modelName: modelName,
        manufacturer: manufacturer,
        updateTime: responseTime
      }
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタ削除エラー');
    addLog('機種マスタ削除でエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack,
      modelName: modelName,
      manufacturer: manufacturer
    });
    
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack,
        modelName: modelName,
        manufacturer: manufacturer
      }
    };
  }
}

// フォーム作成時に機種マスタのデータを取得する関数
function getModelMasterForForm() {
  try {
    const result = getModelMasterData();
    
    if (!result.success) {
      return {
        success: false,
        error: result.error,
        models: []
      };
    }
    
    const models = [];
    if (result.data && result.data.length > 1) {
      for (let i = 1; i < result.data.length; i++) {
        const row = result.data[i];
        models.push({
          modelId: row[0] || '',     // 機種ID
          modelName: row[1] || '',   // 機種名
          manufacturer: row[2] || '', // メーカー
          category: row[3] || '',     // カテゴリ
          displayName: `${row[2] || ''} ${row[1] || ''}`.trim() // メーカー 機種名
        });
      }
    }
    
    return {
      success: true,
      models: models,
      totalCount: models.length
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      models: []
    };
  }
}

// カテゴリ別に機種マスタを取得する関数
function getModelMasterByCategory(category) {
  try {
    const result = getModelMasterForForm();
    
    if (!result.success) {
      return result;
    }
    
    const filteredModels = result.models.filter(model => 
      model.category === category
    );
    
    return {
      success: true,
      models: filteredModels,
      totalCount: filteredModels.length,
      category: category
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      models: []
    };
  }
}

// 機種マスタから利用可能なカテゴリ一覧を取得する関数
function getAvailableCategories() {
  try {
    const result = getModelMasterForForm();
    
    if (!result.success) {
      return {
        success: false,
        error: result.error,
        categories: []
      };
    }
    
    const categories = [...new Set(result.models.map(model => model.category))].filter(cat => cat);
    
    return {
      success: true,
      categories: categories,
      totalCount: categories.length
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      categories: []
    };
  }
}

// 機種マスタの診断機能
function diagnoseModelMaster() {
  const startTime = Date.now();
  const diagnosticLogs = [];
  
  function addDiagnosticLog(message, data = {}) {
    const logEntry = {
      timestamp: new Date().toISOString(),
      message: message,
      data: data
    };
    diagnosticLogs.push(logEntry);
    console.log(`[MODEL-MASTER-DIAGNOSTIC] ${message}`, data);
  }
  
  try {
    addDiagnosticLog('機種マスタ診断開始');
    
    // スプレッドシートID確認
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID_DESTINATION');
    
    addDiagnosticLog('スプレッドシートID確認', {
      spreadsheetId: spreadsheetId ? '設定済み' : '未設定',
      idLength: spreadsheetId ? spreadsheetId.length : 0
    });
    
    if (!spreadsheetId) {
      return {
        success: false,
        error: 'SPREADSHEET_ID_DESTINATIONが設定されていません',
        diagnostics: diagnosticLogs,
        executionTime: Date.now() - startTime + 'ms'
      };
    }
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    addDiagnosticLog('スプレッドシート取得成功', {
      name: spreadsheet.getName(),
      url: spreadsheet.getUrl()
    });
    
    // 全シート取得
    const allSheets = spreadsheet.getSheets();
    const sheetInfo = allSheets.map(sheet => ({
      name: sheet.getName(),
      id: sheet.getSheetId(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn(),
      hidden: sheet.isSheetHidden()
    }));
    
    addDiagnosticLog('全シート情報', {
      totalSheets: allSheets.length,
      sheets: sheetInfo,
      targetSheetName: MASTER_SHEET_NAMES.model
    });
    
    // 機種マスタシート確認
    let modelSheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    let sheetCreated = false;
    
    if (!modelSheet) {
      addDiagnosticLog('機種マスタシートが見つかりません。新規作成します');
      modelSheet = spreadsheet.insertSheet(MASTER_SHEET_NAMES.model);
      const headers = ['機種名', 'メーカー', 'カテゴリ', '作成日', '更新日', '備考'];
      modelSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheetCreated = true;
    }
    
    addDiagnosticLog('機種マスタシート情報', {
      name: modelSheet.getName(),
      id: modelSheet.getSheetId(),
      lastRow: modelSheet.getLastRow(),
      lastColumn: modelSheet.getLastColumn(),
      isHidden: modelSheet.isSheetHidden(),
      wasCreated: sheetCreated
    });
    
    // getModelMasterData関数の実際の呼び出しテスト
    addDiagnosticLog('getModelMasterData関数テスト開始');
    const result = getModelMasterData();
    
    addDiagnosticLog('getModelMasterData関数テスト結果', {
      success: result.success,
      hasData: result.data ? true : false,
      dataLength: result.data ? result.data.length : 0,
      error: result.error || 'なし',
      message: result.message || 'なし'
    });
    
    return {
      success: true,
      message: '機種マスタ診断完了',
      diagnostics: diagnosticLogs,
      sheetInfo: {
        name: modelSheet.getName(),
        lastRow: modelSheet.getLastRow(),
        lastColumn: modelSheet.getLastColumn(),
        wasCreated: sheetCreated
      },
      functionResult: result,
      executionTime: Date.now() - startTime + 'ms'
    };
    
  } catch (error) {
    addDiagnosticLog('診断中にエラー発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.toString(),
      diagnostics: diagnosticLogs,
      executionTime: Date.now() - startTime + 'ms'
    };
  }
}

// 既存の日付データを年月日形式に修正する関数
function fixExistingDateFormats() {
  const startTime = startPerformanceTimer();
  addLog('既存の日付データ修正を開始します');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      return {
        success: false,
        error: '機種マスタシートが見つかりません'
      };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {
        success: true,
        message: '修正対象のデータがありません',
        updatedCount: 0
      };
    }

    let updatedCount = 0;
    
    // 2行目から最後まで（ヘッダー行を除く）をチェック
    for (let row = 2; row <= lastRow; row++) {
      const createdDateCell = sheet.getRange(row, 5); // 機種ID追加により1列ずれる
      const updatedDateCell = sheet.getRange(row, 6);
      
      const createdValue = createdDateCell.getValue();
      const updatedValue = updatedDateCell.getValue();
      
      // 作成日の修正
      if (createdValue instanceof Date) {
        const dateString = Utilities.formatDate(createdValue, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        const formattedDateString = "'" + dateString; // 文字列として確実に保存
        createdDateCell.setValue(formattedDateString);
        createdDateCell.setNumberFormat('@'); // 文字列フォーマット
        updatedCount++;
        addLog(`行${row}の作成日を修正: ${createdValue} → ${dateString}`);
      }
      
      // 更新日の修正
      if (updatedValue instanceof Date) {
        const dateString = Utilities.formatDate(updatedValue, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        const formattedDateString = "'" + dateString; // 文字列として確実に保存
        updatedDateCell.setValue(formattedDateString);
        updatedDateCell.setNumberFormat('@'); // 文字列フォーマット
        updatedCount++;
        addLog(`行${row}の更新日を修正: ${updatedValue} → ${dateString}`);
      }
    }
    
    const responseTime = endPerformanceTimer(startTime, '日付データ修正');
    
    return {
      success: true,
      message: `${updatedCount}個の日付データを修正しました`,
      updatedCount: updatedCount,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '日付データ修正エラー');
    addLog('日付データ修正でエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack
      }
    };
  }
}

function diagnoseSummarySheet() {
  const startTime = Date.now();
  const diagnosticLogs = [];
  
  function addDiagnosticLog(message, data = {}) {
    const logEntry = {
      timestamp: new Date().toISOString(),
      message: message,
      data: data
    };
    diagnosticLogs.push(logEntry);
    console.log(`[DIAGNOSTIC] ${message}`, data);
  }
  
  try {
    addDiagnosticLog('サマリーシート診断開始');
    
    // スプレッドシートID取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID_DESTINATION');
    
    addDiagnosticLog('スプレッドシートID確認', {
      spreadsheetId: spreadsheetId ? '設定済み' : '未設定',
      idLength: spreadsheetId ? spreadsheetId.length : 0
    });
    
    if (!spreadsheetId) {
      throw new Error('SPREADSHEET_ID_DESTINATIONが設定されていません');
    }
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    addDiagnosticLog('スプレッドシート取得成功', {
      name: spreadsheet.getName(),
      url: spreadsheet.getUrl()
    });
    
    // 全シート取得
    const allSheets = spreadsheet.getSheets();
    const sheetInfo = allSheets.map(sheet => ({
      name: sheet.getName(),
      id: sheet.getSheetId(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn(),
      hidden: sheet.isSheetHidden()
    }));
    
    addDiagnosticLog('全シート情報', {
      totalSheets: allSheets.length,
      sheets: sheetInfo
    });
    
    // サマリーシート特定
    const summarySheet = spreadsheet.getSheetByName('サマリー');
    if (!summarySheet) {
      addDiagnosticLog('エラー: サマリーシートが見つかりません', {
        availableSheets: sheetInfo.map(s => s.name)
      });
      throw new Error('「サマリー」シートが見つかりません');
    }
    
    addDiagnosticLog('サマリーシート取得成功', {
      name: summarySheet.getName(),
      id: summarySheet.getSheetId(),
      isHidden: summarySheet.isSheetHidden()
    });
    
    // データ範囲確認
    const lastRow = summarySheet.getLastRow();
    const lastColumn = summarySheet.getLastColumn();
    
    addDiagnosticLog('データ範囲確認', {
      lastRow: lastRow,
      lastColumn: lastColumn,
      isEmpty: lastRow === 0,
      totalCells: lastRow * lastColumn
    });
    
    if (lastRow === 0) {
      addDiagnosticLog('警告: サマリーシートにデータがありません');
      return {
        success: false,
        error: 'サマリーシートにデータがありません',
        diagnostics: diagnosticLogs,
        executionTime: Date.now() - startTime + 'ms'
      };
    }
    
    // 実際のデータを少し取得してみる
    const sampleRange = summarySheet.getRange(1, 1, Math.min(lastRow, 5), Math.min(lastColumn, 5));
    const sampleData = sampleRange.getValues();
    
    addDiagnosticLog('サンプルデータ取得成功', {
      sampleRows: sampleData.length,
      sampleCols: sampleData[0] ? sampleData[0].length : 0,
      firstRow: sampleData[0] || [],
      secondRow: sampleData[1] || []
    });
    
    // セルの型チェック
    let dateCount = 0;
    let errorCount = 0;
    let nullCount = 0;
    
    for (let i = 0; i < sampleData.length; i++) {
      for (let j = 0; j < sampleData[i].length; j++) {
        const cell = sampleData[i][j];
        if (cell instanceof Date) {
          dateCount++;
        } else if (cell === null || cell === undefined) {
          nullCount++;
        } else if (cell && typeof cell === 'object' && cell.toString().startsWith('#')) {
          errorCount++;
        }
      }
    }
    
    addDiagnosticLog('セル型統計', {
      dateCount: dateCount,
      errorCount: errorCount,
      nullCount: nullCount
    });
    
    // 実際のgetDestinationSheetDataを呼び出し
    addDiagnosticLog('実際の関数呼び出しテスト開始');
    const result = getDestinationSheetData('サマリー', 'diagnostic');
    
    addDiagnosticLog('実際の関数呼び出し結果', {
      success: result.success,
      hasData: result.data ? true : false,
      dataLength: result.data ? result.data.length : 0,
      error: result.error || 'なし'
    });
    
    return {
      success: true,
      message: 'サマリーシート診断完了',
      diagnostics: diagnosticLogs,
      sheetInfo: {
        name: summarySheet.getName(),
        lastRow: lastRow,
        lastColumn: lastColumn,
        sampleData: sampleData
      },
      functionResult: result,
      executionTime: Date.now() - startTime + 'ms'
    };
    
  } catch (error) {
    addDiagnosticLog('診断中にエラー発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.toString(),
      diagnostics: diagnosticLogs,
      executionTime: Date.now() - startTime + 'ms'
    };
  }
}


// ========================================
// フォーム保存先設定管理機能
// ========================================

// フォーム保存先設定を取得（拠点別対応）
function getFormStorageSettings(locationId) {
  const startTime = startPerformanceTimer();
  addLog('フォーム保存先設定取得を開始', { locationId });
  
  try {
    // 拠点IDの検証
    if (!locationId) {
      throw new Error('拠点IDが指定されていません');
    }
    
    const properties = PropertiesService.getScriptProperties();
    const settings = {
      terminal: properties.getProperty(`FORM_FOLDER_${locationId.toUpperCase()}_TERMINAL`) || '',
      printer: properties.getProperty(`FORM_FOLDER_${locationId.toUpperCase()}_PRINTER`) || '',
      other: properties.getProperty(`FORM_FOLDER_${locationId.toUpperCase()}_OTHER`) || '',
      default: properties.getProperty(`FORM_FOLDER_${locationId.toUpperCase()}_DEFAULT`) || ''
    };
    
    endPerformanceTimer(startTime, 'フォーム保存先設定取得');
    addLog('フォーム保存先設定取得完了', { locationId, settings });
    
    return settings;
  } catch (error) {
    addLog('フォーム保存先設定取得エラー', { locationId, error: error.toString() });
    throw new Error('フォーム保存先設定の取得に失敗しました: ' + error.message);
  }
}

// フォーム保存先設定を保存（拠点別対応）
function saveFormStorageSettings(locationId, settings) {
  const startTime = startPerformanceTimer();
  addLog('フォーム保存先設定保存を開始', { locationId, settings });
  
  try {
    // バリデーション
    if (!locationId) {
      throw new Error('拠点IDが指定されていません');
    }
    
    if (!settings || typeof settings !== 'object') {
      throw new Error('設定データが不正です');
    }
    
    const properties = PropertiesService.getScriptProperties();
    const locationPrefix = `FORM_FOLDER_${locationId.toUpperCase()}`;
    
    // 各カテゴリの設定を拠点別に保存
    if (settings.terminal) {
      properties.setProperty(`${locationPrefix}_TERMINAL`, settings.terminal);
    } else {
      properties.deleteProperty(`${locationPrefix}_TERMINAL`);
    }
    
    if (settings.printer) {
      properties.setProperty(`${locationPrefix}_PRINTER`, settings.printer);
    } else {
      properties.deleteProperty(`${locationPrefix}_PRINTER`);
    }
    
    if (settings.other) {
      properties.setProperty(`${locationPrefix}_OTHER`, settings.other);
    } else {
      properties.deleteProperty(`${locationPrefix}_OTHER`);
    }
    
    if (settings.default) {
      properties.setProperty(`${locationPrefix}_DEFAULT`, settings.default);
    } else {
      properties.deleteProperty(`${locationPrefix}_DEFAULT`);
    }
    
    endPerformanceTimer(startTime, 'フォーム保存先設定保存');
    addLog('フォーム保存先設定保存完了', { locationId, settings });
    
    return { success: true, message: 'フォーム保存先設定が正常に保存されました' };
  } catch (error) {
    addLog('フォーム保存先設定保存エラー', { locationId, error: error.toString() });
    throw new Error('フォーム保存先設定の保存に失敗しました: ' + error.message);
  }
}

// Google Drive フォルダIDの検証
function validateDriveFolderId(folderId) {
  const startTime = startPerformanceTimer();
  addLog('Google Drive フォルダID検証を開始', { folderId });
  
  try {
    // フォルダIDの形式チェック
    const folderIdPattern = /^[a-zA-Z0-9_-]+$/;
    if (!folderIdPattern.test(folderId)) {
      return {
        valid: false,
        error: 'フォルダIDの形式が正しくありません'
      };
    }
    
    // 実際にフォルダにアクセスして確認
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();
    
    endPerformanceTimer(startTime, 'Google Drive フォルダID検証');
    addLog('Google Drive フォルダID検証完了', { folderId, folderName });
    
    return {
      valid: true,
      name: folderName,
      id: folderId
    };
  } catch (error) {
    endPerformanceTimer(startTime, 'Google Drive フォルダID検証エラー');
    addLog('Google Drive フォルダID検証エラー', { folderId, error: error.toString() });
    
    return {
      valid: false,
      error: 'フォルダにアクセスできません: ' + error.message
    };
  }
}

// フォーム作成時に適切な保存先フォルダIDを取得（拠点別対応・エラーハンドリング強化）
function getFormStorageFolderId(locationId, deviceCategory) {
  const startTime = startPerformanceTimer();
  addLog('フォーム保存先フォルダID取得を開始', { locationId, deviceCategory });
  
  try {
    // 拠点IDの検証
    if (!locationId) {
      throw new Error('拠点IDが指定されていません');
    }
    
    const settings = getFormStorageSettings(locationId);
    let folderId = '';
    let categoryUsed = '';
    
    // デバイスカテゴリに応じてフォルダIDを決定
    switch (deviceCategory) {
      case 'SV':
      case 'CL':
        folderId = settings.terminal;
        categoryUsed = 'terminal';
        break;
      case 'プリンタ':
        folderId = settings.printer;
        categoryUsed = 'printer';
        break;
      case 'その他':
        folderId = settings.other;
        categoryUsed = 'other';
        break;
      default:
        folderId = settings.default;
        categoryUsed = 'default';
        break;
    }
    
    // フォールバック処理
    if (!folderId && categoryUsed !== 'default') {
      folderId = settings.default;
      categoryUsed = 'default';
    }
    
    // 保存先未設定エラーチェック
    if (!folderId) {
      const locationNames = {
        'osaka': '大阪',
        'kobe': '神戸',
        'himeji': '姫路'
      };
      const locationName = locationNames[locationId] || locationId;
      
      throw new Error(`${locationName}拠点の${deviceCategory}カテゴリ用フォーム保存先が設定されていません。設定画面で保存先を設定してください。`);
    }
    
    endPerformanceTimer(startTime, 'フォーム保存先フォルダID取得');
    addLog('フォーム保存先フォルダID取得完了', { locationId, deviceCategory, folderId, categoryUsed });
    
    return {
      success: true,
      folderId: folderId,
      locationId: locationId,
      category: deviceCategory,
      categoryUsed: categoryUsed,
      usedDefault: categoryUsed === 'default'
    };
  } catch (error) {
    endPerformanceTimer(startTime, 'フォーム保存先フォルダID取得エラー');
    addLog('フォーム保存先フォルダID取得エラー', { locationId, deviceCategory, error: error.toString() });
    
    return {
      success: false,
      error: error.toString(),
      locationId: locationId,
      category: deviceCategory
    };
  }
}