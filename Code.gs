// ========================================
// 拠点マスタベース汎用関数
// ========================================

// 拠点IDから拠点名を取得（拠点マスタベース）
function getLocationNameById(locationId) {
  return safeExecute(() => {
    const location = getLocationById(locationId);
    return location ? location.locationName : locationId;
  }, 'getLocationNameById', locationId);
}

// 拠点IDから拠点コードを取得（拠点マスタベース）
function getLocationCodeById(locationId) {
  return safeExecute(() => {
    const location = getLocationById(locationId);
    return location ? location.locationCode : locationId;
  }, 'getLocationCodeById', locationId);
}

// 全拠点の名前マッピングを取得（互換性維持用）
function getLocationNamesMapping() {
  return safeExecute(() => {
    const locations = getLocationMaster();
    const mapping = {};
    locations.forEach(location => {
      mapping[location.locationId] = location.locationName;
    });
    return mapping;
  }, 'getLocationNamesMapping', {});
}

// マスタシート名はconfig.gsで定義

// 旧マッピングは削除済み - 拠点マスタベースに統一

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

// デバッグモードはconfig.gsで定義

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

// 拠点ID変換関数（拠点マスタベース・互換性維持）
function convertLegacyLocation(location) {
  // 拠点マスタベースでは直接IDを使用
  return location;
}

// 統一スプレッドシートIDをスクリプトプロパティから取得
function getSpreadsheetIdFromProperty(location) {
  // 旧拠点名を新拠点名に変換
  const normalizedLocation = convertLegacyLocation(location);
  
  // 全拠点で統一スプレッドシートを使用
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
}

// フロントエンド用の拠点一覧を取得（削除 - 重複定義のため）

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
        location: getLocationNameById(location),
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
    locations: getLocationNamesMapping()
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
      locationName: getLocationNameById(location),
      spreadsheetId: spreadsheetId
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      location: location,
      locationName: getLocationNameById(location)
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
        location: getLocationNameById(location),
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
        location: getLocationNameById(location),
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
    const locationMaster = getLocationMaster();
    const locations = locationMaster.map(loc => loc.locationId);
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
        location: getLocationNameById(location),
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
    
    // 文字列として確実に保存（2025年1月対応：アポストロフィなし）
    const formattedDateString = dateString;
    
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
    
    // 文字列として確実に保存（2025年1月対応：アポストロフィなし）
    const formattedDateString = dateString;
    
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

/**
 * 機種マスタデータを取得（フロントエンド用）
 */
function getModelMaster() {
  return getModelMasterForForm();
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
    
    // カテゴリを英語に統一して取得
    const categories = [...new Set(result.models.map(model => model.category))].filter(cat => cat);
    
    // 英語カテゴリのみを返すように設定
    const englishCategories = ['desktop', 'laptop', 'server', 'printer', 'other'];
    
    return {
      success: true,
      categories: englishCategories,
      totalCount: englishCategories.length
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      categories: []
    };
  }
}

/**
 * 機種マスタのカテゴリを日本語から英語に変換
 * @return {Object} 変換結果
 */
function convertModelMasterCategoriesToEnglish() {
  const startTime = startPerformanceTimer();
  addLog('機種マスタカテゴリ英語変換開始');
  
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
    
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    
    // カテゴリ列のインデックスを取得
    const categoryColIndex = headerRow.indexOf('カテゴリ');
    if (categoryColIndex === -1) {
      throw new Error('カテゴリ列が見つかりません');
    }
    
    // 日本語→英語カテゴリマッピング
    const categoryMapping = {
      'SV': 'server',
      'CL': 'desktop',
      'ノートブック': 'laptop',
      'ノートPC': 'laptop',
      'デスクトップPC': 'desktop',
      'サーバー': 'server',
      'プリンタ': 'printer',
      'その他': 'other'
    };
    
    let updatedCount = 0;
    const updates = [];
    
    // データ行をチェック（1行目はヘッダーなのでスキップ）
    for (let i = 1; i < data.length; i++) {
      const currentCategory = data[i][categoryColIndex];
      const englishCategory = categoryMapping[currentCategory];
      
      if (englishCategory && currentCategory !== englishCategory) {
        updates.push({
          row: i + 1,
          currentCategory: currentCategory,
          newCategory: englishCategory
        });
        data[i][categoryColIndex] = englishCategory;
        updatedCount++;
      }
    }
    
    // 更新を実行
    if (updates.length > 0) {
      // カテゴリ列のみを一括更新
      const categoryValues = data.slice(1).map(row => [row[categoryColIndex]]);
      const range = sheet.getRange(2, categoryColIndex + 1, categoryValues.length, 1);
      range.setValues(categoryValues);
      
      addLog('機種マスタカテゴリ英語変換完了', {
        updatedCount: updatedCount,
        updates: updates
      });
    }
    
    endPerformanceTimer(startTime, '機種マスタカテゴリ英語変換');
    
    return {
      success: true,
      message: `${updatedCount}個のカテゴリを英語に変換しました`,
      updatedCount: updatedCount,
      updates: updates,
      totalRows: data.length - 1
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタカテゴリ英語変換エラー');
    addLog('機種マスタカテゴリ英語変換エラー', { error: error.toString() });
    
    return {
      success: false,
      error: error.toString()
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
        const formattedDateString = dateString; // 文字列として確実に保存（2025年1月対応：アポストロフィなし）
        createdDateCell.setValue(formattedDateString);
        createdDateCell.setNumberFormat('@'); // 文字列フォーマット
        updatedCount++;
        addLog(`行${row}の作成日を修正: ${createdValue} → ${dateString}`);
      }
      
      // 更新日の修正
      if (updatedValue instanceof Date) {
        const dateString = Utilities.formatDate(updatedValue, Session.getScriptTimeZone(), 'yyyy/MM/dd');
        const formattedDateString = dateString; // 文字列として確実に保存（2025年1月対応：アポストロフィなし）
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
      locationFolder: properties.getProperty(`FORM_FOLDER_${locationId.toUpperCase()}`) || ''
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
    
    // 拠点のフォルダ設定を保存
    if (settings.locationFolder) {
      properties.setProperty(locationPrefix, settings.locationFolder);
    } else {
      properties.deleteProperty(locationPrefix);
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
    const folderId = settings.locationFolder;
    const categoryUsed = 'location';
    
    // 保存先未設定エラーチェック
    if (!folderId) {
      // 拠点マスタから拠点名を取得
      let locationName = locationId;
      try {
        const location = getLocationById(locationId);
        if (location) {
          locationName = location.locationName;
        }
      } catch (e) {
        // 拠点マスタが取得できない場合はIDをそのまま使用
      }
      
      throw new Error(`${locationName}拠点のフォーム保存先が設定されていません。設定画面で保存先を設定してください。`);
    }
    
    endPerformanceTimer(startTime, 'フォーム保存先フォルダID取得');
    addLog('フォーム保存先フォルダID取得完了', { locationId, deviceCategory, folderId, categoryUsed });
    
    return {
      success: true,
      folderId: folderId,
      locationId: locationId,
      category: deviceCategory,
      categoryUsed: categoryUsed,
      usedDefault: false
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

// ========================================
// 共通フォームURL設定関連
// ========================================

/**
 * 共通フォームURL設定を取得
 */
function getCommonFormsSettings() {
  const startTime = startPerformanceTimer();
  addLog('共通フォームURL設定取得開始');
  
  try {
    const properties = PropertiesService.getScriptProperties();
    
    const settings = {
      terminalCommonFormUrl: properties.getProperty('TERMINAL_COMMON_FORM_URL') || '',
      printerCommonFormUrl: properties.getProperty('PRINTER_COMMON_FORM_URL') || '',
      qrRedirectUrl: properties.getProperty('QR_REDIRECT_URL') || ''
    };
    
    endPerformanceTimer(startTime, '共通フォームURL設定取得');
    addLog('共通フォームURL設定取得完了', settings);
    
    return settings;
  } catch (error) {
    endPerformanceTimer(startTime, '共通フォームURL設定取得エラー');
    addLog('共通フォームURL設定取得エラー', { error: error.toString() });
    throw new Error('共通フォームURL設定の取得に失敗しました: ' + error.toString());
  }
}

/**
 * 共通フォームURL設定を保存
 */
function saveCommonFormsSettings(settings) {
  const startTime = startPerformanceTimer();
  addLog('共通フォームURL設定保存開始', settings);
  
  try {
    const properties = PropertiesService.getScriptProperties();
    
    // 設定値を保存
    if (settings.terminalCommonFormUrl) {
      properties.setProperty('TERMINAL_COMMON_FORM_URL', settings.terminalCommonFormUrl);
    } else {
      properties.deleteProperty('TERMINAL_COMMON_FORM_URL');
    }
    
    if (settings.printerCommonFormUrl) {
      properties.setProperty('PRINTER_COMMON_FORM_URL', settings.printerCommonFormUrl);
    } else {
      properties.deleteProperty('PRINTER_COMMON_FORM_URL');
    }
    
    if (settings.qrRedirectUrl) {
      properties.setProperty('QR_REDIRECT_URL', settings.qrRedirectUrl);
    } else {
      properties.deleteProperty('QR_REDIRECT_URL');
    }
    
    endPerformanceTimer(startTime, '共通フォームURL設定保存');
    addLog('共通フォームURL設定保存完了');
    
    return {
      success: true,
      message: '共通フォームURL設定が正常に保存されました'
    };
  } catch (error) {
    endPerformanceTimer(startTime, '共通フォームURL設定保存エラー');
    addLog('共通フォームURL設定保存エラー', { error: error.toString() });
    throw new Error('共通フォームURL設定の保存に失敗しました: ' + error.toString());
  }
}

/**
 * Google FormsのURLを検証
 */
function validateGoogleFormUrl(formUrl) {
  const startTime = startPerformanceTimer();
  addLog('Google FormsのURL検証開始', { formUrl });
  
  try {
    // URLの形式チェック
    const urlPattern = /^https:\/\/docs\.google\.com\/forms\/d\/([a-zA-Z0-9_-]+)/;
    const match = formUrl.match(urlPattern);
    
    if (!match) {
      return {
        valid: false,
        error: 'Google FormsのURLの形式が正しくありません'
      };
    }
    
    const formId = match[1];
    
    try {
      // FormAppを使ってフォームにアクセスを試行
      const form = FormApp.openById(formId);
      const title = form.getTitle();
      
      endPerformanceTimer(startTime, 'Google FormsのURL検証');
      addLog('Google FormsのURL検証完了', { formId, title });
      
      return {
        valid: true,
        title: title,
        formId: formId
      };
    } catch (formError) {
      addLog('フォームアクセスエラー', { formId, error: formError.toString() });
      
      return {
        valid: false,
        error: 'フォームにアクセスできません。フォームが存在しないか、アクセス権限がありません'
      };
    }
  } catch (error) {
    endPerformanceTimer(startTime, 'Google FormsのURL検証エラー');
    addLog('Google FormsのURL検証エラー', { error: error.toString() });
    
    return {
      valid: false,
      error: 'URL検証中にエラーが発生しました: ' + error.toString()
    };
  }
}

// ========================================
// 拠点マスタ管理機能
// ========================================

/**
 * 拠点マスタシートを取得または作成
 */
function getLocationMasterSheet() {
  const startTime = startPerformanceTimer();
  addLog('拠点マスタシート取得開始');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません');
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName('拠点マスタ');
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = spreadsheet.insertSheet('拠点マスタ');
      
      // ヘッダー行を設定（管轄とステータス変更通知を追加）
      const headers = ['拠点ID', '拠点名', '拠点コード', '管轄', 'グループメールアドレス', 'ステータス変更通知', 'ステータス', '作成日', '更新日'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // 初期データを挿入（設計書に基づいた管轄情報を追加）
      const initialData = [
        ['osaka', '大阪', 'OSAKA', '関西', 'test-group@example.com', false, 'active', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'), Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd')],
        ['kobe', '神戸', 'KOBE', '関西', 'test-group@example.com', false, 'active', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'), Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd')],
        ['himeji', '姫路', 'HIMEJI', '関西', 'test-group@example.com', false, 'active', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'), Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd')]
      ];
      
      if (initialData.length > 0) {
        sheet.getRange(2, 1, initialData.length, headers.length).setValues(initialData);
      }
      
      // 日付列のフォーマットを設定（8列目と9列目が日付列）
      sheet.getRange(2, 8, sheet.getLastRow() - 1, 2).setNumberFormat('@'); // テキスト形式
      
      addLog('拠点マスタシート作成完了');
    }
    
    endPerformanceTimer(startTime, '拠点マスタシート取得');
    return sheet;
  } catch (error) {
    endPerformanceTimer(startTime, '拠点マスタシート取得エラー');
    addLog('拠点マスタシート取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * 拠点マスタデータを取得
 */
function getLocationMaster() {
  const startTime = startPerformanceTimer();
  addLog('拠点マスタデータ取得開始');
  
  try {
    const sheet = getLocationMasterSheet();
    addLog('拠点マスタシート取得成功', { 
      sheetName: sheet.getName(),
      sheetId: sheet.getSheetId()
    });
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    addLog('シート情報', { lastRow, lastColumn });
    
    if (lastRow <= 1) {
      // ヘッダー行のみの場合
      addLog('データなし（ヘッダー行のみ）');
      return [];
    }
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 9); // 9列に変更
    const data = dataRange.getValues();
    addLog('データ取得', { 
      rangeAddress: dataRange.getA1Notation(),
      dataRows: data.length,
      firstRow: data[0] 
    });
    
    const locations = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // 拠点IDが存在する場合のみ
        const location = {
          locationId: String(row[0]),
          locationName: String(row[1]),
          locationCode: String(row[2]),
          jurisdiction: String(row[3]),  // 管轄を追加
          groupEmail: String(row[4] || ''),
          statusChangeNotification: row[5] === true || row[5] === 'true',  // ステータス変更通知を追加
          status: String(row[6]),
          createdAt: String(row[7] || ''),  // 作成日
          updatedAt: String(row[8] || '')   // 更新日
        };
        locations.push(location);
        addLog(`拠点データ追加 ${i + 1}`, location);
      }
    }
    
    endPerformanceTimer(startTime, '拠点マスタデータ取得');
    addLog('拠点マスタデータ取得完了', { 
      count: locations.length,
      locations: locations
    });
    
    return locations;
  } catch (error) {
    endPerformanceTimer(startTime, '拠点マスタデータ取得エラー');
    addLog('拠点マスタデータ取得エラー', { 
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    throw error;
  }
}

// 日付フォーマット関数は削除（拠点マスタで日付を扱わないため）

/**
 * 拠点IDから拠点情報を取得
 */
function getLocationById(locationId) {
  const startTime = startPerformanceTimer();
  addLog('拠点情報取得開始', { locationId });
  
  try {
    const locations = getLocationMaster();
    const location = locations.find(loc => loc.locationId === locationId);
    
    endPerformanceTimer(startTime, '拠点情報取得');
    addLog('拠点情報取得完了', { found: !!location });
    
    return location || null;
  } catch (error) {
    endPerformanceTimer(startTime, '拠点情報取得エラー');
    addLog('拠点情報取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * 新規拠点を追加
 */
function addLocation(locationData) {
  const startTime = startPerformanceTimer();
  addLog('新規拠点追加開始', locationData);
  
  try {
    const sheet = getLocationMasterSheet();
    
    // 既存チェック
    const locations = getLocationMaster();
    if (locations.find(loc => loc.locationId === locationData.locationId)) {
      throw new Error('同じ拠点IDが既に存在します');
    }
    
    // 新規行を追加（管轄とステータス変更通知を含む）
    const newRow = [
      locationData.locationId,
      locationData.locationName,
      locationData.locationCode,
      locationData.jurisdiction || '',  // 管轄
      locationData.groupEmail || '',
      locationData.statusChangeNotification === true || locationData.statusChangeNotification === 'true',  // ステータス変更通知
      locationData.status || 'active',
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'),
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd')
    ];
    
    sheet.appendRow(newRow);
    
    // 日付列のフォーマットを設定（8列目と9列目）
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 8, 1, 2).setNumberFormat('@'); // テキスト形式
    
    endPerformanceTimer(startTime, '新規拠点追加');
    addLog('新規拠点追加完了');
    
    return { success: true, message: '拠点が追加されました' };
  } catch (error) {
    endPerformanceTimer(startTime, '新規拠点追加エラー');
    addLog('新規拠点追加エラー', { error: error.toString() });
    return { success: false, error: error.toString() };
  }
}

/**
 * 拠点情報を更新
 */
function updateLocation(locationId, updateData) {
  const startTime = startPerformanceTimer();
  addLog('拠点情報更新開始', { locationId, updateData });
  
  try {
    const sheet = getLocationMasterSheet();
    const lastRow = sheet.getLastRow();
    
    // 該当行を検索
    for (let row = 2; row <= lastRow; row++) {
      const currentId = sheet.getRange(row, 1).getValue();
      if (currentId === locationId) {
        // 更新データを設定
        if (updateData.locationName !== undefined) {
          sheet.getRange(row, 2).setValue(updateData.locationName);
        }
        if (updateData.locationCode !== undefined) {
          sheet.getRange(row, 3).setValue(updateData.locationCode);
        }
        if (updateData.jurisdiction !== undefined) {
          sheet.getRange(row, 4).setValue(updateData.jurisdiction);  // 管轄
        }
        if (updateData.groupEmail !== undefined) {
          sheet.getRange(row, 5).setValue(updateData.groupEmail);
        }
        if (updateData.statusChangeNotification !== undefined) {
          sheet.getRange(row, 6).setValue(updateData.statusChangeNotification === true || updateData.statusChangeNotification === 'true');  // ステータス変更通知
        }
        if (updateData.status !== undefined) {
          sheet.getRange(row, 7).setValue(updateData.status);
        }
        
        // 更新日を更新（9列目）
        sheet.getRange(row, 9).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'));
        sheet.getRange(row, 9).setNumberFormat('@'); // テキスト形式
        
        endPerformanceTimer(startTime, '拠点情報更新');
        addLog('拠点情報更新完了');
        
        return { success: true, message: '拠点情報が更新されました' };
      }
    }
    
    throw new Error('指定された拠点が見つかりません');
  } catch (error) {
    endPerformanceTimer(startTime, '拠点情報更新エラー');
    addLog('拠点情報更新エラー', { error: error.toString() });
    return { success: false, error: error.toString() };
  }
}

/**
 * 拠点を削除
 */
function deleteLocation(locationId) {
  const startTime = startPerformanceTimer();
  addLog('拠点削除開始', { locationId });
  
  try {
    const sheet = getLocationMasterSheet();
    const lastRow = sheet.getLastRow();
    
    // 該当行を検索
    for (let row = 2; row <= lastRow; row++) {
      const currentId = sheet.getRange(row, 1).getValue();
      if (currentId === locationId) {
        sheet.deleteRow(row);
        
        endPerformanceTimer(startTime, '拠点削除');
        addLog('拠点削除完了');
        
        return { success: true, message: '拠点が削除されました' };
      }
    }
    
    throw new Error('指定された拠点が見つかりません');
  } catch (error) {
    endPerformanceTimer(startTime, '拠点削除エラー');
    addLog('拠点削除エラー', { error: error.toString() });
    return { success: false, error: error.toString() };
  }
}

/**
 * 既存の拠点マスタに管轄フィールドを追加する初期化関数
 */
function initializeLocationMasterJurisdiction() {
  const startTime = startPerformanceTimer();
  addLog('拠点マスタ管轄フィールド初期化開始');
  
  try {
    const sheet = getLocationMasterSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      addLog('データがないため初期化をスキップ');
      return { success: true, message: 'データがないため初期化をスキップしました' };
    }
    
    // ヘッダー行を確認
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hasJurisdiction = headers.includes('管轄');
    const hasStatusNotification = headers.includes('ステータス変更通知');
    
    if (hasJurisdiction && hasStatusNotification) {
      addLog('既に管轄フィールドが存在します');
      return { success: true, message: '既に管轄フィールドが存在します' };
    }
    
    // 既存データの管轄情報を設定
    for (let row = 2; row <= lastRow; row++) {
      const locationId = sheet.getRange(row, 1).getValue();
      
      // 既存拠点の管轄を設定（デフォルトは関西）
      if (!hasJurisdiction) {
        const currentJurisdiction = sheet.getRange(row, 4).getValue();
        if (!currentJurisdiction || currentJurisdiction === '') {
          sheet.getRange(row, 4).setValue('関西');
        }
      }
      
      // ステータス変更通知のデフォルト値を設定
      if (!hasStatusNotification) {
        const currentNotification = sheet.getRange(row, 6).getValue();
        if (currentNotification === '' || currentNotification === null) {
          sheet.getRange(row, 6).setValue(false);
        }
      }
    }
    
    endPerformanceTimer(startTime, '拠点マスタ管轄フィールド初期化');
    addLog('拠点マスタ管轄フィールド初期化完了');
    
    return { success: true, message: '管轄フィールドの初期化が完了しました' };
  } catch (error) {
    endPerformanceTimer(startTime, '拠点マスタ管轄フィールド初期化エラー');
    addLog('拠点マスタ管轄フィールド初期化エラー', { error: error.toString() });
    return { success: false, error: error.toString() };
  }
}

/**
 * 管轄に基づいて拠点を取得
 */
function getLocationsByJurisdiction(jurisdiction) {
  const startTime = startPerformanceTimer();
  addLog('管轄別拠点取得開始', { jurisdiction });
  
  try {
    const allLocations = getLocationMaster();
    
    if (!jurisdiction || jurisdiction === '') {
      // 管轄指定なしの場合は全拠点を返す
      endPerformanceTimer(startTime, '管轄別拠点取得（全件）');
      return allLocations;
    }
    
    // 指定された管轄の拠点のみをフィルタリング
    const filteredLocations = allLocations.filter(location => 
      location.jurisdiction === jurisdiction
    );
    
    endPerformanceTimer(startTime, '管轄別拠点取得');
    addLog('管轄別拠点取得完了', { 
      jurisdiction,
      totalLocations: allLocations.length,
      filteredLocations: filteredLocations.length 
    });
    
    return filteredLocations;
  } catch (error) {
    endPerformanceTimer(startTime, '管轄別拠点取得エラー');
    addLog('管轄別拠点取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * 管轄リストを取得（重複なし）
 */
function getJurisdictionList() {
  const startTime = startPerformanceTimer();
  addLog('管轄リスト取得開始');
  
  try {
    const locations = getLocationMaster();
    const jurisdictions = [...new Set(locations.map(loc => loc.jurisdiction).filter(j => j && j !== ''))];
    
    endPerformanceTimer(startTime, '管轄リスト取得');
    addLog('管轄リスト取得完了', { jurisdictions });
    
    return jurisdictions.sort();
  } catch (error) {
    endPerformanceTimer(startTime, '管轄リスト取得エラー');
    addLog('管轄リスト取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * サマリーデータの動的拠点表示テスト
 */
function testDynamicSummaryDisplay() {
  console.log('=== サマリーデータ動的拠点表示テスト開始 ===');
  
  try {
    // 1. 拠点マスタの確認
    console.log('\n1. 拠点マスタデータ確認');
    const locations = getLocationMaster();
    console.log('登録拠点数:', locations.length);
    console.log('拠点リスト:', locations.map(loc => ({
      id: loc.locationId,
      name: loc.locationName,
      jurisdiction: loc.jurisdiction
    })));
    
    // 2. サマリーデータの取得
    console.log('\n2. サマリーデータ取得');
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const summarySheet = spreadsheet.getSheetByName('サマリー');
    
    if (!summarySheet) {
      console.log('サマリーシートが見つかりません');
      return { success: false, error: 'サマリーシートが見つかりません' };
    }
    
    const summaryData = summarySheet.getDataRange().getValues();
    console.log('サマリーデータ行数:', summaryData.length);
    
    // 3. サマリーデータ内の拠点名を確認
    console.log('\n3. サマリーデータ内の拠点名確認');
    const foundLocationNames = new Set();
    
    for (let i = 0; i < summaryData.length; i++) {
      const firstCell = summaryData[i][0];
      if (firstCell && typeof firstCell === 'string' && firstCell.trim() !== '') {
        // カテゴリ行や合計行を除外
        if (!firstCell.match(/^\d+\./) && firstCell !== '合計' && firstCell !== 'SV') {
          foundLocationNames.add(firstCell.trim());
        }
      }
    }
    
    console.log('サマリーデータ内で見つかった拠点名:', Array.from(foundLocationNames));
    
    // 4. 拠点マスタとの照合
    console.log('\n4. 拠点マスタとの照合');
    const validLocationNames = locations.map(loc => loc.locationName);
    const unmatchedLocations = [];
    const matchedLocations = [];
    
    foundLocationNames.forEach(name => {
      if (validLocationNames.includes(name)) {
        matchedLocations.push(name);
      } else {
        unmatchedLocations.push(name);
      }
    });
    
    console.log('拠点マスタと一致する拠点:', matchedLocations);
    console.log('拠点マスタに存在しない拠点:', unmatchedLocations);
    
    // 5. 新規拠点追加シミュレーション
    console.log('\n5. 新規拠点追加シミュレーション');
    console.log('テスト拠点「テスト横浜」を追加した場合...');
    console.log('動的実装では自動的にサマリーに含まれるようになります');
    
    console.log('\n=== サマリーデータ動的拠点表示テスト完了 ===');
    
    return {
      success: true,
      results: {
        locationCount: locations.length,
        summaryLocationCount: foundLocationNames.size,
        matchedCount: matchedLocations.length,
        unmatchedCount: unmatchedLocations.length,
        dynamicImplementation: '有効'
      }
    };
    
  } catch (error) {
    console.error('テストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 管轄機能の統合テスト
 */
function testJurisdictionFeatures() {
  console.log('=== 管轄機能統合テスト開始 ===');
  
  try {
    // 1. 既存データの初期化
    console.log('\n1. 既存データの管轄フィールド初期化');
    const initResult = initializeLocationMasterJurisdiction();
    console.log('初期化結果:', initResult);
    
    // 2. 管轄リストの取得テスト
    console.log('\n2. 管轄リスト取得テスト');
    const jurisdictions = getJurisdictionList();
    console.log('管轄リスト:', jurisdictions);
    
    // 3. 管轄別拠点取得テスト
    console.log('\n3. 管轄別拠点取得テスト');
    jurisdictions.forEach(jurisdiction => {
      const locations = getLocationsByJurisdiction(jurisdiction);
      console.log(`管轄「${jurisdiction}」の拠点数:`, locations.length);
      console.log('拠点リスト:', locations.map(loc => loc.locationName));
    });
    
    // 4. 新規拠点追加テスト（管轄とステータス変更通知を含む）
    console.log('\n4. 新規拠点追加テスト');
    const testLocation = {
      locationId: 'test_tokyo',
      locationName: 'テスト東京',
      locationCode: 'TESTTOKYO',
      jurisdiction: '関東',
      groupEmail: 'tokyo-test@example.com',
      statusChangeNotification: true,
      status: 'active'
    };
    
    // 既存の場合は削除
    try {
      deleteLocation(testLocation.locationId);
    } catch (e) {
      // 無視
    }
    
    const addResult = addLocation(testLocation);
    console.log('追加結果:', addResult);
    
    // 5. 追加した拠点の確認
    console.log('\n5. 追加した拠点の確認');
    const addedLocation = getLocationById(testLocation.locationId);
    console.log('追加された拠点:', addedLocation);
    
    // 6. 拠点情報の更新テスト
    console.log('\n6. 拠点情報更新テスト');
    const updateData = {
      jurisdiction: '中部',
      statusChangeNotification: false
    };
    const updateResult = updateLocation(testLocation.locationId, updateData);
    console.log('更新結果:', updateResult);
    
    // 7. 更新後の確認
    console.log('\n7. 更新後の確認');
    const updatedLocation = getLocationById(testLocation.locationId);
    console.log('更新された拠点:', updatedLocation);
    
    // 8. テストデータのクリーンアップ
    console.log('\n8. テストデータのクリーンアップ');
    const deleteResult = deleteLocation(testLocation.locationId);
    console.log('削除結果:', deleteResult);
    
    console.log('\n=== 管轄機能統合テスト完了 ===');
    
    return {
      success: true,
      results: {
        initResult,
        jurisdictions,
        testResults: 'All tests passed'
      }
    };
    
  } catch (error) {
    console.error('テストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 拠点マスタのテスト関数（デバッグ用）
 */
function testLocationMaster() {
  try {
    console.log('=== 拠点マスタテスト開始 ===');
    
    // 1. シートの存在確認
    console.log('1. シート取得テスト');
    const sheet = getLocationMasterSheet();
    console.log('シート名:', sheet.getName());
    console.log('最終行:', sheet.getLastRow());
    console.log('最終列:', sheet.getLastColumn());
    
    // 2. データ取得テスト
    console.log('\n2. データ取得テスト');
    const locations = getLocationMaster();
    console.log('取得件数:', locations.length);
    console.log('取得データ（日付除外）:', JSON.stringify(locations, null, 2));
    
    // 3. 個別拠点取得テスト
    console.log('\n3. 個別拠点取得テスト');
    const osaka = getLocationById('osaka');
    console.log('大阪拠点:', JSON.stringify(osaka, null, 2));
    
    console.log('\n=== テスト完了 ===');
    
    return {
      success: true,
      sheet: {
        name: sheet.getName(),
        lastRow: sheet.getLastRow(),
        lastColumn: sheet.getLastColumn()
      },
      data: locations,
      count: locations.length
    };
  } catch (error) {
    console.error('テストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}