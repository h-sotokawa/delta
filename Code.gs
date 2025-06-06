// 拠点名の日本語表示用マッピング
const LOCATION_NAMES = {
  'osaka_desktop': '大阪(デスクトップ)',
  'osaka_notebook': '大阪(ノート、サーバー)',
  'kobe': '神戸(端末)',
  'himeji': '姫路(端末)',
  'osaka_printer': '大阪(プリンタ、その他)',
  'hyogo_printer': '兵庫(プリンタ、その他)'
};

// メインの代替機一覧シート名
const TARGET_SHEET_NAME = 'main';

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

// 拠点ごとのスプレッドシートIDをスクリプトプロパティから取得
function getSpreadsheetIdFromProperty(location) {
  const propertyKeys = {
    'osaka_desktop': 'SPREADSHEET_ID_SOURCE_OSAKA_DESKTOP',
    'osaka_notebook': 'SPREADSHEET_ID_SOURCE_OSAKA_LAPTOP',
    'kobe': 'SPREADSHEET_ID_SOURCE_KOBE',
    'himeji': 'SPREADSHEET_ID_SOURCE_HIMEJI',
    'osaka_printer': 'SPREADSHEET_ID_SOURCE_OSAKA_PRINTER',
    'hyogo_printer': 'SPREADSHEET_ID_SOURCE_HYOGO_PRINTER'
  };
  
  const propertyKey = propertyKeys[location];
  if (!propertyKey) {
    return null;
  }
  
  return PropertiesService.getScriptProperties().getProperty(propertyKey);
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

function getSpreadsheetData(location, queryType) {
  const startTime = startPerformanceTimer();
  addLog('getSpreadsheetData関数が呼び出されました', { location, queryType });
  
  try {
    // 選択された拠点のスプレッドシートIDをスクリプトプロパティから取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにスプレッドシートIDが設定されていません: ' + location);
    }

    addLog('使用するスプレッドシートID', spreadsheetId);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
    
    if (!sheet) {
      throw new Error('シート「' + TARGET_SHEET_NAME + '」が見つかりません。');
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
function getSpreadsheetDataPaginated(location, queryType, startRow = 1, maxRows = 100) {
  addLog('getSpreadsheetDataPaginated関数が呼び出されました', { location, queryType, startRow, maxRows });
  
  try {
    // 選択された拠点のスプレッドシートIDをスクリプトプロパティから取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにスプレッドシートIDが設定されていません: ' + location);
    }

    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
    
    if (!sheet) {
      throw new Error('シート「' + TARGET_SHEET_NAME + '」が見つかりません。');
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
function updateMachineStatus(rowIndex, newStatus, location) {
  const startTime = startPerformanceTimer();
  
  try {
    // location引数で指定された拠点のスプレッドシートIDをスクリプトプロパティから取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    addLog('updateMachineStatus: ステータスを変更するスプレッドシートID', spreadsheetId);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにスプレッドシートIDが設定されていません: ' + location);
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
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
function updateMultipleStatuses(updates, location) {
  const startTime = startPerformanceTimer();
  
  try {
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにスプレッドシートIDが設定されていません: ' + location);
    }
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
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
function checkDataConsistency(location) {
  try {
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにスプレッドシートIDが設定されていません: ' + location);
    }
    
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
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
    
    addLog('データ読み込み開始', {
      range: `A1:${String.fromCharCode(64 + lastColumn)}${lastRow}`,
      totalCells: lastRow * lastColumn
    });
    
    const range = sheet.getRange(1, 1, lastRow, lastColumn);
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
        availableSheets: sheetNames
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