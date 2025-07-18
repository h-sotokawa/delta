// ========================================
// ユーティリティ関数
// ========================================

// パフォーマンス監視用
let performanceMetrics = {
  totalRequests: 0,
  averageResponseTime: 0,
  lastResetTime: new Date()
};

// ログを保持する配列（本番環境では軽量化）
let serverLogs = [];

/**
 * 軽量化されたログ関数
 * @param {string} message - ログメッセージ
 * @param {*} data - ログデータ
 * @return {Object|null} ログオブジェクトまたはnull
 */
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

/**
 * パフォーマンス測定開始
 * @return {number} 開始時刻のタイムスタンプ
 */
function startPerformanceTimer() {
  return Date.now();
}

/**
 * パフォーマンス測定終了
 * @param {number} startTime - 開始時刻
 * @param {string} operation - 操作名
 * @return {number} 処理時間（ミリ秒）
 */
function endPerformanceTimer(startTime, operation) {
  const duration = Date.now() - startTime;
  performanceMetrics.totalRequests++;
  performanceMetrics.averageResponseTime = 
    (performanceMetrics.averageResponseTime * (performanceMetrics.totalRequests - 1) + duration) / performanceMetrics.totalRequests;
  
  addLog(`パフォーマンス: ${operation}`, `${duration}ms`);
  return duration;
}

/**
 * パフォーマンス統計を取得
 * @return {Object} パフォーマンス統計
 */
function getPerformanceStats() {
  return {
    success: true,
    stats: performanceMetrics
  };
}

/**
 * パフォーマンス統計をリセット
 * @return {Object} リセット結果
 */
function resetPerformanceStats() {
  performanceMetrics = {
    totalRequests: 0,
    averageResponseTime: 0,
    lastResetTime: new Date()
  };
  return { success: true, message: 'パフォーマンス統計をリセットしました' };
}

/**
 * 高速化された日付フォーマット関数
 * @param {Date} date - フォーマットする日付
 * @return {string} フォーマットされた日付文字列
 */
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

/**
 * デバイスタイプに応じてマスタシートを決定する関数（既存の互換性維持）
 * @param {string} deviceType - デバイスタイプ
 * @return {string} シート名
 */
function getTargetSheetName(deviceType) {
  // 統合ビューシートのチェック
  if (deviceType === 'INTEGRATED_VIEW_TERMINAL') {
    return VIEW_SHEET_TYPES.INTEGRATED_TERMINAL;
  } else if (deviceType === 'INTEGRATED_VIEW_PRINTER_OTHER') {
    return VIEW_SHEET_TYPES.INTEGRATED_PRINTER_OTHER;
  } else if (deviceType === 'INTEGRATED_VIEW') {
    return VIEW_SHEET_TYPES.INTEGRATED; // 旧統合ビュー
  } else if (deviceType === 'SUMMARY_VIEW') {
    return VIEW_SHEET_TYPES.SUMMARY;
  }
  
  // 通常のマスタシート
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

/**
 * 詳細カテゴリに応じてマスタシートを決定する関数（フォーム作成用）
 * @param {string} deviceCategory - デバイスカテゴリ
 * @return {string} シート名
 */
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

/**
 * 拠点一覧を取得する関数
 * @return {Object} 拠点一覧
 */
function getLocations() {
  return {
    success: true,
    locations: getLocationNamesMapping()
  };
}

/**
 * システムヘルスチェック
 * @return {Object} ヘルスチェック結果
 */
function systemHealthCheck() {
  const startTime = startPerformanceTimer();
  const healthStatus = {
    timestamp: new Date().toISOString(),
    status: 'healthy',
    checks: []
  };
  
  try {
    // スプレッドシートアクセスチェック
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (spreadsheetId) {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      healthStatus.checks.push({
        name: 'Spreadsheet Access',
        status: 'ok',
        details: `Spreadsheet: ${spreadsheet.getName()}`
      });
    } else {
      healthStatus.checks.push({
        name: 'Spreadsheet Access',
        status: 'error',
        details: 'SPREADSHEET_ID_DESTINATION not configured'
      });
      healthStatus.status = 'unhealthy';
    }
    
    // 拠点マスタチェック
    try {
      const locations = getLocationMaster();
      healthStatus.checks.push({
        name: 'Location Master',
        status: 'ok',
        details: `${locations.length} locations found`
      });
    } catch (error) {
      healthStatus.checks.push({
        name: 'Location Master',
        status: 'error',
        details: error.toString()
      });
      healthStatus.status = 'unhealthy';
    }
    
    // パフォーマンス統計
    healthStatus.checks.push({
      name: 'Performance Metrics',
      status: 'ok',
      details: performanceMetrics
    });
    
    const duration = endPerformanceTimer(startTime, 'ヘルスチェック');
    healthStatus.executionTime = duration;
    
    return {
      success: true,
      health: healthStatus
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'ヘルスチェックエラー');
    return {
      success: false,
      error: error.toString(),
      health: {
        timestamp: new Date().toISOString(),
        status: 'error',
        error: error.toString()
      }
    };
  }
}

// ========================================
// HTMLインクルード機能
// ========================================

/**
 * HTMLファイルをインクルードする関数
 * @param {string} filename - インクルードするファイル名
 * @return {string} HTMLコンテンツ
 */
function include(filename) {
  try {
    // ファイル名に拡張子が含まれていない場合は.htmlを追加
    const fileToInclude = filename.includes('.') ? filename : filename + '.html';
    
    // HtmlServiceを使用してファイルを読み込み
    const content = HtmlService.createHtmlOutputFromFile(fileToInclude).getContent();
    
    addLog('HTMLファイルインクルード成功', { filename: fileToInclude });
    
    return content;
  } catch (error) {
    addLog('HTMLファイルインクルードエラー', { 
      filename: filename, 
      error: error.toString() 
    });
    
    // エラーの場合はHTMLコメントとしてエラー情報を返す
    return `<!-- Error loading ${filename}: ${error.toString()} -->`;
  }
}