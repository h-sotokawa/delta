// ========================================
// ビューシート操作関連
// 統合ビュー、検索インデックス、サマリービューの管理
// ========================================

/**
 * ビューシートの種類を定義
 */
const VIEW_SHEET_TYPES = {
  INTEGRATED: 'integrated_view',
  SEARCH_INDEX: 'search_index',
  SUMMARY: 'summary_view'
};

/**
 * 統合ビューシートを取得または作成
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 統合ビューシート
 */
function getIntegratedViewSheet() {
  const startTime = startPerformanceTimer();
  addLog('統合ビューシート取得開始');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(VIEW_SHEET_TYPES.INTEGRATED);
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = createIntegratedViewSheet(spreadsheet);
    }
    
    endPerformanceTimer(startTime, '統合ビューシート取得');
    return sheet;
  } catch (error) {
    endPerformanceTimer(startTime, '統合ビューシート取得エラー');
    addLog('統合ビューシート取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * 検索インデックスシートを取得または作成
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 検索インデックスシート
 */
function getSearchIndexSheet() {
  const startTime = startPerformanceTimer();
  addLog('検索インデックスシート取得開始');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(VIEW_SHEET_TYPES.SEARCH_INDEX);
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = createSearchIndexSheet(spreadsheet);
    }
    
    endPerformanceTimer(startTime, '検索インデックスシート取得');
    return sheet;
  } catch (error) {
    endPerformanceTimer(startTime, '検索インデックスシート取得エラー');
    addLog('検索インデックスシート取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * サマリービューシートを取得または作成
 * @return {GoogleAppsScript.Spreadsheet.Sheet} サマリービューシート
 */
function getSummaryViewSheet() {
  const startTime = startPerformanceTimer();
  addLog('サマリービューシート取得開始');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(VIEW_SHEET_TYPES.SUMMARY);
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = createSummaryViewSheet(spreadsheet);
    }
    
    endPerformanceTimer(startTime, 'サマリービューシート取得');
    return sheet;
  } catch (error) {
    endPerformanceTimer(startTime, 'サマリービューシート取得エラー');
    addLog('サマリービューシート取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * 統合ビューシートを作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - スプレッドシート
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 作成されたシート
 */
function createIntegratedViewSheet(spreadsheet) {
  addLog('統合ビューシート作成開始');
  
  const sheet = spreadsheet.insertSheet(VIEW_SHEET_TYPES.INTEGRATED);
  
  // ヘッダー行を設定
  const headers = [
    '拠点管理番号',      // A列
    'カテゴリ',          // B列
    '機種名',            // C列
    '製造番号',          // D列
    '資産管理番号',      // E列
    'ソフトウェア',      // F列
    'OS',                // G列
    '最終更新日時',      // H列
    '現在ステータス',    // I列
    '担当者',            // J列
    '顧客名',            // K列
    '顧客番号',          // L列
    '住所',              // M列
    'ユーザー機預り有無', // N列
    '預りユーザー機シリアル', // O列
    'お預かり証No',      // P列
    '社内ステータス',    // Q列
    '棚卸フラグ',        // R列
    '現在拠点',          // S列
    '備考',              // T列
    '貸出日数',          // U列
    '要注意フラグ',      // V列
    '拠点名',            // W列
    '管轄',              // X列
    'formURL',           // Y列
    'QRコードURL'        // Z列
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4472C4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // 列幅の調整
  const columnWidths = [
    150, // 拠点管理番号
    100, // カテゴリ
    150, // 機種名
    120, // 製造番号
    120, // 資産管理番号
    150, // ソフトウェア
    100, // OS
    150, // 最終更新日時
    120, // 現在ステータス
    100, // 担当者
    150, // 顧客名
    100, // 顧客番号
    200, // 住所
    120, // ユーザー機預り有無
    150, // 預りユーザー機シリアル
    120, // お預かり証No
    120, // 社内ステータス
    100, // 棚卸フラグ
    100, // 現在拠点
    200, // 備考
    80,  // 貸出日数
    100, // 要注意フラグ
    100, // 拠点名
    80,  // 管轄
    200, // formURL
    200  // QRコードURL
  ];
  
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  // 統合ビューの初期データを設定（QUERY関数を使用）
  setupIntegratedViewFormulas(sheet);
  
  addLog('統合ビューシート作成完了');
  return sheet;
}

/**
 * 検索インデックスシートを作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - スプレッドシート
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 作成されたシート
 */
function createSearchIndexSheet(spreadsheet) {
  addLog('検索インデックスシート作成開始');
  
  const sheet = spreadsheet.insertSheet(VIEW_SHEET_TYPES.SEARCH_INDEX);
  
  // ヘッダー行を設定
  const headers = [
    '拠点管理番号', // A列
    'カテゴリ',     // B列
    '機種名',       // C列
    'ステータス',   // D列
    '拠点コード',   // E列
    '管轄',         // F列
    '最終更新',     // G列
    '検索キー'      // H列
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#70AD47');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // 列幅の調整
  const columnWidths = [
    150, // 拠点管理番号
    100, // カテゴリ
    150, // 機種名
    120, // ステータス
    100, // 拠点コード
    80,  // 管轄
    150, // 最終更新
    300  // 検索キー
  ];
  
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  addLog('検索インデックスシート作成完了');
  return sheet;
}

/**
 * サマリービューシートを作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - スプレッドシート
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 作成されたシート
 */
function createSummaryViewSheet(spreadsheet) {
  addLog('サマリービューシート作成開始');
  
  const sheet = spreadsheet.insertSheet(VIEW_SHEET_TYPES.SUMMARY);
  
  // ヘッダー行を設定
  const headers = [
    '管轄',     // A列
    '拠点名',   // B列
    'カテゴリ', // C列
    '総数',     // D列
    '貸出中',   // E列
    '社内保管', // F列
    '修理中',   // G列
    'その他',   // H列
    '稼働率',   // I列
    '要注意数', // J列
    '最終更新'  // K列
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ED7D31');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // 列幅の調整
  const columnWidths = [
    80,  // 管轄
    120, // 拠点名
    100, // カテゴリ
    80,  // 総数
    80,  // 貸出中
    80,  // 社内保管
    80,  // 修理中
    80,  // その他
    80,  // 稼働率
    80,  // 要注意数
    150  // 最終更新
  ];
  
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  addLog('サマリービューシート作成完了');
  return sheet;
}

/**
 * 統合ビューシートにQUERY関数を設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 統合ビューシート
 */
function setupIntegratedViewFormulas(sheet) {
  addLog('統合ビュー数式設定開始');
  
  // 実装の簡略化のため、初期は空のままとし、
  // 実際のデータ統合は別途バッチ処理で行う
  // TODO: QUERY関数による自動統合の実装
  
  addLog('統合ビュー数式設定完了');
}

/**
 * ビューシートからデータを取得
 * @param {string} viewType - ビューシートタイプ
 * @param {Object} filters - フィルター条件
 * @return {Object} データ取得結果
 */
function getViewSheetData(viewType, filters = {}) {
  const startTime = startPerformanceTimer();
  addLog('ビューシートデータ取得開始', { viewType, filters });
  
  try {
    let sheet;
    switch (viewType) {
      case VIEW_SHEET_TYPES.INTEGRATED:
        sheet = getIntegratedViewSheet();
        break;
      case VIEW_SHEET_TYPES.SEARCH_INDEX:
        sheet = getSearchIndexSheet();
        break;
      case VIEW_SHEET_TYPES.SUMMARY:
        sheet = getSummaryViewSheet();
        break;
      default:
        throw new Error('不明なビューシートタイプ: ' + viewType);
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      return {
        success: true,
        data: [sheet.getRange(1, 1, 1, lastColumn).getValues()[0]], // ヘッダーのみ
        metadata: {
          viewType: viewType,
          totalRows: 0,
          filteredRows: 0
        }
      };
    }
    
    const data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    
    // フィルター処理（必要に応じて実装）
    let filteredData = data;
    if (filters && Object.keys(filters).length > 0) {
      filteredData = applyFiltersToData(data, filters);
    }
    
    endPerformanceTimer(startTime, 'ビューシートデータ取得');
    
    return {
      success: true,
      data: filteredData,
      metadata: {
        viewType: viewType,
        totalRows: data.length - 1,
        filteredRows: filteredData.length - 1
      }
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'ビューシートデータ取得エラー');
    addLog('ビューシートデータ取得エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * データにフィルターを適用
 * @param {Array<Array>} data - データ配列
 * @param {Object} filters - フィルター条件
 * @return {Array<Array>} フィルター後のデータ
 */
function applyFiltersToData(data, filters) {
  if (!data || data.length <= 1) return data;
  
  const headers = data[0];
  const rows = data.slice(1);
  
  const filteredRows = rows.filter(row => {
    // 管轄フィルター
    if (filters.jurisdiction) {
      const jurisdictionIndex = headers.indexOf('管轄');
      if (jurisdictionIndex >= 0 && row[jurisdictionIndex] !== filters.jurisdiction) {
        return false;
      }
    }
    
    // 拠点フィルター
    if (filters.location) {
      const locationIndex = headers.indexOf('拠点名');
      if (locationIndex >= 0 && row[locationIndex] !== filters.location) {
        return false;
      }
    }
    
    // ステータスフィルター
    if (filters.status) {
      const statusIndex = headers.indexOf('現在ステータス');
      if (statusIndex >= 0 && row[statusIndex] !== filters.status) {
        return false;
      }
    }
    
    // カテゴリフィルター
    if (filters.category) {
      const categoryIndex = headers.indexOf('カテゴリ');
      if (categoryIndex >= 0 && row[categoryIndex] !== filters.category) {
        return false;
      }
    }
    
    // 検索キーワードフィルター（検索インデックス用）
    if (filters.keyword) {
      const searchKeyIndex = headers.indexOf('検索キー');
      if (searchKeyIndex >= 0) {
        const searchKey = row[searchKeyIndex] || '';
        if (!searchKey.toLowerCase().includes(filters.keyword.toLowerCase())) {
          return false;
        }
      }
    }
    
    return true;
  });
  
  return [headers, ...filteredRows];
}

/**
 * 統合ビューのデータを更新（バッチ処理）
 * @return {Object} 更新結果
 */
function updateIntegratedView() {
  const startTime = startPerformanceTimer();
  addLog('統合ビュー更新開始');
  
  try {
    const sheet = getIntegratedViewSheet();
    
    // マスタシートからデータを収集
    const terminalData = getTerminalMasterData();
    const printerData = getPrinterMasterData();
    const otherData = getOtherMasterData();
    
    // ステータス収集シートからデータを取得
    const statusData = getLatestStatusData();
    
    // データを統合
    const integratedData = integrateAllData(
      terminalData,
      printerData,
      otherData,
      statusData
    );
    
    // シートに書き込み
    if (integratedData.length > 0) {
      sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getLastColumn()).clearContent();
      sheet.getRange(2, 1, integratedData.length, integratedData[0].length).setValues(integratedData);
    }
    
    endPerformanceTimer(startTime, '統合ビュー更新');
    
    return {
      success: true,
      rowsUpdated: integratedData.length
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '統合ビュー更新エラー');
    addLog('統合ビュー更新エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 端末マスタデータを取得
 * @return {Array} 端末マスタデータ
 */
function getTerminalMasterData() {
  return getMasterData('端末マスタ');
}

/**
 * プリンタマスタデータを取得
 * @return {Array} プリンタマスタデータ
 */
function getPrinterMasterData() {
  return getMasterData('プリンタマスタ');
}

/**
 * その他マスタデータを取得
 * @return {Array} その他マスタデータ
 */
function getOtherMasterData() {
  return getMasterData('その他マスタ');
}

/**
 * マスタデータを取得（汎用）
 * @param {string} sheetName - シート名
 * @return {Array} マスタデータ
 */
function getMasterData(sheetName) {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // オブジェクト形式に変換
    return data.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
  } catch (error) {
    console.error(`${sheetName}データ取得エラー:`, error);
    return [];
  }
}

/**
 * 最新のステータス収集データを取得
 * @return {Object} ステータスデータ（拠点管理番号をキーとしたマップ）
 */
function getLatestStatusCollectionData() {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // 各種ステータス収集シートを探す
    const statusSheets = ['端末ステータス収集', 'プリンタステータス収集', 'その他ステータス収集'];
    const statusMap = {};
    
    for (const sheetName of statusSheets) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() <= 1) continue;
      
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // 拠点管理番号の列を特定
      const managementNumberIndex = headers.indexOf('2-1.拠点管理番号');
      if (managementNumberIndex < 0) continue;
      
      // 各行を処理
      data.forEach(row => {
        const managementNumber = row[managementNumberIndex];
        if (!managementNumber) return;
        
        // 既存のデータより新しい場合のみ更新
        const timestamp = row[0]; // タイムスタンプは通常A列
        if (!statusMap[managementNumber] || 
            (timestamp && statusMap[managementNumber].timestamp < timestamp)) {
          statusMap[managementNumber] = {
            timestamp: timestamp,
            status: row[headers.indexOf('0-4.ステータス')] || '',
            assignee: row[headers.indexOf('0-1.担当者')] || '',
            customerName: row[headers.indexOf('1-1.顧客名')] || '',
            customerNumber: row[headers.indexOf('1-2.顧客番号')] || '',
            address: row[headers.indexOf('1-3.住所')] || '',
            userMachineFlag: row[headers.indexOf('1-4.ユーザー機の預り有無')] || '',
            userMachineSerial: row[headers.indexOf('1-7.預りユーザー機のシリアルNo.')] || '',
            receiptNumber: row[headers.indexOf('1-8.お預かり証No.')] || '',
            internalStatus: row[headers.indexOf('3-0.社内ステータス')] || '',
            inventoryFlag: row[headers.indexOf('3-0-1.棚卸しフラグ')] || '',
            currentLocation: row[headers.indexOf('3-0-2.拠点')] || '',
            remarks: row[headers.indexOf('備考')] || '',
            loanDate: row[headers.indexOf('1-5.貸出日')] || ''
          };
        }
      });
    }
    
    return statusMap;
  } catch (error) {
    console.error('ステータス収集データ取得エラー:', error);
    return {};
  }
}

/**
 * サマリービューデータを構築
 * @return {Object} 構築結果
 */
function buildSummaryViewData() {
  const startTime = startPerformanceTimer();
  addLog('サマリービューデータ構築開始');
  
  try {
    // 統合ビューからデータを取得
    const integratedSheet = getIntegratedViewSheet();
    const lastRow = integratedSheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: true,
        rowsCreated: 0,
        message: '統合ビューにデータがありません'
      };
    }
    
    const data = integratedSheet.getRange(1, 1, lastRow, integratedSheet.getLastColumn()).getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // 必要な列のインデックスを取得
    const columnIndices = {
      jurisdiction: headers.indexOf('管轄'),
      locationName: headers.indexOf('拠点名'),
      category: headers.indexOf('カテゴリ'),
      status: headers.indexOf('現在ステータス'),
      internalStatus: headers.indexOf('社内ステータス'),
      cautionFlag: headers.indexOf('要注意フラグ')
    };
    
    // 集計用のマップを作成
    const summaryMap = {};
    
    for (const row of rows) {
      const jurisdiction = row[columnIndices.jurisdiction] || '未設定';
      const locationName = row[columnIndices.locationName] || '未設定';
      const category = row[columnIndices.category] || '未設定';
      const status = row[columnIndices.status] || '';
      const internalStatus = row[columnIndices.internalStatus] || '';
      const cautionFlag = row[columnIndices.cautionFlag] || false;
      
      // キーを作成
      const key = `${jurisdiction}\t${locationName}\t${category}`;
      
      if (!summaryMap[key]) {
        summaryMap[key] = {
          jurisdiction: jurisdiction,
          locationName: locationName,
          category: category,
          total: 0,
          lending: 0,
          stored: 0,
          repairing: 0,
          other: 0,
          caution: 0
        };
      }
      
      // カウントを増やす
      summaryMap[key].total++;
      
      // ステータス別カウント
      if (status === '1.貸出中') {
        summaryMap[key].lending++;
      } else if (status === '3.社内にて保管中') {
        summaryMap[key].stored++;
      } else if (internalStatus === '1.修理中') {
        summaryMap[key].repairing++;
      } else {
        summaryMap[key].other++;
      }
      
      // 要注意フラグカウント
      if (cautionFlag === true || cautionFlag === 'TRUE') {
        summaryMap[key].caution++;
      }
    }
    
    // サマリーデータを配列に変換
    const summaryData = [];
    const now = new Date();
    
    for (const key in summaryMap) {
      const summary = summaryMap[key];
      const utilizationRate = summary.total > 0 ? (summary.lending / summary.total * 100).toFixed(1) + '%' : '0.0%';
      
      summaryData.push([
        summary.jurisdiction,     // A: 管轄
        summary.locationName,     // B: 拠点名
        summary.category,         // C: カテゴリ
        summary.total,            // D: 総数
        summary.lending,          // E: 貸出中
        summary.stored,           // F: 社内保管
        summary.repairing,        // G: 修理中
        summary.other,            // H: その他
        utilizationRate,          // I: 稼働率
        summary.caution,          // J: 要注意数
        now                       // K: 最終更新
      ]);
    }
    
    // 管轄、拠点名、カテゴリでソート
    summaryData.sort((a, b) => {
      if (a[0] !== b[0]) return a[0].localeCompare(b[0]); // 管轄
      if (a[1] !== b[1]) return a[1].localeCompare(b[1]); // 拠点名
      return a[2].localeCompare(b[2]); // カテゴリ
    });
    
    // サマリービューシートに書き込み
    const summarySheet = getSummaryViewSheet();
    
    // 既存データをクリア
    if (summarySheet.getLastRow() > 1) {
      summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, summarySheet.getLastColumn()).clearContent();
    }
    
    // 新しいデータを書き込み
    if (summaryData.length > 0) {
      summarySheet.getRange(2, 1, summaryData.length, summaryData[0].length).setValues(summaryData);
    }
    
    endPerformanceTimer(startTime, 'サマリービューデータ構築');
    
    const result = {
      success: true,
      rowsCreated: summaryData.length,
      timestamp: now
    };
    
    addLog('サマリービューデータ構築完了', result);
    return result;
    
  } catch (error) {
    endPerformanceTimer(startTime, 'サマリービューデータ構築エラー');
    addLog('サマリービューデータ構築エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ビューシートのテスト関数
 */
function testViewSheets() {
  console.log('=== ビューシートテスト開始 ===');
  
  // 統合ビューシート取得テスト
  const integratedSheet = getIntegratedViewSheet();
  console.log('統合ビューシート名:', integratedSheet.getName());
  
  // 検索インデックスシート取得テスト
  const searchSheet = getSearchIndexSheet();
  console.log('検索インデックスシート名:', searchSheet.getName());
  
  // サマリービューシート取得テスト
  const summarySheet = getSummaryViewSheet();
  console.log('サマリービューシート名:', summarySheet.getName());
  
  // データ取得テスト
  const integratedData = getViewSheetData(VIEW_SHEET_TYPES.INTEGRATED);
  console.log('統合ビューデータ取得結果:', integratedData);
  
  console.log('=== ビューシートテスト完了 ===');
}