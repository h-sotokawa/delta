// ========================================
// ビューシート操作関連
// 統合ビュー、検索インデックス、サマリービューの管理
// ========================================

/**
 * ビューシートの種類を定義
 */
const VIEW_SHEET_TYPES = {
  INTEGRATED: 'integrated_view',
  INTEGRATED_TERMINAL: 'integrated_view_terminal',
  INTEGRATED_PRINTER_OTHER: 'integrated_view_printer_other',
  SEARCH_INDEX: 'search_index',
  SUMMARY: 'summary_view'
};

/**
 * 統合ビューシートを取得または作成（旧関数、互換性のため維持）
 * @deprecated 端末系とプリンタ・その他系に分割されました
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 端末系統合ビューシート
 */
function getIntegratedViewSheet() {
  return getIntegratedViewTerminalSheet();
}

/**
 * 端末系統合ビューシートを取得または作成
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 端末系統合ビューシート
 */
function getIntegratedViewTerminalSheet() {
  const startTime = startPerformanceTimer();
  addLog('統合ビューシート取得開始');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(VIEW_SHEET_TYPES.INTEGRATED_TERMINAL);
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = createIntegratedViewTerminalSheet(spreadsheet);
    }
    
    endPerformanceTimer(startTime, '端末系統合ビューシート取得');
    return sheet;
  } catch (error) {
    endPerformanceTimer(startTime, '端末系統合ビューシート取得エラー');
    addLog('端末系統合ビューシート取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * プリンタ・その他系統合ビューシートを取得または作成
 * @return {GoogleAppsScript.Spreadsheet.Sheet} プリンタ・その他系統合ビューシート
 */
function getIntegratedViewPrinterOtherSheet() {
  const startTime = startPerformanceTimer();
  addLog('プリンタ・その他系統合ビューシート取得開始');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(VIEW_SHEET_TYPES.INTEGRATED_PRINTER_OTHER);
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = createIntegratedViewPrinterOtherSheet(spreadsheet);
    }
    
    endPerformanceTimer(startTime, 'プリンタ・その他系統合ビューシート取得');
    return sheet;
  } catch (error) {
    endPerformanceTimer(startTime, 'プリンタ・その他系統合ビューシート取得エラー');
    addLog('プリンタ・その他系統合ビューシート取得エラー', { error: error.toString() });
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
 * 統合ビューシートを作成（旧関数、互換性のため維持）
 * @deprecated
 */
function createIntegratedViewSheet(spreadsheet) {
  return createIntegratedViewTerminalSheet(spreadsheet);
}

/**
 * 端末系統合ビューシートを作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - スプレッドシート
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 作成されたシート
 */
function createIntegratedViewTerminalSheet(spreadsheet) {
  addLog('端末系統合ビューシート作成開始');
  
  const sheet = spreadsheet.insertSheet(VIEW_SHEET_TYPES.INTEGRATED_TERMINAL);
  
  // ヘッダー行を設定（設計書通りの47列構造）
  const headers = [
    // マスタシートデータ（A-G列）
    '拠点管理番号',                             // A列
    'カテゴリ',                                 // B列
    '機種名',                                   // C列
    '製造番号',                                 // D列
    '資産管理番号',                             // E列
    'ソフトウェア',                             // F列
    'OS',                                       // G列
    
    // 収集シートデータ（H-AN列）
    'タイムスタンプ',                           // H列
    '9999.管理ID',                              // I列
    '0-0.拠点管理番号',                         // J列
    '0-1.担当者',                               // K列
    '0-2.EMシステムズの社員ですか？',           // L列
    '0-3.所属会社',                             // M列
    '0-4.ステータス',                           // N列
    '1-1.顧客名または貸出先',                   // O列
    '1-2.顧客番号',                             // P列
    '1-3.住所',                                 // Q列
    '1-4.ユーザー機の預り有無',                 // R列
    '1-5.依頼者',                               // S列
    '1-6.備考',                                 // T列
    '1-7.預りユーザー機のシリアルNo.(製造番号)', // U列
    '1-8.お預かり証No.',                        // V列
    '2-1.預り機返却の有無',                     // W列
    '2-2.依頼者',                               // X列
    '2-3.備考',                                 // Y列
    '3-0.社内ステータス',                       // Z列
    '3-0-1.棚卸しフラグ',                       // AA列
    '3-0-2.拠点',                               // AB列
    '3-1-1.ソフト',                             // AC列
    '3-1-2.備考',                               // AD列
    '3-2-1.端末初期化の引継ぎ',                 // AE列
    '3-2-2.備考',                               // AF列
    '3-2-3.引継ぎ担当者',                       // AG列
    '3-2-4.初期化作業の引継ぎ',                 // AH列
    '4-1.所在',                                 // AI列
    '4-2.持ち出し理由',                         // AJ列
    '4-3.備考',                                 // AK列
    '5-1.内容',                                 // AL列
    '5-2.所在',                                 // AM列
    '5-3.備考',                                 // AN列
    
    // 計算フィールド（AO-AP列）
    '貸出日数',                                 // AO列
    '要注意フラグ',                             // AP列
    
    // 参照データ（AQ-AT列）
    '拠点名',                                   // AQ列
    '管轄',                                     // AR列
    'formURL',                                  // AS列
    'QRコードURL'                               // AT列
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4472C4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // 列幅の調整（47列対応）
  const columnWidths = [
    // マスタシートデータ（A-G列）
    150, // A: 拠点管理番号
    100, // B: カテゴリ
    150, // C: 機種名
    120, // D: 製造番号
    120, // E: 資産管理番号
    150, // F: ソフトウェア
    100, // G: OS
    
    // 収集シートデータ（H-AN列）
    150, // H: タイムスタンプ
    100, // I: 9999.管理ID
    150, // J: 0-0.拠点管理番号
    100, // K: 0-1.担当者
    120, // L: 0-2.EMシステムズの社員ですか？
    150, // M: 0-3.所属会社
    120, // N: 0-4.ステータス
    150, // O: 1-1.顧客名または貸出先
    100, // P: 1-2.顧客番号
    200, // Q: 1-3.住所
    120, // R: 1-4.ユーザー機の預り有無
    100, // S: 1-5.依頼者
    200, // T: 1-6.備考
    150, // U: 1-7.預りユーザー機のシリアルNo.
    120, // V: 1-8.お預かり証No.
    120, // W: 2-1.預り機返却の有無
    100, // X: 2-2.依頼者
    200, // Y: 2-3.備考
    120, // Z: 3-0.社内ステータス
    100, // AA: 3-0-1.棚卸しフラグ
    100, // AB: 3-0-2.拠点
    150, // AC: 3-1-1.ソフト
    200, // AD: 3-1-2.備考
    150, // AE: 3-2-1.端末初期化の引継ぎ
    200, // AF: 3-2-2.備考
    120, // AG: 3-2-3.引継ぎ担当者
    150, // AH: 3-2-4.初期化作業の引継ぎ
    120, // AI: 4-1.所在
    150, // AJ: 4-2.持ち出し理由
    200, // AK: 4-3.備考
    150, // AL: 5-1.内容
    120, // AM: 5-2.所在
    200, // AN: 5-3.備考
    
    // 計算フィールド（AO-AP列）
    80,  // AO: 貸出日数
    100, // AP: 要注意フラグ
    
    // 参照データ（AQ-AT列）
    100, // AQ: 拠点名
    80,  // AR: 管轄
    200, // AS: formURL
    200  // AT: QRコードURL
  ];
  
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  // 統合ビューの初期データを設定（QUERY関数を使用）
  setupIntegratedViewFormulas(sheet, 'terminal');
  
  addLog('端末系統合ビューシート作成完了');
  return sheet;
}

/**
 * プリンタ・その他系統合ビューシートを作成
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - スプレッドシート
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 作成されたシート
 */
function createIntegratedViewPrinterOtherSheet(spreadsheet) {
  addLog('プリンタ・その他系統合ビューシート作成開始');
  
  const sheet = spreadsheet.insertSheet(VIEW_SHEET_TYPES.INTEGRATED_PRINTER_OTHER);
  
  // ヘッダー行を設定（設計書通りの47列構造）
  const headers = [
    // マスタシートデータ（A-D列、プリンタ・その他は資産管理番号なし）
    '拠点管理番号',                             // A列
    'カテゴリ',                                 // B列
    '機種名',                                   // C列
    '製造番号',                                 // D列
    
    // 収集シートデータ（E-AO列）
    'タイムスタンプ',                           // E列
    '9999.管理ID',                              // F列
    '0-0.拠点管理番号',                         // G列
    '0-1.担当者',                               // H列
    '0-2.EMシステムズの社員ですか？',           // I列
    '0-3.所属会社',                             // J列
    '0-4.ステータス',                           // K列
    '1-1.顧客名または貸出先',                   // L列
    '1-2.顧客番号',                             // M列
    '1-3.住所',                                 // N列
    '1-4.ユーザー機の預り有無',                 // O列
    '1-5.依頼者',                               // P列
    '1-6.備考',                                 // Q列
    '1-7.預りユーザー機のシリアルNo.(製造番号)', // R列
    '2-1.預り機返却の有無',                     // S列
    '2-2.備考',                                 // T列
    '2-3.修理の必要性',                         // U列
    '2-4.備考',                                 // V列
    '3-0.社内ステータス',                       // W列
    '3-0-1.棚卸フラグ',                         // X列
    '3-0-2.拠点',                               // Y列
    '3-1-1.備考',                               // Z列
    '3-2-1.修理依頼の引継ぎ',                   // AA列
    '3-2-2.症状',                               // AB列
    '3-2-3.備考',                               // AC列
    '4-1.所在',                                 // AD列
    '4-2.修理内容',                             // AE列
    '4-3.備考',                                 // AF列
    '5-1.所在',                                 // AG列
    '5-2.持ち出し理由',                         // AH列
    '5-3.備考',                                 // AI列
    '6-1.所在',                                 // AJ列
    '6-2.依頼者',                               // AK列
    '6-3.備考',                                 // AL列
    '7-1.内容',                                 // AM列
    '7-2.所在',                                 // AN列
    '7-3.備考',                                 // AO列
    
    // 計算フィールド（AP-AQ列）
    '貸出日数',                                 // AP列
    '要注意フラグ',                             // AQ列
    
    // 参照データ（AR-AU列）
    '拠点名',                                   // AR列
    '管轄',                                     // AS列
    'formURL',                                  // AT列
    'QRコードURL'                               // AU列
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行のフォーマット
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#70AD47'); // プリンタ系は緑色
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  
  // 列幅の調整（47列対応）
  const columnWidths = [
    // マスタシートデータ（A-D列）
    150, // A: 拠点管理番号
    100, // B: カテゴリ
    150, // C: 機種名
    120, // D: 製造番号
    
    // 収集シートデータ（E-AO列）
    150, // E: タイムスタンプ
    100, // F: 9999.管理ID
    150, // G: 0-0.拠点管理番号
    100, // H: 0-1.担当者
    120, // I: 0-2.EMシステムズの社員ですか？
    150, // J: 0-3.所属会社
    120, // K: 0-4.ステータス
    150, // L: 1-1.顧客名または貸出先
    100, // M: 1-2.顧客番号
    200, // N: 1-3.住所
    120, // O: 1-4.ユーザー機の預り有無
    100, // P: 1-5.依頼者
    200, // Q: 1-6.備考
    150, // R: 1-7.預りユーザー機のシリアルNo.
    120, // S: 2-1.預り機返却の有無
    200, // T: 2-2.備考
    120, // U: 2-3.修理の必要性
    200, // V: 2-4.備考
    120, // W: 3-0.社内ステータス
    100, // X: 3-0-1.棚卸フラグ
    100, // Y: 3-0-2.拠点
    200, // Z: 3-1-1.備考
    150, // AA: 3-2-1.修理依頼の引継ぎ
    150, // AB: 3-2-2.症状
    200, // AC: 3-2-3.備考
    120, // AD: 4-1.所在
    150, // AE: 4-2.修理内容
    200, // AF: 4-3.備考
    120, // AG: 5-1.所在
    150, // AH: 5-2.持ち出し理由
    200, // AI: 5-3.備考
    120, // AJ: 6-1.所在
    100, // AK: 6-2.依頼者
    200, // AL: 6-3.備考
    150, // AM: 7-1.内容
    120, // AN: 7-2.所在
    200, // AO: 7-3.備考
    
    // 計算フィールド（AP-AQ列）
    80,  // AP: 貸出日数
    100, // AQ: 要注意フラグ
    
    // 参照データ（AR-AU列）
    100, // AR: 拠点名
    80,  // AS: 管轄
    200, // AT: formURL
    200  // AU: QRコードURL
  ];
  
  columnWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  // 統合ビューの初期データを設定（QUERY関数を使用）
  setupIntegratedViewFormulas(sheet, 'printer_other');
  
  addLog('プリンタ・その他系統合ビューシート作成完了');
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
 * 統合ビューシートの初期設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 統合ビューシート
 * @param {string} type - シートタイプ ('terminal' または 'printer_other')
 */
function setupIntegratedViewFormulas(sheet, type) {
  addLog('統合ビュー初期設定開始', { type });
  
  // 統合ビューはQUERY関数を使用せず、フォーム回答のトリガーで動的に更新
  // 初期設定では特に処理は不要
  
  addLog('統合ビュー初期設定完了', { type });
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
      case VIEW_SHEET_TYPES.INTEGRATED_TERMINAL:
        sheet = getIntegratedViewTerminalSheet();
        break;
      case VIEW_SHEET_TYPES.INTEGRATED_PRINTER_OTHER:
        sheet = getIntegratedViewPrinterOtherSheet();
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
      const jurisdictionIndex = getColumnIndex(headers, '管轄');
      if (jurisdictionIndex >= 0 && row[jurisdictionIndex] !== filters.jurisdiction) {
        return false;
      }
    }
    
    // 拠点フィルター
    if (filters.location) {
      let locationMatch = false;
      
      // 拠点名での比較
      const locationNameIndex = getColumnIndex(headers, '拠点名');
      if (locationNameIndex >= 0 && row[locationNameIndex] === filters.location) {
        locationMatch = true;
      }
      
      // 拠点管理番号の拠点コード部分での比較
      const managementNumberIndex = getColumnIndex(headers, '拠点管理番号');
      if (!locationMatch && managementNumberIndex >= 0) {
        const managementNumber = row[managementNumberIndex] || '';
        const locationCode = managementNumber.split('_')[0];
        
        // 拠点コードと拠点IDを比較（例: 'OSAKA' === 'osaka' を考慮）
        if (locationCode && locationCode.toLowerCase() === filters.location.toLowerCase()) {
          locationMatch = true;
        }
        
        // 拠点名と拠点コードの対応を確認
        if (!locationMatch) {
          const locationMapping = {
            'osaka': ['OSAKA', 'Osaka', '大阪本社', '大阪'],
            'kobe': ['KOBE', 'Kobe', '神戸支社', '神戸'],
            'himeji': ['HIMEJI', 'Himeji', '姫路営業所', '姫路'],
            'kyoto': ['KYOTO', 'Kyoto', '京都支店', '京都'],
            'tokyo': ['TOKYO', 'Tokyo', '東京支店', '東京']
          };
          
          const locationKey = filters.location.toLowerCase();
          if (locationMapping[locationKey]) {
            locationMatch = locationMapping[locationKey].includes(locationCode) || 
                          locationMapping[locationKey].includes(filters.location) || // 元の値でも比較
                          (locationNameIndex >= 0 && locationMapping[locationKey].includes(row[locationNameIndex]));
          }
          
          // さらに、フィルター値自体が拠点コードと一致するかチェック（大文字小文字を区別しない）
          if (!locationMatch && locationCode && 
              (locationCode.toLowerCase() === filters.location.toLowerCase() ||
               locationCode === filters.location)) {
            locationMatch = true;
          }
        }
      }
      
      if (!locationMatch) {
        return false;
      }
    }
    
    // ステータスフィルター
    if (filters.status) {
      const statusIndex = getColumnIndex(headers, '現在ステータス') || getColumnIndex(headers, '0-4.ステータス');
      if (statusIndex >= 0 && row[statusIndex] !== filters.status) {
        return false;
      }
    }
    
    // カテゴリフィルター
    if (filters.category) {
      const categoryIndex = getColumnIndex(headers, 'カテゴリ');
      if (categoryIndex >= 0 && row[categoryIndex] !== filters.category) {
        return false;
      }
    }
    
    // 検索キーワードフィルター（検索インデックス用）
    if (filters.keyword) {
      const searchKeyIndex = getColumnIndex(headers, '検索キー');
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
 * @deprecated 両方の統合ビューを更新するためのupdateAllIntegratedViewsを使用
 * @return {Object} 更新結果
 */
function updateIntegratedView() {
  return updateAllIntegratedViews();
}

/**
 * すべての統合ビューを更新
 * @return {Object} 更新結果
 */
function updateAllIntegratedViews() {
  const startTime = startPerformanceTimer();
  addLog('すべての統合ビュー更新開始');
  
  try {
    // 端末系統合ビューを更新
    const terminalResult = updateIntegratedViewTerminal();
    if (!terminalResult.success) {
      throw new Error('端末系統合ビュー更新失敗: ' + terminalResult.error);
    }
    
    // プリンタ・その他系統合ビューを更新
    const printerOtherResult = updateIntegratedViewPrinterOther();
    if (!printerOtherResult.success) {
      throw new Error('プリンタ・その他系統合ビュー更新失敗: ' + printerOtherResult.error);
    }
    
    endPerformanceTimer(startTime, 'すべての統合ビュー更新');
    
    return {
      success: true,
      terminal: terminalResult,
      printerOther: printerOtherResult,
      totalRowsUpdated: terminalResult.rowsUpdated + printerOtherResult.rowsUpdated
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'すべての統合ビュー更新エラー');
    addLog('すべての統合ビュー更新エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 端末系統合ビューのデータを更新
 * @return {Object} 更新結果
 */
function updateIntegratedViewTerminal() {
  const startTime = startPerformanceTimer();
  addLog('端末系統合ビュー更新開始');
  
  try {
    const sheet = getIntegratedViewTerminalSheet();
    
    // 端末マスタシートからデータを収集
    const terminalData = getTerminalMasterData();
    
    // ステータス収集シートからデータを取得
    const statusData = getLatestStatusCollectionData();
    
    // 拠点マスタデータも取得
    const locationMasterData = getLocationMasterData();
    
    // データを統合
    const integratedData = integrateDeviceDataForView(terminalData, statusData, locationMasterData, 'terminal');
    
    // シートに書き込み
    if (integratedData.length > 0) {
      sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getLastColumn()).clearContent();
      sheet.getRange(2, 1, integratedData.length, integratedData[0].length).setValues(integratedData);
    }
    
    endPerformanceTimer(startTime, '端末系統合ビュー更新');
    
    return {
      success: true,
      rowsUpdated: integratedData.length
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '端末系統合ビュー更新エラー');
    addLog('端末系統合ビュー更新エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * プリンタ・その他系統合ビューのデータを更新
 * @return {Object} 更新結果
 */
function updateIntegratedViewPrinterOther() {
  const startTime = startPerformanceTimer();
  addLog('プリンタ・その他系統合ビュー更新開始');
  
  try {
    const sheet = getIntegratedViewPrinterOtherSheet();
    
    // プリンタマスタとその他マスタからデータを収集
    const printerData = getPrinterMasterData();
    const otherData = getOtherMasterData();
    
    // ステータス収集シートからデータを取得
    const statusData = getLatestStatusCollectionData();
    
    // 拠点マスタデータも取得
    const locationMasterData = getLocationMasterData();
    
    // データを統合
    const printerIntegratedData = integrateDeviceDataForView(printerData, statusData, locationMasterData, 'printer');
    const otherIntegratedData = integrateDeviceDataForView(otherData, statusData, locationMasterData, 'other');
    const integratedData = [...printerIntegratedData, ...otherIntegratedData];
    
    // シートに書き込み
    if (integratedData.length > 0) {
      sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getLastColumn()).clearContent();
      sheet.getRange(2, 1, integratedData.length, integratedData[0].length).setValues(integratedData);
    }
    
    endPerformanceTimer(startTime, 'プリンタ・その他系統合ビュー更新');
    
    return {
      success: true,
      rowsUpdated: integratedData.length
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'プリンタ・その他系統合ビュー更新エラー');
    addLog('プリンタ・その他系統合ビュー更新エラー', { error: error.toString() });
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
 * 拠点マスタデータを取得
 * @return {Object} 拠点マスタデータ（拠点コードをキーとしたマップ）
 */
function getLocationMasterData() {
  try {
    const data = getMasterData('拠点マスタ');
    const locationMap = {};
    
    data.forEach(location => {
      const locationCode = location['拠点コード'];
      if (locationCode) {
        locationMap[locationCode] = {
          locationName: location['拠点名'] || '',
          jurisdiction: location['管轄'] || ''
        };
      }
    });
    
    return locationMap;
  } catch (error) {
    console.error('拠点マスタデータ取得エラー:', error);
    return {};
  }
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
    const statusSheets = ['端末ステータス収集', 'プリンタその他ステータス収集'];
    const statusMap = {};
    
    for (const sheetName of statusSheets) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() <= 1) continue;
      
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // 拠点管理番号の列を特定
      const managementNumberIndex = getColumnIndex(headers, '0-0.拠点管理番号');
      if (managementNumberIndex < 0) continue;
      
      // 各行を処理
      data.forEach(row => {
        const managementNumber = row[managementNumberIndex];
        if (!managementNumber) return;
        
        // 既存のデータより新しい場合のみ更新
        const timestamp = getValueByColumnName(row, headers, 'タイムスタンプ');
        if (!statusMap[managementNumber] || 
            (timestamp && statusMap[managementNumber]['タイムスタンプ'] < timestamp)) {
          // 行データをオブジェクトに変換（ヘッダーをキーとして使用）
          const statusRecord = rowToObject(row, headers);
          
          statusMap[managementNumber] = statusRecord;
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
    
    // 必要な列のインデックスを動的に取得
    const columnIndices = {
      jurisdiction: getColumnIndex(headers, '管轄'),
      locationName: getColumnIndex(headers, '拠点名'),
      category: getColumnIndex(headers, 'カテゴリ'),
      status: getColumnIndex(headers, '現在ステータス'),
      internalStatus: getColumnIndex(headers, '社内ステータス'),
      cautionFlag: getColumnIndex(headers, '要注意フラグ')
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
 * 指定された拠点管理番号で統合ビューデータを取得
 * @param {Array<string>} managementNumbers - 拠点管理番号の配列
 * @return {Object} データ取得結果
 */
function getIntegratedViewByIds(managementNumbers) {
  const startTime = startPerformanceTimer();
  addLog('ID指定統合ビューデータ取得開始', { count: managementNumbers.length });
  
  try {
    const sheet = getIntegratedViewSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: true,
        data: [sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]], // ヘッダーのみ
        metadata: {
          totalRows: 0,
          filteredRows: 0
        }
      };
    }
    
    const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // 拠点管理番号の列インデックスを動的に取得
    const managementNumberIndex = getColumnIndex(headers, '拠点管理番号');
    if (managementNumberIndex < 0) {
      throw new Error('拠点管理番号列が見つかりません');
    }
    
    // 指定されたIDに一致する行のみを抽出
    const filteredRows = rows.filter(row => 
      managementNumbers.includes(row[managementNumberIndex])
    );
    
    // ヘッダーと抽出されたデータを結合
    const resultData = [headers, ...filteredRows];
    
    endPerformanceTimer(startTime, 'ID指定統合ビューデータ取得');
    
    return {
      success: true,
      data: resultData,
      metadata: {
        totalRows: rows.length,
        filteredRows: filteredRows.length
      }
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'ID指定統合ビューデータ取得エラー');
    addLog('ID指定統合ビューデータ取得エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * デバイスデータとステータスデータを統合（新列構造対応）
 * @param {Array} deviceData - デバイスマスタデータ
 * @param {Object} statusData - ステータス収集データ
 * @param {Object} locationMap - 拠点マスタマップ（現在は使用しない）
 * @param {string} deviceType - デバイスタイプ（terminal/printer/other）
 * @return {Array} 統合されたデータ
 */
function integrateDeviceDataForView(deviceData, statusData, locationMap, deviceType) {
  const integratedRows = [];
  
  for (const device of deviceData) {
    const managementNumber = device['拠点管理番号'];
    if (!managementNumber) continue;
    
    // ステータスデータから最新情報を取得
    const latestStatus = statusData[managementNumber] || {};
    
    // 貸出日数を計算
    let loanDays = 0;
    if (latestStatus && latestStatus['0-4.ステータス'] === '1.貸出中' && latestStatus['タイムスタンプ']) {
      const loanDate = new Date(latestStatus['タイムスタンプ']);
      const today = new Date();
      loanDays = Math.floor((today - loanDate) / (1000 * 60 * 60 * 24));
    }
    
    // 要注意フラグを判定
    const cautionFlag = loanDays >= 90 || (latestStatus && latestStatus['3-0.社内ステータス'] === '1.修理中');
    
    if (deviceType === 'terminal') {
      // 端末系統合ビュー（46列）
      integratedRows.push([
        // マスタシートデータ（A-G列）
        managementNumber,                              // A: 拠点管理番号
        device['カテゴリ'] || '',                     // B: カテゴリ
        device['機種名'] || '',                       // C: 機種名
        device['製造番号'] || '',                     // D: 製造番号
        device['資産管理番号'] || '',                 // E: 資産管理番号
        device['ソフトウェア'] || '',                 // F: ソフトウェア
        device['OS'] || '',                           // G: OS
        
        // 収集シートデータ（H-AN列）- ステータスデータの列名をそのまま使用
        latestStatus['タイムスタンプ'] || '',                           // H: タイムスタンプ
        latestStatus['9999.管理ID'] || '',                              // I: 9999.管理ID
        latestStatus['0-0.拠点管理番号'] || '',                         // J: 0-0.拠点管理番号
        latestStatus['0-1.担当者'] || '',                               // K: 0-1.担当者
        latestStatus['0-2.EMシステムズの社員ですか？'] || '',           // L: 0-2.EMシステムズの社員ですか？
        latestStatus['0-3.所属会社'] || '',                             // M: 0-3.所属会社
        latestStatus['0-4.ステータス'] || '',                           // N: 0-4.ステータス
        latestStatus['1-1.顧客名または貸出先'] || '',                   // O: 1-1.顧客名または貸出先
        latestStatus['1-2.顧客番号'] || '',                             // P: 1-2.顧客番号
        latestStatus['1-3.住所'] || '',                                 // Q: 1-3.住所
        latestStatus['1-4.ユーザー機の預り有無'] || '',                 // R: 1-4.ユーザー機の預り有無
        latestStatus['1-5.依頼者'] || '',                               // S: 1-5.依頼者
        latestStatus['1-6.備考'] || '',                                 // T: 1-6.備考
        latestStatus['1-7.預りユーザー機のシリアルNo.(製造番号)'] || '', // U: 1-7.預りユーザー機のシリアルNo.
        latestStatus['1-8.お預かり証No.'] || '',                        // V: 1-8.お預かり証No.
        latestStatus['2-1.預り機返却の有無'] || '',                     // W: 2-1.預り機返却の有無
        latestStatus['2-2.依頼者'] || '',                               // X: 2-2.依頼者
        latestStatus['2-3.備考'] || '',                                 // Y: 2-3.備考
        latestStatus['3-0.社内ステータス'] || '',                       // Z: 3-0.社内ステータス
        latestStatus['3-0-1.棚卸しフラグ'] || '',                       // AA: 3-0-1.棚卸しフラグ
        latestStatus['3-0-2.拠点'] || '',                               // AB: 3-0-2.拠点（収集データをそのまま使用）
        latestStatus['3-1-1.ソフト'] || '',                             // AC: 3-1-1.ソフト
        latestStatus['3-1-2.備考'] || '',                               // AD: 3-1-2.備考
        latestStatus['3-2-1.端末初期化の引継ぎ'] || '',                 // AE: 3-2-1.端末初期化の引継ぎ
        latestStatus['3-2-2.備考'] || '',                               // AF: 3-2-2.備考
        latestStatus['3-2-3.引継ぎ担当者'] || '',                       // AG: 3-2-3.引継ぎ担当者
        latestStatus['3-2-4.初期化作業の引継ぎ'] || '',                 // AH: 3-2-4.初期化作業の引継ぎ
        latestStatus['4-1.所在'] || '',                                 // AI: 4-1.所在
        latestStatus['4-2.持ち出し理由'] || '',                         // AJ: 4-2.持ち出し理由
        latestStatus['4-3.備考'] || '',                                 // AK: 4-3.備考
        latestStatus['5-1.内容'] || '',                                 // AL: 5-1.内容
        latestStatus['5-2.所在'] || '',                                 // AM: 5-2.所在
        latestStatus['5-3.備考'] || '',                                 // AN: 5-3.備考
        
        // 計算フィールド（AO-AP列）
        loanDays,                                      // AO: 貸出日数
        cautionFlag,                                   // AP: 要注意フラグ
        
        // 参照データ（AQ-AT列）
        latestStatus['拠点名'] || '',                        // AQ: 拠点名（収集データから取得）
        latestStatus['管轄'] || '',                          // AR: 管轄（収集データから取得）
        device['formURL'] || '',                            // AS: formURL
        device['QRコードURL'] || ''                         // AT: QRコードURL
      ]);
    } else {
      // プリンタ・その他系統合ビュー（47列）
      integratedRows.push([
        // マスタシートデータ（A-D列）
        managementNumber,                              // A: 拠点管理番号
        device['カテゴリ'] || '',                     // B: カテゴリ
        device['機種名'] || '',                       // C: 機種名
        device['製造番号'] || '',                     // D: 製造番号
        
        // 収集シートデータ（E-AO列）- ステータスデータの列名をそのまま使用
        latestStatus['タイムスタンプ'] || '',                           // E: タイムスタンプ
        latestStatus['9999.管理ID'] || '',                              // F: 9999.管理ID
        latestStatus['0-0.拠点管理番号'] || '',                         // G: 0-0.拠点管理番号
        latestStatus['0-1.担当者'] || '',                               // H: 0-1.担当者
        latestStatus['0-2.EMシステムズの社員ですか？'] || '',           // I: 0-2.EMシステムズの社員ですか？
        latestStatus['0-3.所属会社'] || '',                             // J: 0-3.所属会社
        latestStatus['0-4.ステータス'] || '',                           // K: 0-4.ステータス
        latestStatus['1-1.顧客名または貸出先'] || '',                   // L: 1-1.顧客名または貸出先
        latestStatus['1-2.顧客番号'] || '',                             // M: 1-2.顧客番号
        latestStatus['1-3.住所'] || '',                                 // N: 1-3.住所
        latestStatus['1-4.ユーザー機の預り有無'] || '',                 // O: 1-4.ユーザー機の預り有無
        latestStatus['1-5.依頼者'] || '',                               // P: 1-5.依頼者
        latestStatus['1-6.備考'] || '',                                 // Q: 1-6.備考
        latestStatus['1-7.預りユーザー機のシリアルNo.(製造番号)'] || '', // R: 1-7.預りユーザー機のシリアルNo.
        latestStatus['2-1.預り機返却の有無'] || '',                     // S: 2-1.預り機返却の有無
        latestStatus['2-2.備考'] || '',                                 // T: 2-2.備考
        latestStatus['2-3.修理の必要性'] || '',                         // U: 2-3.修理の必要性
        latestStatus['2-4.備考'] || '',                                 // V: 2-4.備考
        latestStatus['3-0.社内ステータス'] || '',                       // W: 3-0.社内ステータス
        latestStatus['3-0-1.棚卸しフラグ'] || '',                       // X: 3-0-1.棚卸しフラグ
        latestStatus['3-0-2.拠点'] || '',                               // Y: 3-0-2.拠点（収集データをそのまま使用）
        latestStatus['3-1-1.備考'] || '',                               // Z: 3-1-1.備考
        latestStatus['3-2-1.修理依頼の引継ぎ'] || '',                   // AA: 3-2-1.修理依頼の引継ぎ
        latestStatus['3-2-2.症状'] || '',                               // AB: 3-2-2.症状
        latestStatus['3-2-3.備考'] || '',                               // AC: 3-2-3.備考
        latestStatus['4-1.所在'] || '',                                 // AD: 4-1.所在
        latestStatus['4-2.修理内容'] || '',                             // AE: 4-2.修理内容
        latestStatus['4-3.備考'] || '',                                 // AF: 4-3.備考
        latestStatus['5-1.所在'] || '',                                 // AG: 5-1.所在
        latestStatus['5-2.持ち出し理由'] || '',                         // AH: 5-2.持ち出し理由
        latestStatus['5-3.備考'] || '',                                 // AI: 5-3.備考
        latestStatus['6-1.所在'] || '',                                 // AJ: 6-1.所在
        latestStatus['6-2.依頼者'] || '',                               // AK: 6-2.依頼者
        latestStatus['6-3.備考'] || '',                                 // AL: 6-3.備考
        latestStatus['7-1.内容'] || '',                                 // AM: 7-1.内容
        latestStatus['7-2.所在'] || '',                                 // AN: 7-2.所在
        latestStatus['7-3.備考'] || '',                                 // AO: 7-3.備考
        
        // 計算フィールド（AP-AQ列）
        loanDays,                                      // AP: 貸出日数
        cautionFlag,                                   // AQ: 要注意フラグ
        
        // 参照データ（AR-AU列）
        latestStatus['拠点名'] || '',                        // AR: 拠点名（収集データから取得）
        latestStatus['管轄'] || '',                          // AS: 管轄（収集データから取得）
        device['formURL'] || '',                            // AT: formURL
        device['QRコードURL'] || ''                         // AU: QRコードURL
      ]);
    }
  }
  
  return integratedRows;
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

/**
 * 端末系統合ビューシートの診断関数
 */
function diagnoseTerminalIntegratedViewSheet() {
  console.log('=== 端末系統合ビューシート診断開始 ===');
  
  try {
    // シートの存在確認
    const sheet = getIntegratedViewTerminalSheet();
    console.log('1. シート取得成功:', sheet.getName());
    
    // 基本情報
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    console.log('2. シート情報:');
    console.log('   - 最終行:', lastRow);
    console.log('   - 最終列:', lastColumn);
    
    // ヘッダー確認
    if (lastRow >= 1) {
      const headers = sheet.getRange(1, 1, 1, Math.min(lastColumn, 10)).getValues()[0];
      console.log('3. ヘッダー（最初の10列）:', headers);
    } else {
      console.log('3. シートが完全に空です');
    }
    
    // データ行の確認
    if (lastRow > 1) {
      console.log('4. データ行数:', lastRow - 1);
      const firstDataRow = sheet.getRange(2, 1, 1, Math.min(lastColumn, 5)).getValues()[0];
      console.log('   最初のデータ行（最初の5列）:', firstDataRow);
    } else {
      console.log('4. データ行がありません（ヘッダーのみ）');
    }
    
    // getIntegratedViewData関数のテスト
    console.log('5. getIntegratedViewData関数テスト:');
    const result = getIntegratedViewData('', 'terminal');
    console.log('   - success:', result.success);
    console.log('   - データ行数:', result.data ? result.data.length : 'null');
    console.log('   - メタデータ:', result.metadata);
    if (result.error) {
      console.log('   - エラー:', result.error);
    }
    
    // 端末マスタデータの確認
    console.log('6. 端末マスタデータ確認:');
    const terminalData = getTerminalMasterData();
    console.log('   - 端末マスタデータ件数:', terminalData.length);
    
    // ステータス収集データの確認
    console.log('7. ステータス収集データ確認:');
    const statusData = getLatestStatusCollectionData();
    console.log('   - ステータスデータ件数:', Object.keys(statusData).length);
    
  } catch (error) {
    console.error('診断中にエラーが発生しました:', error);
    console.error('スタックトレース:', error.stack);
  }
  
  console.log('=== 端末系統合ビューシート診断完了 ===');
}

/**
 * プリンタ・その他系統合ビューシートの診断関数
 */
function diagnosePrinterOtherIntegratedViewSheet() {
  console.log('=== プリンタ・その他系統合ビューシート診断開始 ===');
  
  try {
    // シートの存在確認
    const sheet = getIntegratedViewPrinterOtherSheet();
    console.log('1. シート取得成功:', sheet.getName());
    
    // 基本情報
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    console.log('2. シート情報:');
    console.log('   - 最終行:', lastRow);
    console.log('   - 最終列:', lastColumn);
    
    // ヘッダー確認
    if (lastRow >= 1) {
      const headers = sheet.getRange(1, 1, 1, Math.min(lastColumn, 10)).getValues()[0];
      console.log('3. ヘッダー（最初の10列）:', headers);
    } else {
      console.log('3. シートが完全に空です');
    }
    
    // データ行の確認
    if (lastRow > 1) {
      console.log('4. データ行数:', lastRow - 1);
      const firstDataRow = sheet.getRange(2, 1, 1, Math.min(lastColumn, 5)).getValues()[0];
      console.log('   最初のデータ行（最初の5列）:', firstDataRow);
    } else {
      console.log('4. データ行がありません（ヘッダーのみ）');
    }
    
    // getIntegratedViewData関数のテスト
    console.log('5. getIntegratedViewData関数テスト:');
    const result = getIntegratedViewData('', 'printer_other');
    console.log('   - success:', result.success);
    console.log('   - データ行数:', result.data ? result.data.length : 'null');
    console.log('   - メタデータ:', result.metadata);
    if (result.error) {
      console.log('   - エラー:', result.error);
    }
    
  } catch (error) {
    console.error('診断中にエラーが発生しました:', error);
    console.error('スタックトレース:', error.stack);
  }
  
  console.log('=== プリンタ・その他系統合ビューシート診断完了 ===');
}

/**
 * 端末系統合ビューシートを手動で更新（テスト用）
 */
function manuallyUpdateTerminalIntegratedView() {
  console.log('=== 端末系統合ビューシート手動更新開始 ===');
  
  try {
    const result = updateIntegratedViewTerminal();
    console.log('更新結果:', result);
    
    if (result.success) {
      console.log('更新成功 - 更新行数:', result.rowsUpdated);
      
      // 更新後のシート状態を確認
      const sheet = getIntegratedViewTerminalSheet();
      console.log('更新後のシート情報:');
      console.log('- 最終行:', sheet.getLastRow());
      console.log('- 最終列:', sheet.getLastColumn());
      
      if (sheet.getLastRow() > 1) {
        const firstDataRow = sheet.getRange(2, 1, 1, Math.min(sheet.getLastColumn(), 5)).getValues()[0];
        console.log('- 最初のデータ行（最初の5列）:', firstDataRow);
      }
    } else {
      console.error('更新失敗:', result.error);
    }
    
  } catch (error) {
    console.error('手動更新中にエラーが発生しました:', error);
    console.error('スタックトレース:', error.stack);
  }
  
  console.log('=== 端末系統合ビューシート手動更新完了 ===');
}

/**
 * 統合ビューデータを取得（シンプル版）
 * @param {string} location - 拠点名（空文字の場合は全拠点）
 * @param {string} viewType - ビュータイプ（terminal/printer_other）
 * @return {Object} レスポンスオブジェクト
 */
/**
 * 拠点コードを正規化（大文字小文字の統一と例外処理）
 * @param {string} locationCode - 拠点コード
 * @return {string} 正規化された拠点コード
 */
function normalizeLocationCode(locationCode) {
  if (!locationCode) return '';
  
  const code = locationCode.toString().toUpperCase();
  
  // 例外処理：Osaka, Kobe, Himejiの場合
  const exceptions = {
    'OSAKA': 'OSAKA',
    'KOBE': 'KOBE',
    'HIMEJI': 'HIMEJI'
  };
  
  // 既に正しい形式の場合はそのまま返す
  if (exceptions[code]) {
    return code;
  }
  
  // その他の一般的な形式も許容
  return code;
}

/**
 * 統合ビューから利用可能な拠点リストを取得
 * @return {Array} 拠点コードの配列
 */
function getAvailableLocationsFromIntegratedView() {
  try {
    // 端末とプリンタ両方のシートから拠点を収集
    const terminalSheet = getIntegratedViewTerminalSheet();
    const printerSheet = getIntegratedViewPrinterOtherSheet();
    
    const locations = new Set();
    
    // 端末シートから拠点を収集
    if (terminalSheet && terminalSheet.getLastRow() > 1) {
      const terminalData = terminalSheet.getRange(2, 1, terminalSheet.getLastRow() - 1, 1).getValues();
      terminalData.forEach(row => {
        const managementNumber = row[0];
        if (managementNumber) {
          const locationCode = managementNumber.toString().split('_')[0];
          if (locationCode) {
            locations.add(locationCode);
          }
        }
      });
    }
    
    // プリンタシートから拠点を収集
    if (printerSheet && printerSheet.getLastRow() > 1) {
      const printerData = printerSheet.getRange(2, 1, printerSheet.getLastRow() - 1, 1).getValues();
      printerData.forEach(row => {
        const managementNumber = row[0];
        if (managementNumber) {
          const locationCode = managementNumber.toString().split('_')[0];
          if (locationCode) {
            locations.add(locationCode);
          }
        }
      });
    }
    
    // 配列に変換してソート
    return Array.from(locations).sort();
    
  } catch (error) {
    console.error('拠点リスト取得エラー:', error);
    return [];
  }
}

function getIntegratedViewData(location, viewType) {
  const startTime = startPerformanceTimer();
  addLog('統合ビューデータ取得開始', { location, viewType });
  
  try {
    // ビューシートを取得
    let sheet;
    if (viewType === 'terminal') {
      sheet = getIntegratedViewTerminalSheet();
      addLog('端末系統合ビューシート取得', { sheetName: sheet ? sheet.getName() : 'null' });
    } else if (viewType === 'printer_other') {
      sheet = getIntegratedViewPrinterOtherSheet();
      addLog('プリンタ・その他系統合ビューシート取得', { sheetName: sheet ? sheet.getName() : 'null' });
    } else {
      throw new Error('不正なビュータイプ: ' + viewType);
    }
    
    if (!sheet) {
      throw new Error('統合ビューシートが見つかりません');
    }
    
    // 全データを取得
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    addLog('シート情報', { 
      sheetName: sheet.getName(),
      lastRow: lastRow,
      lastColumn: lastColumn 
    });
    
    if (lastRow === 0) {
      addLog('シートが空です（行がありません）');
      return {
        success: true,
        data: [],
        metadata: {
          viewType: viewType,
          location: location,
          totalRows: 0
        }
      };
    }
    
    if (lastRow === 1) {
      // ヘッダーのみの場合
      addLog('シートにはヘッダーのみがあります');
      const headers = sheet.getRange(1, 1, 1, lastColumn).getValues();
      return {
        success: true,
        data: headers,
        metadata: {
          viewType: viewType,
          location: location,
          totalRows: 0
        }
      };
    }
    
    const allData = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    addLog('全データ取得完了', { 
      totalRows: allData.length,
      headerRow: allData[0] ? allData[0].slice(0, 5) : 'no header' // 最初の5列のみログに記録
    });
    
    // 拠点でフィルタリング
    let filteredData = allData;
    if (location) {
      // ヘッダー行を保持
      filteredData = [allData[0]];
      
      // 拠点管理番号列のインデックスを検索（A列 = 0）
      const managementNumberIndex = 0;
      
      // 拠点コードの正規化（例外処理）
      const normalizedLocationCode = normalizeLocationCode(location);
      addLog('拠点フィルタリング開始', { 
        location: location,
        normalizedLocationCode: normalizedLocationCode 
      });
      
      // データ行をフィルタリング
      for (let i = 1; i < allData.length; i++) {
        const managementNumber = allData[i][managementNumberIndex];
        if (managementNumber) {
          // 拠点管理番号から拠点コードを抽出（最初の_まで）
          const dataLocationCode = managementNumber.toString().split('_')[0];
          
          if (dataLocationCode) {
            // 拠点コードも正規化して比較
            const normalizedDataCode = normalizeLocationCode(dataLocationCode);
            
            if (normalizedDataCode === normalizedLocationCode) {
              filteredData.push(allData[i]);
            }
          }
        }
      }
      
      addLog('拠点フィルタリング完了', { 
        originalRows: allData.length - 1,
        filteredRows: filteredData.length - 1
      });
    }
    
    endPerformanceTimer(startTime, '統合ビューデータ取得完了');
    
    const result = {
      success: true,
      data: filteredData,
      metadata: {
        viewType: viewType,
        location: location || '全拠点',
        totalRows: filteredData.length - 1 // ヘッダー行を除く
      }
    };
    
    addLog('返却データ', {
      success: result.success,
      dataRows: result.data ? result.data.length : 0,
      metadata: result.metadata
    });
    
    return result;
    
  } catch (error) {
    endPerformanceTimer(startTime, '統合ビューデータ取得エラー');
    addLog('統合ビューデータ取得エラー', { 
      error: error.toString(),
      stack: error.stack || 'no stack trace'
    });
    
    return {
      success: false,
      error: error.toString(),
      data: null
    };
  }
}