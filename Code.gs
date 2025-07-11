// 代替機管理システム
// メインコード

// グローバル設定
const CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'),
  DEBUG_MODE: PropertiesService.getScriptProperties().getProperty('DEBUG_MODE') === 'true',
  STATUS_CHANGE_NOTIFICATION_ENABLED: PropertiesService.getScriptProperties().getProperty('STATUS_CHANGE_NOTIFICATION_ENABLED') === 'true'
};

// シート名の定義
const SHEET_NAMES = {
  LOCATION_MASTER: '拠点マスタ',
  MODEL_MASTER: '機種マスタ',
  DATA_TYPE_MASTER: 'データタイプマスタ',
  TERMINAL_MASTER: '端末マスタ',
  PRINTER_MASTER: 'プリンタマスタ',
  OTHER_MASTER: 'そのたマスタ',
  TERMINAL_STATUS: '端末ステータス収集',
  PRINTER_STATUS: 'プリンタステータス収集',
  OTHER_STATUS: 'そのたステータス収集'
};

// エントリーポイント
function doGet(e) {
  // パラメータによるページ分岐（将来の拡張用）
  const page = e.parameter.page || 'index';
  
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('代替機管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    logError('doGet', error);
    return HtmlService.createHtmlOutput('エラーが発生しました。');
  }
}

// ユーザー情報の取得
function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    logError('getUserEmail', error);
    return null;
  }
}

// エラーログ
function logError(functionName, error) {
  console.error(`[${functionName}] ${error.toString()}`);
  
  if (CONFIG.DEBUG_MODE) {
    // デバッグモードではコンソールに詳細を出力
    console.error(error.stack);
  }
  
  // エラー通知（設定されている場合）
  const errorEmail = PropertiesService.getScriptProperties().getProperty('ERROR_NOTIFICATION_EMAIL');
  if (errorEmail && CONFIG.STATUS_CHANGE_NOTIFICATION_ENABLED) {
    try {
      MailApp.sendEmail({
        to: errorEmail,
        subject: `[代替機管理システム] エラー通知 - ${functionName}`,
        body: `以下のエラーが発生しました:\n\n${error.toString()}\n\nスタック:\n${error.stack}\n\n発生時刻: ${new Date().toLocaleString('ja-JP')}`
      });
    } catch (mailError) {
      console.error('エラー通知の送信に失敗しました:', mailError);
    }
  }
}

// スプレッドシートの取得
function getSpreadsheet() {
  if (!CONFIG.SPREADSHEET_ID) {
    throw new Error('SPREADSHEET_IDが設定されていません。');
  }
  
  try {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (error) {
    logError('getSpreadsheet', error);
    throw new Error('スプレッドシートへのアクセスに失敗しました。');
  }
}

// シートの取得
function getSheet(sheetName) {
  const spreadsheet = getSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。`);
  }
  
  return sheet;
}

// HTMLファイルの取得
function getSpreadsheetViewerHtml() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('spreadsheet-viewer').getContent();
    const js = HtmlService.createHtmlOutputFromFile('spreadsheet-viewer-functions').getContent();
    return html + js;
  } catch (error) {
    logError('getSpreadsheetViewerHtml', error);
    return '<div class="alert alert-danger">スプレッドシートビューアーの読み込みに失敗しました。</div>';
  }
}

function getUrlGeneratorHtml() {
  try {
    return HtmlService.createHtmlOutputFromFile('url-generator').getContent();
  } catch (error) {
    logError('getUrlGeneratorHtml', error);
    return '<div class="alert alert-danger">URL生成画面の読み込みに失敗しました。</div>';
  }
}

function getModelMasterHtml() {
  try {
    return HtmlService.createHtmlOutputFromFile('model-master').getContent();
  } catch (error) {
    logError('getModelMasterHtml', error);
    return '<div class="alert alert-danger">機種マスタ画面の読み込みに失敗しました。</div>';
  }
}

function getLocationMasterHtml() {
  try {
    return HtmlService.createHtmlOutputFromFile('location-master').getContent();
  } catch (error) {
    logError('getLocationMasterHtml', error);
    return '<div class="alert alert-danger">拠点マスタ画面の読み込みに失敗しました。</div>';
  }
}

function getSettingsHtml() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('settings').getContent();
    const js = HtmlService.createHtmlOutputFromFile('settings-functions').getContent();
    return html + js;
  } catch (error) {
    logError('getSettingsHtml', error);
    return '<div class="alert alert-danger">設定画面の読み込みに失敗しました。</div>';
  }
}

// 拠点リストの取得
function getLocationList() {
  try {
    const sheet = getSheet(SHEET_NAMES.LOCATION_MASTER);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return []; // ヘッダーのみの場合
    
    const headers = data[0];
    const locations = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // 拠点IDが空の場合はスキップ
      
      locations.push({
        id: row[0],
        code: row[1],
        name: row[2],
        region: row[3] || '',
        email: row[4] || '',
        notificationEnabled: row[5] === true || row[5] === 'TRUE',
        status: row[6] === true || row[6] === 'TRUE',
        createdAt: row[7],
        updatedAt: row[8]
      });
    }
    
    // 有効な拠点のみを返す
    return locations.filter(loc => loc.status);
  } catch (error) {
    logError('getLocationList', error);
    return [];
  }
}

// データタイプリストの取得
function getDataTypeList() {
  try {
    const sheet = getSheet(SHEET_NAMES.DATA_TYPE_MASTER);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    const dataTypes = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      dataTypes.push({
        id: row[0],
        name: row[1],
        description: row[2] || '',
        displayOrder: row[3] || 999,
        status: row[4] === true || row[4] === 'TRUE',
        createdAt: row[5],
        updatedAt: row[6]
      });
    }
    
    // 有効なデータタイプのみを表示順でソート
    return dataTypes
      .filter(dt => dt.status)
      .sort((a, b) => a.displayOrder - b.displayOrder);
  } catch (error) {
    logError('getDataTypeList', error);
    return [];
  }
}

// スプレッドシートデータの取得
function getSpreadsheetData(locationId, dataTypeId) {
  try {
    // データタイプに応じた処理
    if (dataTypeId === 'SUMMARY') {
      return getSummaryData(locationId);
    }
    
    // 通常データと監査データの取得
    const allData = [];
    const headers = getDataHeaders(dataTypeId);
    
    // 端末マスタからデータ取得
    const terminalData = getDeviceData(SHEET_NAMES.TERMINAL_MASTER, '端末', locationId);
    allData.push(...terminalData);
    
    // プリンタマスタからデータ取得
    const printerData = getDeviceData(SHEET_NAMES.PRINTER_MASTER, 'プリンタ', locationId);
    allData.push(...printerData);
    
    // その他マスタからデータ取得
    const otherData = getDeviceData(SHEET_NAMES.OTHER_MASTER, 'その他', locationId);
    allData.push(...otherData);
    
    // フィルタリング（拠点指定がある場合）
    const filteredData = locationId ? 
      allData.filter(row => row[1] === locationId) : // 拠点IDでフィルタ
      allData;
    
    return {
      headers: headers,
      data: filteredData,
      dataType: dataTypeId,
      location: locationId,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    logError('getSpreadsheetData', error);
    return {
      headers: [],
      data: [],
      error: error.toString()
    };
  }
}

// データヘッダーの取得
function getDataHeaders(dataTypeId) {
  if (dataTypeId === 'AUDIT') {
    return [
      'デバイス種別',
      '拠点ID',
      '拠点管理番号',
      '機種名',
      'メーカー',
      '製造番号',
      'ステータス',
      '資産管理番号',
      '担当者',
      '貸出先',
      'ユーザー機預り有無',
      '預り機ステータス',
      '備考',
      '更新日時'
    ];
  } else {
    // 通常データのヘッダー
    return [
      'デバイス種別',
      '拠点ID',
      '拠点管理番号',
      '機種名',
      'メーカー',
      '製造番号',
      'ステータス',
      '更新日時'
    ];
  }
}

// デバイスデータの取得
function getDeviceData(sheetName, deviceType, locationId) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // 拠点管理番号が空の場合はスキップ
      
      // 拠点IDの抽出（拠点管理番号から）
      const locationCode = row[0].split('_')[0];
      
      // 基本データ
      const deviceData = [
        deviceType,
        locationCode,
        row[0], // 拠点管理番号
        row[1], // 機種名
        row[2], // メーカー
        row[3], // 製造番号
        row[4] || '未設定', // ステータス
        row[10] || new Date() // 更新日時
      ];
      
      result.push(deviceData);
    }
    
    return result;
  } catch (error) {
    logError('getDeviceData', error);
    return [];
  }
}

// サマリーデータの取得
function getSummaryData(locationId) {
  try {
    // サマリーシートからデータを取得
    const summarySheet = getSheet('サマリー');
    if (!summarySheet) {
      return {
        headers: [],
        data: [],
        error: 'サマリーシートが見つかりません'
      };
    }
    
    const data = summarySheet.getDataRange().getValues();
    
    // 拠点でフィルタリング（必要な場合）
    let filteredData = data;
    if (locationId) {
      // 拠点フィルタリングロジック（サマリーデータの構造に応じて実装）
    }
    
    return {
      headers: data[0] || [],
      data: filteredData.slice(1),
      dataType: 'SUMMARY',
      location: locationId,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    logError('getSummaryData', error);
    return {
      headers: [],
      data: [],
      error: error.toString()
    };
  }
}

// 預り機ステータスの更新
function updateDepositStatus(updates) {
  try {
    let updatedCount = 0;
    const errors = [];
    
    updates.forEach(update => {
      try {
        // 対象デバイスの特定
        const deviceType = identifyDeviceType(update.id);
        const sheetName = getSheetNameByDeviceType(deviceType);
        const sheet = getSheet(sheetName);
        
        // 該当行の検索
        const data = sheet.getDataRange().getValues();
        let rowIndex = -1;
        
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === update.id) {
            rowIndex = i + 1; // シートの行番号は1から始まる
            break;
          }
        }
        
        if (rowIndex > 0) {
          // 預り機ステータスの更新（仮の列インデックス）
          const statusColumn = 12; // 実際の列に合わせて調整
          sheet.getRange(rowIndex, statusColumn).setValue(update.newStatus);
          
          // 更新日時の記録
          const updateTimeColumn = 11;
          sheet.getRange(rowIndex, updateTimeColumn).setValue(new Date());
          
          // 変更履歴の記録（別シートまたは備考欄に記録）
          const remarkColumn = 13;
          const currentRemark = sheet.getRange(rowIndex, remarkColumn).getValue();
          const newRemark = `${currentRemark}\n[${new Date().toLocaleString('ja-JP')}] 預り機ステータス変更: ${update.newStatus} (理由: ${update.reason})`;
          sheet.getRange(rowIndex, remarkColumn).setValue(newRemark);
          
          updatedCount++;
        }
      } catch (error) {
        errors.push({
          id: update.id,
          error: error.toString()
        });
      }
    });
    
    return {
      success: errors.length === 0,
      updatedCount: updatedCount,
      errors: errors,
      error: errors.length > 0 ? `${errors.length}件のエラーが発生しました` : null
    };
  } catch (error) {
    logError('updateDepositStatus', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// デバイスタイプの識別
function identifyDeviceType(managementId) {
  // 拠点管理番号からデバイスタイプを推定
  const parts = managementId.split('_');
  if (parts.length < 2) return 'その他';
  
  const category = parts[1];
  if (['SV', 'DT', 'NB', 'TB'].includes(category)) {
    return '端末';
  } else if (['PR', 'MFP'].includes(category)) {
    return 'プリンタ';
  } else {
    return 'その他';
  }
}

// デバイスタイプに応じたシート名の取得
function getSheetNameByDeviceType(deviceType) {
  const sheetMap = {
    '端末': SHEET_NAMES.TERMINAL_MASTER,
    'プリンタ': SHEET_NAMES.PRINTER_MASTER,
    'その他': SHEET_NAMES.OTHER_MASTER
  };
  return sheetMap[deviceType] || SHEET_NAMES.OTHER_MASTER;
}

// データのエクスポート
function exportSpreadsheetData(locationId, dataTypeId, deviceType, searchValue) {
  try {
    // データの取得
    const result = getSpreadsheetData(locationId, dataTypeId);
    
    if (!result || !result.data || result.data.length === 0) {
      return {
        success: false,
        error: 'エクスポートするデータがありません'
      };
    }
    
    // フィルタリング
    let filteredData = result.data;
    
    if (deviceType) {
      const deviceTypeIndex = result.headers.indexOf('デバイス種別');
      if (deviceTypeIndex >= 0) {
        filteredData = filteredData.filter(row => row[deviceTypeIndex] === deviceType);
      }
    }
    
    if (searchValue) {
      filteredData = filteredData.filter(row => 
        row.some(cell => cell && cell.toString().toLowerCase().includes(searchValue.toLowerCase()))
      );
    }
    
    // CSV形式に変換
    const csv = convertToCSV(result.headers, filteredData);
    
    // ファイル名の生成
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    const filename = `device_data_export_${timestamp}.csv`;
    
    return {
      success: true,
      data: csv,
      filename: filename
    };
  } catch (error) {
    logError('exportSpreadsheetData', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// CSV変換
function convertToCSV(headers, data) {
  const csvRows = [];
  
  // ヘッダー行
  csvRows.push(headers.map(h => `"${h}"`).join(','));
  
  // データ行
  data.forEach(row => {
    const csvRow = row.map(cell => {
      if (cell === null || cell === undefined) return '""';
      const str = cell.toString().replace(/"/g, '""');
      return `"${str}"`;
    }).join(',');
    csvRows.push(csvRow);
  });
  
  // BOM付きUTF-8で返す（Excelでの文字化け対策）
  return '\uFEFF' + csvRows.join('\n');
}

// システム設定の取得
function getSystemSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    
    return {
      debugMode: props.getProperty('DEBUG_MODE') === 'true',
      terminalFormUrl: props.getProperty('TERMINAL_FORM_URL') || '',
      printerFormUrl: props.getProperty('PRINTER_FORM_URL') || '',
      qrPageUrl: props.getProperty('QR_PAGE_URL') || '',
      errorNotificationEmail: props.getProperty('ERROR_NOTIFICATION_EMAIL') || '',
      alertNotificationEmail: props.getProperty('ALERT_NOTIFICATION_EMAIL') || '',
      statusChangeNotificationEnabled: props.getProperty('STATUS_CHANGE_NOTIFICATION_ENABLED') === 'true',
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      lastUpdate: props.getProperty('LAST_SYSTEM_UPDATE') || null
    };
  } catch (error) {
    logError('getSystemSettings', error);
    return {};
  }
}

// デバッグモードの更新
function updateDebugMode(enabled) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('DEBUG_MODE', enabled.toString());
    props.setProperty('LAST_SYSTEM_UPDATE', new Date().toISOString());
    
    // グローバル変数も更新
    CONFIG.DEBUG_MODE = enabled;
    
    return { success: true };
  } catch (error) {
    logError('updateDebugMode', error);
    return { success: false, error: error.toString() };
  }
}

// フォームURL設定の保存
function saveFormUrlSettings(settings) {
  try {
    // デバッグモードチェック
    if (!CONFIG.DEBUG_MODE) {
      return { success: false, error: 'デバッグモードが無効です' };
    }
    
    const props = PropertiesService.getScriptProperties();
    
    if (settings.terminalFormUrl) {
      props.setProperty('TERMINAL_FORM_URL', settings.terminalFormUrl);
    }
    if (settings.printerFormUrl) {
      props.setProperty('PRINTER_FORM_URL', settings.printerFormUrl);
    }
    if (settings.qrPageUrl) {
      props.setProperty('QR_PAGE_URL', settings.qrPageUrl);
    }
    
    props.setProperty('LAST_SYSTEM_UPDATE', new Date().toISOString());
    
    return { success: true };
  } catch (error) {
    logError('saveFormUrlSettings', error);
    return { success: false, error: error.toString() };
  }
}

// ログ通知設定の保存
function saveLogNotificationSettings(settings) {
  try {
    // デバッグモードチェック
    if (!CONFIG.DEBUG_MODE) {
      return { success: false, error: 'デバッグモードが無効です' };
    }
    
    const props = PropertiesService.getScriptProperties();
    
    if (settings.errorNotificationEmail !== undefined) {
      props.setProperty('ERROR_NOTIFICATION_EMAIL', settings.errorNotificationEmail);
    }
    if (settings.alertNotificationEmail !== undefined) {
      props.setProperty('ALERT_NOTIFICATION_EMAIL', settings.alertNotificationEmail);
    }
    
    props.setProperty('LAST_SYSTEM_UPDATE', new Date().toISOString());
    
    return { success: true };
  } catch (error) {
    logError('saveLogNotificationSettings', error);
    return { success: false, error: error.toString() };
  }
}

// テストメール送信
function sendTestEmail(type, emails) {
  try {
    const subject = type === 'error' ? 
      '[代替機管理システム] エラー通知テスト' : 
      '[代替機管理システム] アラート通知テスト';
    
    const body = `これは${type === 'error' ? 'エラー' : 'アラート'}通知のテストメールです。\n\n` +
                 `送信日時: ${new Date().toLocaleString('ja-JP')}\n` +
                 `送信先: ${emails}\n\n` +
                 `このメールが届いていれば、通知設定は正常に機能しています。`;
    
    // 複数のメールアドレスに対応
    const emailList = emails.split(',').map(e => e.trim()).filter(e => e);
    
    emailList.forEach(email => {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body
      });
    });
    
    return { success: true, sentTo: emailList };
  } catch (error) {
    logError('sendTestEmail', error);
    return { success: false, error: error.toString() };
  }
}

// グローバル通知設定の保存
function saveGlobalNotificationSetting(enabled) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('STATUS_CHANGE_NOTIFICATION_ENABLED', enabled.toString());
    props.setProperty('LAST_SYSTEM_UPDATE', new Date().toISOString());
    
    // グローバル変数も更新
    CONFIG.STATUS_CHANGE_NOTIFICATION_ENABLED = enabled;
    
    return { success: true };
  } catch (error) {
    logError('saveGlobalNotificationSetting', error);
    return { success: false, error: error.toString() };
  }
}

// システム設定の保存
function saveSystemSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    
    // 現在はデバッグモードのみ
    if (settings.debugMode !== undefined) {
      return updateDebugMode(settings.debugMode);
    }
    
    props.setProperty('LAST_SYSTEM_UPDATE', new Date().toISOString());
    
    return { success: true };
  } catch (error) {
    logError('saveSystemSettings', error);
    return { success: false, error: error.toString() };
  }
}

// 初期設定のセットアップ（初回実行時）
function setupInitialSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    
    // 既に設定されている場合はスキップ
    if (props.getProperty('INITIALIZED') === 'true') {
      return;
    }
    
    // デフォルト値の設定
    props.setProperties({
      'DEBUG_MODE': 'false',
      'STATUS_CHANGE_NOTIFICATION_ENABLED': 'true',
      'INITIALIZED': 'true',
      'LAST_SYSTEM_UPDATE': new Date().toISOString()
    });
    
    console.log('初期設定が完了しました。');
  } catch (error) {
    logError('setupInitialSettings', error);
  }
}

// スクリプト実行時に初期設定を確認
setupInitialSettings();