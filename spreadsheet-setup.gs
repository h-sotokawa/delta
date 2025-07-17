/**
 * 代替機管理システム - スプレッドシート初期セットアップ
 * 新しいスプレッドシートで使用する際の初期化スクリプト
 */

/**
 * メイン初期化関数
 * 新しいスプレッドシートで最初に実行する
 */
function initializeSpreadsheet() {
  try {
    console.log('=== スプレッドシート初期化開始 ===');
    
    // 1. PropertiesService初期設定
    setupPropertiesService();
    
    // 2. 必要なシートの作成
    createAllRequiredSheets();
    
    // 3. 統合ビューの設定
    setupIntegratedViews();
    
    // 4. トリガーの設定
    setupTriggers();
    
    console.log('=== スプレッドシート初期化完了 ===');
    return {
      success: true,
      message: 'スプレッドシートの初期化が正常に完了しました'
    };
    
  } catch (error) {
    console.error('初期化エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * PropertiesServiceの初期設定
 */
function setupPropertiesService() {
  console.log('PropertiesService設定開始...');
  
  const properties = PropertiesService.getScriptProperties();
  
  // 既存のSPREADSHEET_ID_DESTINATIONを使用、または新しく設定
  let spreadsheetId = properties.getProperty('SPREADSHEET_ID_DESTINATION');
  
  if (!spreadsheetId) {
    // SPREADSHEET_ID_DESTINATIONが未設定の場合、アクティブスプレッドシートから取得を試行
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (spreadsheet) {
        spreadsheetId = spreadsheet.getId();
        properties.setProperty('SPREADSHEET_ID_DESTINATION', spreadsheetId);
        console.log('アクティブスプレッドシートからIDを設定:', spreadsheetId);
      }
    } catch (error) {
      throw new Error('スプレッドシートIDが設定されておらず、アクティブスプレッドシートも取得できません。SPREADSHEET_ID_DESTINATIONを手動で設定してください。');
    }
  }
  
  // SPREADSHEET_ID_MAINにも同じIDを設定
  properties.setProperty('SPREADSHEET_ID_MAIN', spreadsheetId);
  
  // デフォルト設定値（後で管理画面から変更可能）
  properties.setProperty('TERMINAL_COMMON_FORM_URL', '');
  properties.setProperty('PRINTER_COMMON_FORM_URL', '');
  properties.setProperty('QR_REDIRECT_URL', '');
  properties.setProperty('ERROR_NOTIFICATION_EMAIL', '');
  properties.setProperty('ALERT_NOTIFICATION_EMAIL', '');
  properties.setProperty('DEBUG_MODE', 'false');
  
  console.log('PropertiesService設定完了');
}

/**
 * 必要なシートをすべて作成
 */
function createAllRequiredSheets() {
  console.log('必要なシートの作成開始...');
  
  // PropertiesServiceからスプレッドシートIDを取得
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID_DESTINATION') || properties.getProperty('SPREADSHEET_ID_MAIN');
  
  if (!spreadsheetId) {
    throw new Error('スプレッドシートIDが設定されていません。先にsetupPropertiesService()を実行してください。');
  }
  
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  // 作成するシートの定義
  const sheetsToCreate = [
    {
      name: '拠点マスタ',
      headers: ['拠点ID', '拠点コード', '拠点名', 'グループメールアドレス', 'ステータス', '管轄', 'ステータス変更通知', '作成日時', '更新日時'],
      frozen: { rows: 1, columns: 2 }
    },
    {
      name: '機種マスタ',
      headers: ['機種名', 'カテゴリ', 'メーカー', 'モデル名', '仕様', 'ステータス', '作成日時', '更新日時'],
      frozen: { rows: 1, columns: 2 }
    },
    {
      name: '端末マスタ',
      headers: ['拠点管理番号', '拠点', 'カテゴリ', '機種名', '資産管理番号', '製造番号', '作成日時'],
      frozen: { rows: 1, columns: 1 }
    },
    {
      name: 'プリンタマスタ',
      headers: ['拠点管理番号', '拠点', '機種名', '作成日時'],
      frozen: { rows: 1, columns: 1 }
    },
    {
      name: 'その他マスタ',
      headers: ['拠点管理番号', '拠点', 'カテゴリ', '機種名', '作成日時'],
      frozen: { rows: 1, columns: 1 }
    },
    {
      name: 'integrated_view_terminal',
      headers: generateIntegratedViewHeaders('terminal'),
      frozen: { rows: 1, columns: 3 }
    },
    {
      name: 'integrated_view_printer_other',
      headers: generateIntegratedViewHeaders('printer_other'),
      frozen: { rows: 1, columns: 3 }
    },
    {
      name: 'search_index',
      headers: ['機器ID', '検索キー', 'カテゴリ', '拠点', '状態', '最終更新日時'],
      frozen: { rows: 1, columns: 1 }
    }
  ];
  
  // 既存シートの確認と作成
  sheetsToCreate.forEach(sheetConfig => {
    let sheet = spreadsheet.getSheetByName(sheetConfig.name);
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = spreadsheet.insertSheet(sheetConfig.name);
      console.log(`シート作成: ${sheetConfig.name}`);
    } else {
      console.log(`シート既存: ${sheetConfig.name}`);
    }
    
    // ヘッダー設定
    if (sheetConfig.headers && sheetConfig.headers.length > 0) {
      const headerRange = sheet.getRange(1, 1, 1, sheetConfig.headers.length);
      headerRange.setValues([sheetConfig.headers]);
      headerRange.setBackground('#4285F4');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
    }
    
    // 固定行・列の設定
    if (sheetConfig.frozen) {
      if (sheetConfig.frozen.rows > 0) {
        sheet.setFrozenRows(sheetConfig.frozen.rows);
      }
      if (sheetConfig.frozen.columns > 0) {
        sheet.setFrozenColumns(sheetConfig.frozen.columns);
      }
    }
  });
  
  console.log('必要なシートの作成完了');
}

/**
 * 統合ビューのヘッダーを生成（設計書準拠）
 */
function generateIntegratedViewHeaders(type) {
  if (type === 'terminal') {
    // 端末系統合ビュー: A-AT (46列) - 設計書準拠
    return [
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
  } else {
    // プリンタ・その他系統合ビュー: A-AU (47列) - 設計書準拠
    return [
      // マスタシートデータ（A-D列）
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
  }
}


/**
 * 統合ビューの設定
 */
function setupIntegratedViews() {
  console.log('統合ビューの設定開始...');
  
  // PropertiesServiceからスプレッドシートIDを取得
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID_DESTINATION') || properties.getProperty('SPREADSHEET_ID_MAIN');
  
  console.log('取得したスプレッドシートID:', spreadsheetId);
  
  if (!spreadsheetId) {
    throw new Error('スプレッドシートIDが設定されていません。先にsetupPropertiesService()を実行してください。');
  }
  
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log('スプレッドシート取得成功:', spreadsheet ? spreadsheet.getName() : 'null');
  } catch (error) {
    throw new Error(`スプレッドシートを開けませんでした。ID: ${spreadsheetId}, エラー: ${error.toString()}`);
  }
  
  if (!spreadsheet) {
    throw new Error(`スプレッドシートがnullです。ID: ${spreadsheetId}`);
  }
  
  // 端末系統合ビューの設定
  const terminalSheet = spreadsheet.getSheetByName('integrated_view_terminal');
  if (terminalSheet) {
    if (typeof setupIntegratedViewFormulas === 'function') {
      setupIntegratedViewFormulas(terminalSheet, 'terminal');
      console.log('端末系統合ビュー設定完了');
    } else {
      console.log('setupIntegratedViewFormulas関数が見つかりません。統合ビューの設定をスキップします。');
    }
  }
  
  // プリンタ・その他系統合ビューの設定
  const printerSheet = spreadsheet.getSheetByName('integrated_view_printer_other');
  if (printerSheet) {
    if (typeof setupIntegratedViewFormulas === 'function') {
      setupIntegratedViewFormulas(printerSheet, 'printer_other');
      console.log('プリンタ・その他系統合ビュー設定完了');
    } else {
      console.log('setupIntegratedViewFormulas関数が見つかりません。統合ビューの設定をスキップします。');
    }
  }
  
  console.log('統合ビューの設定完了');
}

/**
 * トリガーの設定
 */
function setupTriggers() {
  console.log('トリガーの設定開始...');
  
  // PropertiesServiceからスプレッドシートIDを取得
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID_DESTINATION') || properties.getProperty('SPREADSHEET_ID_MAIN');
  
  if (!spreadsheetId) {
    throw new Error('スプレッドシートIDが設定されていません。先にsetupPropertiesService()を実行してください。');
  }
  
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw new Error(`スプレッドシートを開けませんでした。ID: ${spreadsheetId}, エラー: ${error.toString()}`);
  }
  
  if (!spreadsheet) {
    throw new Error(`スプレッドシートがnullです。ID: ${spreadsheetId}`);
  }
  
  // 既存トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  console.log('既存トリガーを削除しました');
  
  // onFormSubmitトリガー（フォーム送信時）
  // 端末ステータス収集シートやプリンタステータス収集シートが存在する場合に設定
  const terminalSheet = spreadsheet.getSheetByName('端末ステータス収集');
  const printerSheet = spreadsheet.getSheetByName('プリンタステータス収集');
  
  if (terminalSheet || printerSheet) {
    ScriptApp.newTrigger('autoUpdateIntegratedViewOnSubmit')
      .forSpreadsheet(spreadsheet)
      .onFormSubmit()
      .create();
    console.log('フォーム送信のonFormSubmitトリガー設定完了');
  } else {
    console.log('収集シートが見つからないため、onFormSubmitトリガーはスキップしました');
  }
  
  // onChangeトリガー（マスタシート変更監視）
  ScriptApp.newTrigger('autoUpdateIntegratedViewOnChange')
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();
  console.log('マスタシート変更監視のonChangeトリガー設定完了');
  
  // timeBasedトリガー（深夜2:00の日次再構築）
  ScriptApp.newTrigger('autoRebuildAllIntegratedViews')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  console.log('日次再構築のtimeBasedトリガー設定完了');
  
  console.log('トリガーの設定完了');
}

/**
 * 設定確認用関数
 */
function checkSpreadsheetSetup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const properties = PropertiesService.getScriptProperties().getProperties();
  
  console.log('=== スプレッドシート設定確認 ===');
  console.log('スプレッドシートID:', spreadsheet.getId());
  console.log('スプレッドシート名:', spreadsheet.getName());
  
  console.log('\n--- PropertiesService設定 ---');
  for (const key in properties) {
    if (key.includes('EMAIL') || key.includes('URL')) {
      console.log(`${key}: ${properties[key] || '(未設定)'}`);
    } else {
      console.log(`${key}: ${properties[key]}`);
    }
  }
  
  console.log('\n--- シート一覧 ---');
  const sheets = spreadsheet.getSheets();
  sheets.forEach(sheet => {
    console.log(`- ${sheet.getName()} (${sheet.getLastRow()}行)`);
  });
  
  console.log('\n--- トリガー一覧 ---');
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    console.log(`- ${trigger.getHandlerFunction()} (${trigger.getEventType()})`);
  });
  
  return '設定確認完了';
}