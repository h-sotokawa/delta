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
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = spreadsheet.getId();
  
  // 必須設定項目
  const properties = PropertiesService.getScriptProperties();
  
  // スプレッドシートID
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
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
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
      name: '端末ステータス収集',
      headers: ['タイムスタンプ', '拠点管理番号', '拠点', 'ステータス', '貸出開始日時', '貸出予定終了日時', '利用者', '利用部署'],
      frozen: { rows: 1, columns: 2 }
    },
    {
      name: 'プリンタステータス収集',
      headers: ['タイムスタンプ', '拠点管理番号', '拠点', 'ステータス', '設置場所', '利用部署'],
      frozen: { rows: 1, columns: 2 }
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
 * 統合ビューのヘッダーを生成
 */
function generateIntegratedViewHeaders(type) {
  const baseHeaders = ['拠点管理番号', '拠点', 'カテゴリ', '機種名'];
  
  if (type === 'terminal') {
    // 端末系統合ビュー: A-AT (46列)
    return baseHeaders.concat([
      '資産管理番号', '製造番号', '作成日時',
      'タイムスタンプ', '状態', '貸出開始日時', '貸出予定終了日時',
      '貸出日数', '利用者', '利用者ID', '利用部署',
      '預り票番号', '預り開始日', '希望返却日', 'キッティング希望',
      '端末初期化の引継ぎ', '備考', '初期化パスワード', '初期化作業の引継ぎ',
      '修理開始日', '修理完了予定日', '修理内容', '修理業者',
      '修理費用', '修理備考', '廃棄予定日', '廃棄理由',
      '廃棄承認者', '廃棄備考', '紛失日', '紛失場所',
      '警察届出', '紛失備考', '発見日', '発見場所',
      '発見状態', '発見備考', '返却確認日', '返却確認者',
      '返却備考', '拠点名', '管轄'
    ]);
  } else {
    // プリンタ・その他系統合ビュー: A-AU (47列)
    return baseHeaders.concat([
      '作成日時', 'タイムスタンプ', '状態', '設置場所',
      '利用部署', '設置日', '設置担当者', '設置備考',
      '移動元場所', '移動先場所', '移動日', '移動担当者',
      '移動理由', '移動備考', '修理開始日', '修理完了予定日',
      '修理内容', '修理業者', '修理費用', '修理備考',
      '廃棄予定日', '廃棄理由', '廃棄承認者', '廃棄備考',
      '紛失日', '紛失場所', '警察届出', '紛失備考',
      '発見日', '発見場所', '発見状態', '発見備考',
      '返却確認日', '返却確認者', '返却備考', '預り開始日',
      '希望返却日', '預り理由', '預り備考', '拠点名',
      '管轄'
    ]);
  }
}


/**
 * 統合ビューの設定
 */
function setupIntegratedViews() {
  console.log('統合ビューの設定開始...');
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 端末系統合ビューの設定
  const terminalSheet = spreadsheet.getSheetByName('integrated_view_terminal');
  if (terminalSheet) {
    setupIntegratedViewFormulas(terminalSheet, 'terminal');
    console.log('端末系統合ビュー設定完了');
  }
  
  // プリンタ・その他系統合ビューの設定
  const printerSheet = spreadsheet.getSheetByName('integrated_view_printer_other');
  if (printerSheet) {
    setupIntegratedViewFormulas(printerSheet, 'printer_other');
    console.log('プリンタ・その他系統合ビュー設定完了');
  }
  
  console.log('統合ビューの設定完了');
}

/**
 * トリガーの設定
 */
function setupTriggers() {
  console.log('トリガーの設定開始...');
  
  // 既存トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  // onFormSubmitトリガー（端末ステータス収集）
  const terminalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('端末ステータス収集');
  if (terminalSheet) {
    ScriptApp.newTrigger('updateIntegratedViewOnSubmit')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onFormSubmit()
      .create();
    console.log('端末ステータス収集のonFormSubmitトリガー設定完了');
  }
  
  // onChangeトリガー（マスタシート変更監視）
  ScriptApp.newTrigger('updateIntegratedViewOnChange')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();
  console.log('マスタシート変更監視のonChangeトリガー設定完了');
  
  // timeBasedトリガー（深夜2:00の日次再構築）
  ScriptApp.newTrigger('rebuildAllIntegratedViews')
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