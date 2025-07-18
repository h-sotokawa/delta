// ========================================
// テスト用マスタデータ作成スクリプト
// ========================================

/**
 * すべてのマスタシートにテストデータを作成
 */
function setupAllTestMasterData() {
  console.log('=== すべてのマスタシートのテストデータ作成開始 ===');
  
  try {
    // 端末マスタのテストデータを作成
    setupTerminalMasterTestData();
    
    // プリンタマスタのテストデータを作成
    setupPrinterMasterTestData();
    
    // その他マスタのテストデータを作成
    setupOtherMasterTestData();
    
    // 拠点マスタのテストデータを作成
    setupLocationMasterTestData();
    
    console.log('=== すべてのマスタシートのテストデータ作成完了 ===');
    
  } catch (error) {
    console.error('テストデータ作成エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

/**
 * 端末マスタのテストデータを作成
 */
function setupTerminalMasterTestData() {
  console.log('端末マスタのテストデータ作成開始');
  
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName('端末マスタ');
  
  if (!sheet) {
    console.error('端末マスタシートが見つかりません');
    return;
  }
  
  // ヘッダー行を設定
  const headers = [
    '拠点管理番号', 'カテゴリ', '機種名', '製造番号', '資産管理番号',
    'ソフトウェア', 'OS', 'formURL', 'QRコードURL'
  ];
  
  // 既存データをクリア
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // テストデータ
  const testData = [
    [
      'OSAKA_desktop_ThinkPad_ABC123_001',
      'desktop',
      'ThinkPad X1 Carbon',
      'ABC123',
      'ASSET001',
      'Windows 11 Pro',
      'Windows',
      'https://forms.gle/terminal001',
      'https://qr.codes/terminal001'
    ],
    [
      'OSAKA_desktop_ThinkPad_DEF456_002',
      'desktop',
      'ThinkPad T14',
      'DEF456',
      'ASSET002',
      'Windows 11 Pro',
      'Windows',
      'https://forms.gle/terminal002',
      'https://qr.codes/terminal002'
    ],
    [
      'KOBE_laptop_MacBook_GHI789_001',
      'laptop',
      'MacBook Pro 14',
      'GHI789',
      'ASSET003',
      'macOS Sonoma',
      'macOS',
      'https://forms.gle/terminal003',
      'https://qr.codes/terminal003'
    ],
    [
      'KOBE_laptop_MacBook_JKL012_002',
      'laptop',
      'MacBook Air 13',
      'JKL012',
      'ASSET004',
      'macOS Sonoma',
      'macOS',
      'https://forms.gle/terminal004',
      'https://qr.codes/terminal004'
    ],
    [
      'HIMEJI_server_PowerEdge_MNO345_001',
      'server',
      'Dell PowerEdge R750',
      'MNO345',
      'ASSET005',
      'Windows Server 2022',
      'Windows Server',
      'https://forms.gle/terminal005',
      'https://qr.codes/terminal005'
    ],
    [
      'HIMEJI_server_PowerEdge_PQR678_002',
      'server',
      'Dell PowerEdge R730',
      'PQR678',
      'ASSET006',
      'Windows Server 2019',
      'Windows Server',
      'https://forms.gle/terminal006',
      'https://qr.codes/terminal006'
    ],
    [
      'OSAKA_tablet_iPad_STU901_001',
      'tablet',
      'iPad Pro 11',
      'STU901',
      'ASSET007',
      'iPadOS 17',
      'iPadOS',
      'https://forms.gle/terminal007',
      'https://qr.codes/terminal007'
    ],
    [
      'KOBE_tablet_Surface_VWX234_001',
      'tablet',
      'Surface Pro 9',
      'VWX234',
      'ASSET008',
      'Windows 11 Pro',
      'Windows',
      'https://forms.gle/terminal008',
      'https://qr.codes/terminal008'
    ]
  ];
  
  // データを書き込み
  sheet.getRange(2, 1, testData.length, testData[0].length).setValues(testData);
  
  console.log('端末マスタのテストデータ作成完了: ' + testData.length + '件');
}

/**
 * プリンタマスタのテストデータを作成
 */
function setupPrinterMasterTestData() {
  console.log('プリンタマスタのテストデータ作成開始');
  
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName('プリンタマスタ');
  
  if (!sheet) {
    console.error('プリンタマスタシートが見つかりません');
    return;
  }
  
  // ヘッダー行を設定（プリンタは資産管理番号なし）
  const headers = [
    '拠点管理番号', 'カテゴリ', '機種名', '製造番号',
    'formURL', 'QRコードURL'
  ];
  
  // 既存データをクリア
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // テストデータ
  const testData = [
    [
      'OSAKA_printer_LaserJet_PRT001_001',
      'printer',
      'HP LaserJet Pro 4025n',
      'PRT001',
      'https://forms.gle/printer001',
      'https://qr.codes/printer001'
    ],
    [
      'OSAKA_printer_LaserJet_PRT002_002',
      'printer',
      'HP LaserJet Pro 4035n',
      'PRT002',
      'https://forms.gle/printer002',
      'https://qr.codes/printer002'
    ],
    [
      'KOBE_printer_ColorLaser_PRT003_001',
      'printer',
      'Canon Color imageCLASS',
      'PRT003',
      'https://forms.gle/printer003',
      'https://qr.codes/printer003'
    ],
    [
      'KOBE_printer_InkJet_PRT004_002',
      'printer',
      'Epson WorkForce Pro',
      'PRT004',
      'https://forms.gle/printer004',
      'https://qr.codes/printer004'
    ],
    [
      'HIMEJI_printer_Multifunction_PRT005_001',
      'printer',
      'Brother MFC-L8900CDW',
      'PRT005',
      'https://forms.gle/printer005',
      'https://qr.codes/printer005'
    ]
  ];
  
  // データを書き込み
  sheet.getRange(2, 1, testData.length, testData[0].length).setValues(testData);
  
  console.log('プリンタマスタのテストデータ作成完了: ' + testData.length + '件');
}

/**
 * その他マスタのテストデータを作成
 */
function setupOtherMasterTestData() {
  console.log('その他マスタのテストデータ作成開始');
  
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName('その他マスタ');
  
  if (!sheet) {
    console.error('その他マスタシートが見つかりません');
    return;
  }
  
  // ヘッダー行を設定（その他も資産管理番号なし）
  const headers = [
    '拠点管理番号', 'カテゴリ', '機種名', '製造番号',
    'formURL', 'QRコードURL'
  ];
  
  // 既存データをクリア
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // テストデータ
  const testData = [
    [
      'OSAKA_router_Cisco_RTR001_001',
      'router',
      'Cisco ISR 4331',
      'RTR001',
      'https://forms.gle/router001',
      'https://qr.codes/router001'
    ],
    [
      'OSAKA_router_Cisco_RTR002_002',
      'router',
      'Cisco ISR 4321',
      'RTR002',
      'https://forms.gle/router002',
      'https://qr.codes/router002'
    ],
    [
      'KOBE_hub_NetGear_HUB001_001',
      'hub',
      'NetGear ProSafe 24-Port',
      'HUB001',
      'https://forms.gle/hub001',
      'https://qr.codes/hub001'
    ],
    [
      'KOBE_hub_HPE_HUB002_002',
      'hub',
      'HPE OfficeConnect 1420',
      'HUB002',
      'https://forms.gle/hub002',
      'https://qr.codes/hub002'
    ],
    [
      'HIMEJI_other_Monitor_MON001_001',
      'other',
      'Dell UltraSharp 27" 4K',
      'MON001',
      'https://forms.gle/other001',
      'https://qr.codes/other001'
    ],
    [
      'HIMEJI_other_UPS_UPS001_001',
      'other',
      'APC Smart-UPS 1500VA',
      'UPS001',
      'https://forms.gle/other002',
      'https://qr.codes/other002'
    ]
  ];
  
  // データを書き込み
  sheet.getRange(2, 1, testData.length, testData[0].length).setValues(testData);
  
  console.log('その他マスタのテストデータ作成完了: ' + testData.length + '件');
}

/**
 * 拠点マスタのテストデータを作成
 */
function setupLocationMasterTestData() {
  console.log('拠点マスタのテストデータ作成開始');
  
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName('拠点マスタ');
  
  if (!sheet) {
    console.error('拠点マスタシートが見つかりません');
    return;
  }
  
  // ヘッダー行を設定
  const headers = [
    '拠点コード', '拠点名', '管轄', 'ステータス変更通知'
  ];
  
  // 既存データをクリア
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // テストデータ
  const testData = [
    ['OSAKA', '大阪本社', '関西', 'TRUE'],
    ['KOBE', '神戸支社', '関西', 'TRUE'],
    ['HIMEJI', '姫路営業所', '関西', 'FALSE'],
    ['TOKYO', '東京支社', '関東', 'TRUE']
  ];
  
  // データを書き込み
  sheet.getRange(2, 1, testData.length, testData[0].length).setValues(testData);
  
  console.log('拠点マスタのテストデータ作成完了: ' + testData.length + '件');
}

/**
 * テスト用ステータス収集データを作成
 */
function setupTestStatusCollectionData() {
  console.log('=== テスト用ステータス収集データ作成開始 ===');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // 端末ステータス収集データを作成
    setupTerminalStatusTestData(spreadsheet);
    
    // プリンタステータス収集データを作成
    setupPrinterStatusTestData(spreadsheet);
    
    // その他ステータス収集データを作成
    setupOtherStatusTestData(spreadsheet);
    
    console.log('=== テスト用ステータス収集データ作成完了 ===');
    
  } catch (error) {
    console.error('ステータス収集データ作成エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

/**
 * 端末ステータス収集のテストデータを作成
 */
function setupTerminalStatusTestData(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('端末ステータス収集');
  
  if (!sheet) {
    console.error('端末ステータス収集シートが見つかりません');
    return;
  }
  
  // 既存データをクリア（ヘッダーは保持）
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  
  // テストデータ（端末マスタの拠点管理番号と一致するように）
  const testData = [
    [
      new Date('2025-01-15 10:30:00'), // タイムスタンプ
      'MGT001', // 9999.管理ID
      'OSAKA_desktop_ThinkPad_ABC123_001', // 0-0.拠点管理番号
      '田中太郎', // 0-1.担当者
      'はい', // 0-2.EMシステムズの社員ですか？
      'EMシステムズ', // 0-3.所属会社
      '1.貸出中', // 0-4.ステータス
      '株式会社テスト', // 1-1.顧客名または貸出先
      'CUST001', // 1-2.顧客番号
      '大阪府大阪市北区梅田1-1-1', // 1-3.住所
      '有り', // 1-4.ユーザー機の預り有無
      '山田次郎', // 1-5.依頼者
      'テスト用貸出', // 1-6.備考
      'USER001', // 1-7.預りユーザー機のシリアルNo.
      'RECEIPT001', // 1-8.お預かり証No.
      '無し', // 2-1.預り機返却の有無
      '', // 2-2.依頼者
      '', // 2-3.備考
      '2.回収後社内保管', // 3-0.社内ステータス
      'FALSE', // 3-0-1.棚卸しフラグ
      'OSAKA', // 3-0-2.拠点
      'Microsoft Office 2021', // 3-1-1.ソフト
      'プリインストール済み', // 3-1-2.備考
      '', // 3-2-1.端末初期化の引継ぎ
      '', // 3-2-2.備考
      '', // 3-2-3.引継ぎ担当者
      '', // 3-2-4.初期化作業の引継ぎ
      '', // 4-1.所在
      '', // 4-2.持ち出し理由
      '', // 4-3.備考
      '', // 5-1.内容
      '', // 5-2.所在
      ''  // 5-3.備考
    ],
    [
      new Date('2025-01-15 11:00:00'), // タイムスタンプ
      'MGT002', // 9999.管理ID
      'KOBE_laptop_MacBook_GHI789_001', // 0-0.拠点管理番号
      '佐藤花子', // 0-1.担当者
      'はい', // 0-2.EMシステムズの社員ですか？
      'EMシステムズ', // 0-3.所属会社
      '3.社内にて保管中', // 0-4.ステータス
      '', // 1-1.顧客名または貸出先
      '', // 1-2.顧客番号
      '', // 1-3.住所
      '無し', // 1-4.ユーザー機の預り有無
      '', // 1-5.依頼者
      '在庫管理中', // 1-6.備考
      '', // 1-7.預りユーザー機のシリアルNo.
      '', // 1-8.お預かり証No.
      '', // 2-1.預り機返却の有無
      '', // 2-2.依頼者
      '', // 2-3.備考
      '2.回収後社内保管', // 3-0.社内ステータス
      'TRUE', // 3-0-1.棚卸しフラグ
      'KOBE', // 3-0-2.拠点
      'macOS Sonoma', // 3-1-1.ソフト
      '標準構成', // 3-1-2.備考
      '', // 3-2-1.端末初期化の引継ぎ
      '', // 3-2-2.備考
      '', // 3-2-3.引継ぎ担当者
      '', // 3-2-4.初期化作業の引継ぎ
      '', // 4-1.所在
      '', // 4-2.持ち出し理由
      '', // 4-3.備考
      '', // 5-1.内容
      '', // 5-2.所在
      ''  // 5-3.備考
    ],
    [
      new Date('2025-01-15 14:20:00'), // タイムスタンプ
      'MGT003', // 9999.管理ID
      'HIMEJI_server_PowerEdge_MNO345_001', // 0-0.拠点管理番号
      '鈴木一郎', // 0-1.担当者
      'はい', // 0-2.EMシステムズの社員ですか？
      'EMシステムズ', // 0-3.所属会社
      '5.修理中', // 0-4.ステータス
      '', // 1-1.顧客名または貸出先
      '', // 1-2.顧客番号
      '', // 1-3.住所
      '無し', // 1-4.ユーザー機の預り有無
      '高橋三郎', // 1-5.依頼者
      'ハードディスク故障のため修理中', // 1-6.備考
      '', // 1-7.預りユーザー機のシリアルNo.
      '', // 1-8.お預かり証No.
      '', // 2-1.預り機返却の有無
      '', // 2-2.依頼者
      '', // 2-3.備考
      '1.修理中', // 3-0.社内ステータス
      'FALSE', // 3-0-1.棚卸しフラグ
      'HIMEJI', // 3-0-2.拠点
      'Windows Server 2022', // 3-1-1.ソフト
      'データベースサーバー用', // 3-1-2.備考
      '', // 3-2-1.端末初期化の引継ぎ
      '', // 3-2-2.備考
      '', // 3-2-3.引継ぎ担当者
      '', // 3-2-4.初期化作業の引継ぎ
      '', // 4-1.所在
      '', // 4-2.持ち出し理由
      '', // 4-3.備考
      '', // 5-1.内容
      '', // 5-2.所在
      ''  // 5-3.備考
    ]
  ];
  
  // データを書き込み
  sheet.getRange(2, 1, testData.length, testData[0].length).setValues(testData);
  
  console.log('端末ステータス収集のテストデータ作成完了: ' + testData.length + '件');
}

/**
 * プリンタステータス収集のテストデータを作成
 */
function setupPrinterStatusTestData(spreadsheet) {
  // プリンタステータス収集シートの処理（簡略化）
  console.log('プリンタステータス収集のテストデータ作成をスキップ');
}

/**
 * その他ステータス収集のテストデータを作成
 */
function setupOtherStatusTestData(spreadsheet) {
  // その他ステータス収集シートの処理（簡略化）
  console.log('その他ステータス収集のテストデータ作成をスキップ');
}

/**
 * すべてのテストデータを一括作成
 */
function createAllTestData() {
  console.log('=== すべてのテストデータ一括作成開始 ===');
  
  setupAllTestMasterData();
  setupTestStatusCollectionData();
  
  console.log('=== すべてのテストデータ一括作成完了 ===');
  console.log('統合ビューの更新を実行してください: updateAllIntegratedViews()');
}