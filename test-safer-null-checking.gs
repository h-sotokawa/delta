// 安全なnullチェック修正後のテスト
function testSaferNullChecking() {
  console.log('=== 安全なnullチェック修正後のテスト ===');
  
  try {
    // 1. 空のステータスデータでテスト
    const testDevice = {
      '拠点管理番号': 'TEST_terminal_Device_TEST123_001',
      'カテゴリ': 'terminal',
      '機種名': 'Test Device',
      '製造番号': 'TEST123',
      '資産管理番号': 'ASSET_TEST',
      'ソフトウェア': 'Test Software',
      'OS': 'Test OS',
      'formURL': 'https://forms.gle/test',
      'QRコードURL': 'https://qr.test.com'
    };
    
    const emptyStatusData = {
      'TEST_terminal_Device_TEST123_001': {} // 空のオブジェクト
    };
    
    const testLocationData = {
      'TEST': {
        locationName: 'テスト拠点',
        jurisdiction: 'テスト管轄'
      }
    };
    
    console.log('\\n1. 空のステータスデータでのテスト:');
    const result1 = integrateDeviceData([testDevice], emptyStatusData, testLocationData, 'terminal');
    console.log('結果の行数:', result1.length);
    if (result1.length > 0) {
      console.log('結果の列数:', result1[0].length);
      console.log('期待値: 46列');
      console.log('結果:', result1[0].length === 46 ? '✅ 正常' : '❌ 異常');
      
      // 最初の10列を確認
      console.log('\\n最初の10列:');
      for (let i = 0; i < Math.min(10, result1[0].length); i++) {
        console.log(`  ${i+1}: ${result1[0][i] || '(空)'}`)
      }
    }
    
    // 2. ステータスデータがない場合でテスト
    console.log('\\n2. ステータスデータがない場合でのテスト:');
    const noStatusData = {};
    const result2 = integrateDeviceData([testDevice], noStatusData, testLocationData, 'terminal');
    console.log('結果の行数:', result2.length);
    if (result2.length > 0) {
      console.log('結果の列数:', result2[0].length);
      console.log('期待値: 46列');
      console.log('結果:', result2[0].length === 46 ? '✅ 正常' : '❌ 異常');
    }
    
    // 3. 実際のデータでテスト
    console.log('\\n3. 実際のデータでのテスト:');
    const terminalData = getTerminalMasterData();
    const statusData = getLatestStatusCollectionData();
    const locationData = getLocationMasterData();
    
    if (terminalData.length > 0) {
      const result3 = integrateDeviceData([terminalData[0]], statusData, locationData, 'terminal');
      console.log('結果の行数:', result3.length);
      if (result3.length > 0) {
        console.log('結果の列数:', result3[0].length);
        console.log('期待値: 46列');
        console.log('結果:', result3[0].length === 46 ? '✅ 正常' : '❌ 異常');
      }
    }
    
    console.log('\\n=== テスト完了 ===');
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

// プリンタ・その他系のテスト
function testPrinterOtherSaferNullChecking() {
  console.log('=== プリンタ・その他系の安全なnullチェックテスト ===');
  
  try {
    const testDevice = {
      '拠点管理番号': 'TEST_printer_Device_TEST123_001',
      'カテゴリ': 'printer',
      '機種名': 'Test Printer',
      '製造番号': 'TEST123',
      'formURL': 'https://forms.gle/test',
      'QRコードURL': 'https://qr.test.com'
    };
    
    const emptyStatusData = {
      'TEST_printer_Device_TEST123_001': {} // 空のオブジェクト
    };
    
    const testLocationData = {
      'TEST': {
        locationName: 'テスト拠点',
        jurisdiction: 'テスト管轄'
      }
    };
    
    console.log('\\n1. 空のステータスデータでのテスト:');
    const result1 = integrateDeviceData([testDevice], emptyStatusData, testLocationData, 'printer');
    console.log('結果の行数:', result1.length);
    if (result1.length > 0) {
      console.log('結果の列数:', result1[0].length);
      console.log('期待値: 47列');
      console.log('結果:', result1[0].length === 47 ? '✅ 正常' : '❌ 異常');
    }
    
    console.log('\\n=== プリンタ・その他系テスト完了 ===');
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}