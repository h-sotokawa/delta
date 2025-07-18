// 列数検証用の簡単なテスト
function validateColumnCount() {
  console.log('=== 列数検証テスト ===');
  
  // テスト用データ
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
  
  const testLocationData = {
    'TEST': {
      locationName: 'テスト拠点',
      jurisdiction: 'テスト管轄'
    }
  };
  
  // 1. 空のステータスデータ
  console.log('1. 空のステータスデータでのテスト:');
  const emptyStatusData = {
    'TEST_terminal_Device_TEST123_001': {}
  };
  
  const result1 = integrateDeviceData([testDevice], emptyStatusData, testLocationData, 'terminal');
  console.log('行数:', result1.length);
  console.log('列数:', result1.length > 0 ? result1[0].length : 0);
  console.log('期待値: 46列');
  console.log('結果:', result1.length > 0 && result1[0].length === 46 ? '✅ 正常' : '❌ 異常');
  
  // 2. ステータスデータなし
  console.log('\\n2. ステータスデータなしでのテスト:');
  const noStatusData = {};
  
  const result2 = integrateDeviceData([testDevice], noStatusData, testLocationData, 'terminal');
  console.log('行数:', result2.length);
  console.log('列数:', result2.length > 0 ? result2[0].length : 0);
  console.log('期待値: 46列');
  console.log('結果:', result2.length > 0 && result2[0].length === 46 ? '✅ 正常' : '❌ 異常');
  
  // 3. プリンタ・その他系のテスト
  console.log('\\n3. プリンタ・その他系のテスト:');
  const printerDevice = {
    '拠点管理番号': 'TEST_printer_Device_TEST123_001',
    'カテゴリ': 'printer',
    '機種名': 'Test Printer',
    '製造番号': 'TEST123',
    'formURL': 'https://forms.gle/test',
    'QRコードURL': 'https://qr.test.com'
  };
  
  const result3 = integrateDeviceData([printerDevice], {}, testLocationData, 'printer');
  console.log('行数:', result3.length);
  console.log('列数:', result3.length > 0 ? result3[0].length : 0);
  console.log('期待値: 47列');
  console.log('結果:', result3.length > 0 && result3[0].length === 47 ? '✅ 正常' : '❌ 異常');
  
  console.log('\\n=== 検証完了 ===');
}