// 統合処理の詳細デバッグ関数
function debugIntegrationProcess() {
  console.log('=== 統合処理の詳細デバッグ ===');
  
  try {
    // 1. ステータス収集データの確認
    console.log('\n1. ステータス収集データの確認:');
    const statusData = getLatestStatusCollectionData();
    const statusKeys = Object.keys(statusData);
    console.log('ステータスデータ件数:', statusKeys.length);
    console.log('キー一覧:', statusKeys);
    
    if (statusKeys.length > 0) {
      const firstKey = statusKeys[0];
      const firstData = statusData[firstKey];
      console.log(`\n最初のデータ (${firstKey}):`);
      console.log('フィールド数:', Object.keys(firstData).length);
      
      // 重要なフィールドを確認
      const importantFields = ['タイムスタンプ', '9999.管理ID', '0-4.ステータス', '0-1.担当者'];
      for (const field of importantFields) {
        console.log(`  ${field}: ${firstData[field] || '(空)'}`);
      }
    }
    
    // 2. 端末マスタデータの確認
    console.log('\n2. 端末マスタデータの確認:');
    const terminalData = getTerminalMasterData();
    console.log('端末マスタデータ件数:', terminalData.length);
    
    if (terminalData.length > 0) {
      console.log('最初の端末データ:');
      const firstTerminal = terminalData[0];
      console.log(`  拠点管理番号: ${firstTerminal['拠点管理番号']}`);
      console.log(`  カテゴリ: ${firstTerminal['カテゴリ']}`);
    }
    
    // 3. 拠点マスタデータの確認
    console.log('\n3. 拠点マスタデータの確認:');
    const locationData = getLocationMasterData();
    const locationKeys = Object.keys(locationData);
    console.log('拠点マスタデータ件数:', locationKeys.length);
    console.log('拠点コード一覧:', locationKeys);
    
    // 4. integrateDeviceData関数の詳細テスト
    console.log('\n4. integrateDeviceData関数の詳細テスト:');
    if (terminalData.length > 0) {
      const testDevices = terminalData.slice(0, 1); // 最初の1件のみテスト
      const integratedData = integrateDeviceData(testDevices, statusData, locationData, 'terminal');
      
      console.log('統合データ件数:', integratedData.length);
      if (integratedData.length > 0) {
        const row = integratedData[0];
        console.log('統合データの列数:', row.length);
        console.log('統合データの内容（最初の20列）:');
        for (let i = 0; i < Math.min(20, row.length); i++) {
          console.log(`  列${i+1}: ${row[i] || '(空)'}`);
        }
      }
    }
    
    // 5. 統合ビューシートの実際のデータ確認
    console.log('\n5. 統合ビューシートの実際のデータ確認:');
    const sheet = getIntegratedViewTerminalSheet();
    if (sheet.getLastRow() > 1) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const firstRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      console.log('シートの列数:', headers.length);
      console.log('最初のデータ行（最初の15列）:');
      for (let i = 0; i < Math.min(15, firstRow.length); i++) {
        console.log(`  ${headers[i]}: ${firstRow[i] || '(空)'}`);
      }
    }
    
    console.log('\n=== デバッグ完了 ===');
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

// 簡単なテスト関数
function testSingleDeviceIntegration() {
  console.log('=== 単一デバイス統合テスト ===');
  
  try {
    // テスト用のデータを手動で作成
    const testDevice = {
      '拠点管理番号': 'OSAKA_desktop_ThinkPad_ABC123_001',
      'カテゴリ': 'desktop',
      '機種名': 'ThinkPad X1 Carbon',
      '製造番号': 'ABC123',
      '資産管理番号': 'ASSET001',
      'ソフトウェア': 'Windows 11 Pro',
      'OS': 'Windows',
      'formURL': 'https://forms.gle/terminal001',
      'QRコードURL': 'https://qr.codes/terminal001'
    };
    
    const testStatus = {
      'タイムスタンプ': '2025-01-15 10:30:00',
      '9999.管理ID': 'MGT001',
      '0-0.拠点管理番号': 'OSAKA_desktop_ThinkPad_ABC123_001',
      '0-1.担当者': '田中太郎',
      '0-2.EMシステムズの社員ですか？': 'はい',
      '0-3.所属会社': 'EMシステムズ',
      '0-4.ステータス': '1.貸出中',
      '1-1.顧客名または貸出先': '株式会社テスト',
      '1-2.顧客番号': 'CUST001',
      '1-3.住所': '大阪府大阪市北区梅田1-1-1'
    };
    
    const testLocation = {
      'OSAKA': {
        locationName: '大阪本社',
        jurisdiction: '関西'
      }
    };
    
    console.log('テストデータ準備完了');
    
    // 統合処理実行
    const statusMap = {
      'OSAKA_desktop_ThinkPad_ABC123_001': testStatus
    };
    
    const result = integrateDeviceData([testDevice], statusMap, testLocation, 'terminal');
    
    console.log('統合結果:');
    console.log('行数:', result.length);
    if (result.length > 0) {
      console.log('列数:', result[0].length);
      console.log('最初の15列:');
      for (let i = 0; i < Math.min(15, result[0].length); i++) {
        console.log(`  ${i+1}: ${result[0][i] || '(空)'}`);
      }
    }
    
  } catch (error) {
    console.error('テストエラー:', error);
  }
}