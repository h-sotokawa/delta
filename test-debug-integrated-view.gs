// ========================================
// 統合ビューのデバッグ用テスト関数
// ========================================

/**
 * 統合ビュー更新のデバッグ
 */
function debugIntegratedViewUpdate() {
  console.log('=== 統合ビュー更新デバッグ開始 ===');
  
  try {
    // 1. マスタデータの確認
    console.log('\n1. 端末マスタデータの確認:');
    const terminalData = getTerminalMasterData();
    console.log('端末マスタデータ件数:', terminalData.length);
    if (terminalData.length > 0) {
      console.log('最初のデータ:');
      const firstDevice = terminalData[0];
      console.log('  拠点管理番号:', firstDevice['拠点管理番号']);
      console.log('  カテゴリ:', firstDevice['カテゴリ']);
      console.log('  機種名:', firstDevice['機種名']);
      console.log('  製造番号:', firstDevice['製造番号']);
      console.log('  資産管理番号:', firstDevice['資産管理番号']);
    }
    
    // 2. ステータス収集データの確認
    console.log('\n2. ステータス収集データの確認:');
    const statusData = getLatestStatusCollectionData();
    const statusKeys = Object.keys(statusData);
    console.log('ステータスデータ件数:', statusKeys.length);
    if (statusKeys.length > 0) {
      const firstKey = statusKeys[0];
      const firstStatus = statusData[firstKey];
      console.log('最初のステータスデータ（キー: ' + firstKey + '):');
      console.log('  タイムスタンプ:', firstStatus['タイムスタンプ']);
      console.log('  0-4.ステータス:', firstStatus['0-4.ステータス']);
      console.log('  1-1.顧客名または貸出先:', firstStatus['1-1.顧客名または貸出先']);
      console.log('  3-0.社内ステータス:', firstStatus['3-0.社内ステータス']);
    }
    
    // 3. 拠点マスタデータの確認
    console.log('\n3. 拠点マスタデータの確認:');
    const locationData = getLocationMasterData();
    const locationKeys = Object.keys(locationData);
    console.log('拠点マスタデータ件数:', locationKeys.length);
    if (locationKeys.length > 0) {
      const firstLocKey = locationKeys[0];
      console.log('最初の拠点データ（キー: ' + firstLocKey + '):');
      console.log('  拠点名:', locationData[firstLocKey].locationName);
      console.log('  管轄:', locationData[firstLocKey].jurisdiction);
    }
    
    // 4. integrateDeviceData関数の詳細テスト
    console.log('\n4. integrateDeviceData関数の詳細テスト:');
    if (terminalData.length > 0) {
      // 最初の3件のデータで統合処理をテスト
      const testDevices = terminalData.slice(0, 3);
      const integratedData = integrateDeviceData(testDevices, statusData, locationData, 'terminal');
      
      console.log('統合されたデータ件数:', integratedData.length);
      if (integratedData.length > 0) {
        console.log('\n統合データの詳細（最初の行）:');
        const firstRow = integratedData[0];
        console.log('列数:', firstRow.length);
        
        // 期待される列構造と実際のデータを比較
        const expectedHeaders = [
          '拠点管理番号', 'カテゴリ', '機種名', '製造番号', '資産管理番号',
          'ソフトウェア', 'OS', 'タイムスタンプ', '9999.管理ID',
          '0-0.拠点管理番号', '0-1.担当者', '0-2.EMシステムズの社員ですか？',
          '0-3.所属会社', '0-4.ステータス'
        ];
        
        for (let i = 0; i < Math.min(14, firstRow.length); i++) {
          console.log(`  列${i+1} (${expectedHeaders[i] || '?'}): ${firstRow[i]}`);
        }
      }
    }
    
    // 5. 実際に統合ビューを更新してみる
    console.log('\n5. 統合ビュー更新の実行:');
    const updateResult = updateIntegratedViewTerminal();
    console.log('更新結果:', updateResult);
    
    // 6. 更新後のシートデータを確認
    if (updateResult.success) {
      console.log('\n6. 更新後のシートデータ確認:');
      const sheet = getIntegratedViewTerminalSheet();
      if (sheet.getLastRow() > 1) {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const firstDataRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        console.log('シートの列数:', headers.length);
        console.log('\n最初のデータ行の内容:');
        for (let i = 0; i < Math.min(20, headers.length); i++) {
          console.log(`  ${headers[i]}: ${firstDataRow[i]}`);
        }
      }
    }
    
    console.log('\n=== デバッグ完了 ===');
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

/**
 * 単一デバイスの統合処理をデバッグ
 */
function debugSingleDeviceIntegration() {
  console.log('=== 単一デバイスの統合処理デバッグ ===');
  
  try {
    // テスト用のデバイスデータを手動で作成
    const testDevice = {
      '拠点管理番号': 'OSAKA_desktop_ST210E_aaa_001',
      'カテゴリ': 'desktop',
      '機種名': 'ST210E',
      '製造番号': 'aaa',
      '資産管理番号': 'ASSET001',
      'ソフトウェア': 'Windows 10',
      'OS': 'Windows',
      'formURL': 'https://example.com/form',
      'QRコードURL': 'https://example.com/qr'
    };
    
    // テスト用のステータスデータ
    const testStatus = {
      'タイムスタンプ': '2025/07/18 10:00:00',
      '9999.管理ID': 'MGT001',
      '0-0.拠点管理番号': 'OSAKA_desktop_ST210E_aaa_001',
      '0-1.担当者': '山田太郎',
      '0-2.EMシステムズの社員ですか？': 'はい',
      '0-3.所属会社': 'EMシステムズ',
      '0-4.ステータス': '1.貸出中',
      '1-1.顧客名または貸出先': 'テスト顧客',
      '1-2.顧客番号': 'CUST001',
      '1-3.住所': '大阪府大阪市',
      '3-0.社内ステータス': '2.回収後社内保管',
      '3-0-1.棚卸しフラグ': 'TRUE',
      '3-0-2.拠点': 'OSAKA'
    };
    
    // テスト用の拠点データ
    const testLocation = {
      'OSAKA': {
        locationName: '大阪',
        jurisdiction: '関西'
      }
    };
    
    console.log('テストデバイス:', testDevice);
    console.log('\nテストステータス:', testStatus);
    console.log('\nテスト拠点:', testLocation);
    
    // 統合処理を実行
    const statusMap = {
      'OSAKA_desktop_ST210E_aaa_001': testStatus
    };
    
    const result = integrateDeviceData([testDevice], statusMap, testLocation, 'terminal');
    console.log('\n統合結果:');
    console.log('行数:', result.length);
    if (result.length > 0) {
      console.log('列数:', result[0].length);
      console.log('\n統合データの内容:');
      const row = result[0];
      const headers = [
        '拠点管理番号', 'カテゴリ', '機種名', '製造番号', '資産管理番号',
        'ソフトウェア', 'OS', 'タイムスタンプ', '9999.管理ID',
        '0-0.拠点管理番号', '0-1.担当者', '0-2.EMシステムズの社員ですか？',
        '0-3.所属会社', '0-4.ステータス', '1-1.顧客名または貸出先'
      ];
      
      for (let i = 0; i < Math.min(15, row.length); i++) {
        console.log(`  ${i+1}. ${headers[i] || '?'}: ${row[i]}`);
      }
    }
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}