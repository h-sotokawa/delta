// ========================================
// ステータス収集データの詳細テスト
// ========================================

/**
 * ステータス収集シートの詳細を確認
 */
function testStatusCollectionSheet() {
  console.log('=== ステータス収集シートの詳細確認 ===');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // 端末ステータス収集シートを確認
    const sheetName = '端末ステータス収集';
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log('端末ステータス収集シートが見つかりません');
      return;
    }
    
    console.log('シート名:', sheetName);
    console.log('最終行:', sheet.getLastRow());
    console.log('最終列:', sheet.getLastColumn());
    
    // ヘッダー行を取得
    if (sheet.getLastRow() > 0) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      console.log('\nヘッダー列（全' + headers.length + '列）:');
      headers.forEach((header, index) => {
        console.log(`  列${index + 1}: ${header}`);
      });
      
      // データ行のサンプルを取得
      if (sheet.getLastRow() > 1) {
        const sampleRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
        console.log('\n最初のデータ行の内容:');
        headers.forEach((header, index) => {
          if (sampleRow[index]) {
            console.log(`  ${header}: ${sampleRow[index]}`);
          }
        });
      }
    }
    
    // getLatestStatusCollectionData関数のテスト
    console.log('\n=== getLatestStatusCollectionData関数のテスト ===');
    const statusData = getLatestStatusCollectionData();
    const keys = Object.keys(statusData);
    console.log('取得されたデータ件数:', keys.length);
    
    if (keys.length > 0) {
      const firstKey = keys[0];
      const firstData = statusData[firstKey];
      console.log('\n最初のデータ（キー: ' + firstKey + '):');
      console.log('オブジェクトのキー数:', Object.keys(firstData).length);
      
      // 全てのキーと値を表示
      console.log('\n全てのフィールド:');
      for (const key in firstData) {
        console.log(`  ${key}: ${firstData[key]}`);
      }
    }
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

/**
 * 拠点管理番号のマッチングをテスト
 */
function testManagementNumberMatching() {
  console.log('=== 拠点管理番号マッチングテスト ===');
  
  try {
    // マスタデータから拠点管理番号を取得
    const terminalData = getTerminalMasterData();
    console.log('端末マスタデータ件数:', terminalData.length);
    
    if (terminalData.length > 0) {
      console.log('\n端末マスタの拠点管理番号:');
      terminalData.forEach((device, index) => {
        console.log(`  ${index + 1}. ${device['拠点管理番号']}`);
      });
    }
    
    // ステータス収集データのキーを確認
    const statusData = getLatestStatusCollectionData();
    const statusKeys = Object.keys(statusData);
    console.log('\nステータス収集データのキー:');
    statusKeys.forEach((key, index) => {
      console.log(`  ${index + 1}. ${key}`);
    });
    
    // マッチング確認
    console.log('\nマッチング結果:');
    terminalData.forEach(device => {
      const mgmtNum = device['拠点管理番号'];
      if (statusData[mgmtNum]) {
        console.log(`  ✓ ${mgmtNum} - マッチしました`);
      } else {
        console.log(`  ✗ ${mgmtNum} - マッチしません`);
        // 類似のキーを探す
        const similarKeys = statusKeys.filter(key => 
          key.toLowerCase().includes(mgmtNum.toLowerCase()) || 
          mgmtNum.toLowerCase().includes(key.toLowerCase())
        );
        if (similarKeys.length > 0) {
          console.log(`    類似キー: ${similarKeys.join(', ')}`);
        }
      }
    });
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}