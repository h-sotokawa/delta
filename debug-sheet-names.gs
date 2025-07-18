// シート名の確認用デバッグ関数
function debugSheetNames() {
  console.log('=== 全シート名確認 ===');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    const sheets = spreadsheet.getSheets();
    console.log('総シート数:', sheets.length);
    
    sheets.forEach((sheet, index) => {
      console.log(`${index + 1}. ${sheet.getName()} (行数: ${sheet.getLastRow()})`);
    });
    
    console.log('\n=== ステータス収集シート確認 ===');
    const statusSheets = ['端末ステータス収集', 'プリンタステータス収集', 'その他ステータス収集'];
    
    for (const sheetName of statusSheets) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        console.log(`✓ ${sheetName}: ${sheet.getLastRow()}行`);
        if (sheet.getLastRow() > 1) {
          const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
          console.log(`  ヘッダー数: ${headers.length}`);
          console.log(`  拠点管理番号列: ${headers.indexOf('0-0.拠点管理番号')}`);
          
          // 最初のデータ行を確認
          const firstRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
          const mgmtNumIndex = headers.indexOf('0-0.拠点管理番号');
          if (mgmtNumIndex >= 0) {
            console.log(`  最初のデータの拠点管理番号: ${firstRow[mgmtNumIndex]}`);
          }
        }
      } else {
        console.log(`✗ ${sheetName}: 見つかりません`);
      }
    }
    
  } catch (error) {
    console.error('エラー:', error);
  }
}

// ステータス収集データの詳細確認
function debugStatusCollectionData() {
  console.log('=== ステータス収集データの詳細確認 ===');
  
  try {
    const statusData = getLatestStatusCollectionData();
    const keys = Object.keys(statusData);
    
    console.log('取得されたステータスデータ件数:', keys.length);
    console.log('キー一覧:', keys);
    
    if (keys.length > 0) {
      const firstKey = keys[0];
      const firstData = statusData[firstKey];
      console.log(`\n最初のデータ (${firstKey}):`);
      console.log('フィールド数:', Object.keys(firstData).length);
      
      // 重要なフィールドのみ表示
      const importantFields = ['タイムスタンプ', '9999.管理ID', '0-4.ステータス', '0-1.担当者'];
      for (const field of importantFields) {
        console.log(`  ${field}: ${firstData[field] || '(空)'}`);
      }
    }
    
  } catch (error) {
    console.error('エラー:', error);
  }
}