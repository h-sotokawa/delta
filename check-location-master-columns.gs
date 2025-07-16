/**
 * 拠点マスタシートの現在の列構成を確認する診断関数
 */
function checkLocationMasterColumns() {
  try {
    console.log('=== 拠点マスタシート列構成チェック ===');
    
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('拠点マスタ');
    
    if (!sheet) {
      console.log('拠点マスタシートが存在しません');
      return;
    }
    
    // ヘッダー行を取得
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log('現在のヘッダー:', headers);
    console.log('列数:', headers.length);
    
    // 各列のインデックスを表示
    headers.forEach((header, index) => {
      console.log(`列${index + 1}: ${header}`);
    });
    
    // 実際のデータ行を1行取得して確認
    if (sheet.getLastRow() > 1) {
      const firstDataRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      console.log('\n最初のデータ行:');
      headers.forEach((header, index) => {
        console.log(`${header}: ${firstDataRow[index]}`);
      });
    }
    
    // 期待される列構成
    console.log('\n期待される列構成:');
    console.log('1: 拠点ID');
    console.log('2: 拠点名');
    console.log('3: 拠点コード');
    console.log('4: 管轄');
    console.log('5: 作成日時');
    console.log('6: ステータス変更通知');
    
    return {
      headers: headers,
      expectedHeaders: ['拠点ID', '拠点名', '拠点コード', '管轄', '作成日時', 'ステータス変更通知']
    };
    
  } catch (error) {
    console.error('エラー:', error);
  }
}