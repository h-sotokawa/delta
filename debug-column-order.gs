// 統合ビューシートの列順序を確認する関数
function debugColumnOrder() {
  console.log('=== 統合ビューシート列順序確認 ===');
  
  try {
    // 統合ビューシートを取得
    const sheet = getIntegratedViewTerminalSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    console.log('総列数:', headers.length);
    console.log('\n列順序（インデックス: ヘッダー名）:');
    
    headers.forEach((header, index) => {
      console.log(`${index}: ${header}`);
    });
    
    // 最初のデータ行も確認
    if (sheet.getLastRow() > 1) {
      const firstRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      console.log('\n最初のデータ行の内容:');
      headers.forEach((header, index) => {
        console.log(`${header}: ${firstRow[index] || '(空)'}`);
      });
    }
    
  } catch (error) {
    console.error('エラー:', error);
  }
}

// フロントエンドでの列順序を確認
function checkFrontendColumnOrder() {
  console.log('=== フロントエンド列順序確認 ===');
  
  try {
    // テスト用にデータを取得
    const response = getSpreadsheetData('osaka', 'all', 'INTEGRATED_VIEW_TERMINAL');
    
    if (response.success && response.data && response.data.length > 0) {
      const headers = response.data[0];
      console.log('フロントエンドに送信される列順序:');
      headers.forEach((header, index) => {
        console.log(`${index}: ${header}`);
      });
      
      if (response.data.length > 1) {
        const firstDataRow = response.data[1];
        console.log('\n最初のデータ行:');
        headers.forEach((header, index) => {
          console.log(`${header}: ${firstDataRow[index] || '(空)'}`);
        });
      }
    }
    
  } catch (error) {
    console.error('エラー:', error);
  }
}