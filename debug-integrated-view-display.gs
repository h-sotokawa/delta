// 統合ビューの表示データをデバッグする関数
function debugIntegratedViewDisplay() {
  console.log('=== 統合ビュー表示デバッグ ===');
  
  try {
    // 1. 統合ビューシートのデータを取得
    const sheet = getIntegratedViewTerminalSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    console.log('シート情報:');
    console.log('  行数:', lastRow);
    console.log('  列数:', lastCol);
    
    if (lastRow > 1) {
      // ヘッダーを取得
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      console.log('\nヘッダー列数:', headers.length);
      console.log('ヘッダー（最初の15列）:');
      for (let i = 0; i < Math.min(15, headers.length); i++) {
        console.log(`  ${i+1}: ${headers[i]}`);
      }
      
      // 最初のデータ行を取得
      const firstRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
      console.log('\n最初のデータ行（最初の15列）:');
      for (let i = 0; i < Math.min(15, firstRow.length); i++) {
        console.log(`  ${i+1}: ${firstRow[i] || '(空)'}`);
      }
      
      // ステータス関連の列（H-AN列、8-40列）を確認
      console.log('\nステータス関連列（H-N列、8-14列）:');
      for (let i = 7; i < Math.min(14, firstRow.length); i++) {
        console.log(`  ${headers[i]}: ${firstRow[i] || '(空)'}`);
      }
    }
    
    // 2. 現在のデータ更新状況を確認
    console.log('\n=== データ更新テスト ===');
    const terminalResult = updateIntegratedViewTerminal();
    console.log('更新結果:', terminalResult);
    
    // 3. 更新後のデータを確認
    if (terminalResult.success) {
      console.log('\n更新後のデータ確認:');
      const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const updatedFirstRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      console.log('更新後の列数:', updatedHeaders.length);
      console.log('更新後のステータス関連列（H-N列）:');
      for (let i = 7; i < Math.min(14, updatedFirstRow.length); i++) {
        console.log(`  ${updatedHeaders[i]}: ${updatedFirstRow[i] || '(空)'}`);
      }
    }
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

// 統合ビューの実際の列構造を確認
function checkIntegratedViewStructure() {
  console.log('=== 統合ビュー列構造確認 ===');
  
  try {
    const sheet = getIntegratedViewTerminalSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    console.log('総列数:', headers.length);
    console.log('\n全ヘッダー一覧:');
    headers.forEach((header, index) => {
      console.log(`  ${index + 1} (${String.fromCharCode(65 + index)}列): ${header}`);
    });
    
    // 期待される46列と実際の列数を比較
    console.log('\n期待値: 46列');
    console.log('実際: ' + headers.length + '列');
    console.log('差分: ' + (headers.length - 46));
    
  } catch (error) {
    console.error('エラー:', error);
  }
}