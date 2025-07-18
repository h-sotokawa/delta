// 統合ビュー表示のテスト関数
function testIntegratedViewDisplay() {
  console.log('=== 統合ビュー表示テスト ===');
  
  try {
    // 1. バックエンドでのデータ取得をテスト
    console.log('\n1. バックエンドデータ取得テスト:');
    const backendResponse = getSpreadsheetData('osaka', 'all', 'INTEGRATED_VIEW_TERMINAL');
    
    if (backendResponse.success && backendResponse.data && backendResponse.data.length > 0) {
      const headers = backendResponse.data[0];
      console.log('  取得成功');
      console.log('  列数:', headers.length);
      console.log('  最初の10列:', headers.slice(0, 10));
      console.log('  最後の10列:', headers.slice(-10));
      
      // データ行の確認
      if (backendResponse.data.length > 1) {
        const firstDataRow = backendResponse.data[1];
        console.log('\n  最初のデータ行（最初の10列）:');
        for (let i = 0; i < Math.min(10, firstDataRow.length); i++) {
          console.log(`    ${headers[i]}: ${firstDataRow[i] || '(空)'}`);
        }
      }
    } else {
      console.log('  取得失敗:', backendResponse.error || 'データなし');
    }
    
    // 2. 統合ビューシートの直接確認
    console.log('\n2. 統合ビューシート直接確認:');
    const sheet = getIntegratedViewTerminalSheet();
    const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log('  シート列数:', sheetHeaders.length);
    
    // 3. 列順序の確認（reorderColumnsForNormalDataの影響を確認）
    console.log('\n3. 列順序の確認:');
    console.log('  期待される列順（最初の15列）:');
    const expectedOrder = [
      '端末番号', '管理番号_HW', '管理番号_SN', '現在の設置先', '機器管理番号', 
      '導入時デスク番号', '廃却日', 'タイムスタンプ', '連番', '拠点', 
      '拠点名', 'サーバ室内外', 'フロア階数', '島番号', '座席/デスク番号'
    ];
    
    for (let i = 0; i < Math.min(expectedOrder.length, sheetHeaders.length); i++) {
      const match = sheetHeaders[i] === expectedOrder[i] ? '✓' : '✗';
      console.log(`    ${i+1}. ${match} 期待: ${expectedOrder[i]}, 実際: ${sheetHeaders[i]}`);
    }
    
    // 4. フロントエンドで適用される処理の確認
    console.log('\n4. フロントエンド処理の影響:');
    console.log('  dataTypeId: INTEGRATED_VIEW_TERMINAL');
    console.log('  createTable内でreorderColumnsForNormalDataがスキップされているか確認');
    
    // 5. 全列の一覧
    console.log('\n5. 全列一覧（46列期待）:');
    sheetHeaders.forEach((header, index) => {
      console.log(`  ${index + 1}. ${header}`);
    });
    
    return {
      success: true,
      backendColumns: backendResponse.data ? backendResponse.data[0].length : 0,
      sheetColumns: sheetHeaders.length,
      expectedColumns: 46,
      message: `バックエンド: ${backendResponse.data ? backendResponse.data[0].length : 0}列, シート: ${sheetHeaders.length}列`
    };
    
  } catch (error) {
    console.error('テストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// 列順序の問題を詳細に調査
function analyzeColumnOrderIssue() {
  console.log('=== 列順序問題の詳細調査 ===');
  
  try {
    // 1. 統合ビューシートの元データ
    const sheet = getIntegratedViewTerminalSheet();
    const originalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 2. getSpreadsheetDataでの取得
    const response = getSpreadsheetData('osaka', 'all', 'INTEGRATED_VIEW_TERMINAL');
    const fetchedHeaders = response.data ? response.data[0] : [];
    
    console.log('元のシート列数:', originalHeaders.length);
    console.log('取得データ列数:', fetchedHeaders.length);
    
    // 3. 差分を確認
    if (originalHeaders.length !== fetchedHeaders.length) {
      console.log('\n列数が異なります！');
      console.log('差:', originalHeaders.length - fetchedHeaders.length);
    }
    
    // 4. 順序の違いを確認
    console.log('\n列順序の比較:');
    const maxLength = Math.max(originalHeaders.length, fetchedHeaders.length);
    let differences = 0;
    
    for (let i = 0; i < maxLength; i++) {
      const original = originalHeaders[i] || '(なし)';
      const fetched = fetchedHeaders[i] || '(なし)';
      
      if (original !== fetched) {
        differences++;
        console.log(`  位置 ${i+1}: 元="${original}" → 取得="${fetched}"`);
      }
    }
    
    console.log(`\n合計 ${differences} 個の違いが見つかりました。`);
    
    // 5. 特定の重要な列の位置を確認
    const importantColumns = ['タイムスタンプ', '連番', '拠点', '拠点名', 'サーバ室内外'];
    console.log('\n重要な列の位置:');
    
    importantColumns.forEach(colName => {
      const originalIndex = originalHeaders.indexOf(colName);
      const fetchedIndex = fetchedHeaders.indexOf(colName);
      console.log(`  ${colName}:`);
      console.log(`    元の位置: ${originalIndex + 1} (${originalIndex >= 0 ? originalHeaders[originalIndex] : 'なし'})`);
      console.log(`    取得位置: ${fetchedIndex + 1} (${fetchedIndex >= 0 ? fetchedHeaders[fetchedIndex] : 'なし'})`);
    });
    
  } catch (error) {
    console.error('調査エラー:', error);
  }
}
