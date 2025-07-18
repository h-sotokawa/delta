// 拠点フィルタリングのテスト関数
function testLocationFiltering() {
  console.log('=== 拠点フィルタリングテスト ===');
  
  try {
    // 1. 統合ビューの全データを確認
    console.log('\n1. 統合ビューの全データ確認:');
    const sheet = getIntegratedViewTerminalSheet();
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    
    // 拠点管理番号列のインデックスを取得
    const managementNumberIndex = headers.indexOf('拠点管理番号');
    const locationNameIndex = headers.indexOf('拠点名');
    
    console.log('  総行数:', allData.length - 1); // ヘッダーを除く
    
    // 拠点ごとのデータ数を集計
    const locationCount = {};
    for (let i = 1; i < allData.length; i++) {
      const managementNumber = allData[i][managementNumberIndex];
      const locationName = allData[i][locationNameIndex];
      
      if (managementNumber) {
        const locationCode = managementNumber.split('_')[0];
        locationCount[locationCode] = (locationCount[locationCode] || 0) + 1;
        console.log(`  行${i}: ${managementNumber} → 拠点コード: ${locationCode}, 拠点名: ${locationName}`);
      }
    }
    
    console.log('\n  拠点ごとのデータ数:');
    Object.entries(locationCount).forEach(([code, count]) => {
      console.log(`    ${code}: ${count}件`);
    });
    
    // 2. getSpreadsheetData で拠点指定してデータ取得
    console.log('\n2. 拠点フィルタリングテスト（osaka指定）:');
    const osakaResult = getSpreadsheetData('osaka', 'all', 'INTEGRATED_VIEW_TERMINAL');
    
    if (osakaResult.success) {
      console.log('  取得成功');
      console.log('  データ行数:', osakaResult.data.length - 1); // ヘッダーを除く
      
      // 最初の数行を確認
      if (osakaResult.data.length > 1) {
        console.log('\n  最初の3件:');
        for (let i = 1; i <= Math.min(3, osakaResult.data.length - 1); i++) {
          const row = osakaResult.data[i];
          console.log(`    ${i}. ${row[managementNumberIndex]} (${row[locationNameIndex]})`);
        }
      }
      
      // すべての行が大阪のデータか確認
      let nonOsakaCount = 0;
      for (let i = 1; i < osakaResult.data.length; i++) {
        const managementNumber = osakaResult.data[i][managementNumberIndex];
        if (managementNumber) {
          const locationCode = managementNumber.split('_')[0];
          if (locationCode.toLowerCase() !== 'osaka') {
            nonOsakaCount++;
            console.log(`  警告: 大阪以外のデータ発見 - ${managementNumber}`);
          }
        }
      }
      
      if (nonOsakaCount === 0) {
        console.log('  ✅ すべて大阪のデータです');
      } else {
        console.log(`  ❌ ${nonOsakaCount}件の大阪以外のデータが含まれています`);
      }
    } else {
      console.log('  取得失敗:', osakaResult.error);
    }
    
    // 3. 他の拠点でもテスト
    console.log('\n3. 他の拠点でのフィルタリングテスト:');
    const testLocations = ['kobe', 'himeji'];
    
    testLocations.forEach(location => {
      const result = getSpreadsheetData(location, 'all', 'INTEGRATED_VIEW_TERMINAL');
      if (result.success) {
        console.log(`\n  ${location}:`);
        console.log(`    データ行数: ${result.data.length - 1}`);
        
        // 拠点コードを確認
        const locationCodes = new Set();
        for (let i = 1; i < result.data.length; i++) {
          const managementNumber = result.data[i][managementNumberIndex];
          if (managementNumber) {
            const locationCode = managementNumber.split('_')[0];
            locationCodes.add(locationCode);
          }
        }
        console.log(`    含まれる拠点コード: ${Array.from(locationCodes).join(', ')}`);
      } else {
        console.log(`  ${location}: 取得失敗 - ${result.error}`);
      }
    });
    
    // 4. allでの取得テスト
    console.log('\n4. 全拠点取得テスト（all指定）:');
    const allResult = getSpreadsheetData('all', 'all', 'INTEGRATED_VIEW_TERMINAL');
    if (allResult.success) {
      console.log('  取得成功');
      console.log('  データ行数:', allResult.data.length - 1);
      
      // すべての拠点コードを集計
      const allLocationCodes = new Set();
      for (let i = 1; i < allResult.data.length; i++) {
        const managementNumber = allResult.data[i][managementNumberIndex];
        if (managementNumber) {
          const locationCode = managementNumber.split('_')[0];
          allLocationCodes.add(locationCode);
        }
      }
      console.log('  含まれる拠点コード:', Array.from(allLocationCodes).join(', '));
    } else {
      console.log('  取得失敗:', allResult.error);
    }
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}