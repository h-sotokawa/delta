// 拠点フィルタリング問題のデバッグ関数
function debugLocationFilterIssue() {
  console.log('=== 拠点フィルタリング問題デバッグ ===');
  
  try {
    // 1. 統合ビューシートの内容を確認
    console.log('\n1. 統合ビューシートの内容確認:');
    const sheet = getIntegratedViewTerminalSheet();
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    
    console.log('  総行数（ヘッダー含む）:', allData.length);
    console.log('  列数:', headers.length);
    
    // ヘッダーのインデックスを取得
    const managementNumberIndex = headers.indexOf('拠点管理番号');
    const locationNameIndex = headers.indexOf('拠点名');
    
    console.log('  拠点管理番号列インデックス:', managementNumberIndex);
    console.log('  拠点名列インデックス:', locationNameIndex);
    
    // 最初の5行のデータを表示
    console.log('\n  最初の5行のデータ:');
    for (let i = 1; i <= Math.min(5, allData.length - 1); i++) {
      const row = allData[i];
      console.log(`    行${i}: 拠点管理番号="${row[managementNumberIndex]}", 拠点名="${row[locationNameIndex]}"`);
    }
    
    // 2. フィルター関数の動作確認
    console.log('\n2. applyFiltersToData関数のテスト:');
    
    // osaka でフィルタリング
    const osakaFilter = { location: 'osaka' };
    const osakaFiltered = applyFiltersToData(allData, osakaFilter);
    console.log(`  osaka フィルター結果: ${osakaFiltered.length - 1}行（ヘッダー除く）`);
    
    // 大阪本社 でフィルタリング
    const osakahonshaFilter = { location: '大阪本社' };
    const osakahonshaFiltered = applyFiltersToData(allData, osakahonshaFilter);
    console.log(`  大阪本社 フィルター結果: ${osakahonshaFiltered.length - 1}行（ヘッダー除く）`);
    
    // OSAKA でフィルタリング
    const OSAKAFilter = { location: 'OSAKA' };
    const OSAKAFiltered = applyFiltersToData(allData, OSAKAFilter);
    console.log(`  OSAKA フィルター結果: ${OSAKAFiltered.length - 1}行（ヘッダー除く）`);
    
    // 3. getViewSheetData の動作確認
    console.log('\n3. getViewSheetData関数のテスト:');
    
    // フィルターなし
    const noFilterResult = getViewSheetData('integrated_view_terminal', {});
    console.log('  フィルターなし結果:', noFilterResult.success ? `${noFilterResult.data.length - 1}行` : 'エラー');
    
    // osaka フィルター
    const osakaViewResult = getViewSheetData('integrated_view_terminal', { location: 'osaka' });
    console.log('  osaka フィルター結果:', osakaViewResult.success ? `${osakaViewResult.data.length - 1}行` : 'エラー');
    
    // 4. getSpreadsheetData の動作確認
    console.log('\n4. getSpreadsheetData関数のテスト:');
    
    // all指定
    const allResult = getSpreadsheetData('all', 'all', 'INTEGRATED_VIEW_TERMINAL');
    console.log('  all指定結果:', allResult.success ? `${allResult.data.length - 1}行` : allResult.error);
    
    // osaka指定
    const osakaResult = getSpreadsheetData('osaka', 'all', 'INTEGRATED_VIEW_TERMINAL');
    console.log('  osaka指定結果:', osakaResult.success ? `${osakaResult.data.length - 1}行` : osakaResult.error);
    
    if (osakaResult.success && osakaResult.data.length > 1) {
      console.log('\n  osaka指定で取得されたデータの確認:');
      for (let i = 1; i <= Math.min(3, osakaResult.data.length - 1); i++) {
        const row = osakaResult.data[i];
        console.log(`    行${i}: 拠点管理番号="${row[0]}", 拠点名="${row[locationNameIndex]}"`);
      }
    }
    
    // 5. 拠点コードの分布を確認
    console.log('\n5. 統合ビュー内の拠点コード分布:');
    const locationCodeCount = {};
    for (let i = 1; i < allData.length; i++) {
      const managementNumber = allData[i][managementNumberIndex];
      if (managementNumber) {
        const locationCode = managementNumber.split('_')[0];
        locationCodeCount[locationCode] = (locationCodeCount[locationCode] || 0) + 1;
      }
    }
    
    Object.entries(locationCodeCount).forEach(([code, count]) => {
      console.log(`  ${code}: ${count}件`);
    });
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

// フィルター条件の詳細デバッグ
function debugFilterCondition() {
  console.log('=== フィルター条件の詳細デバッグ ===');
  
  try {
    const sheet = getIntegratedViewTerminalSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const managementNumberIndex = headers.indexOf('拠点管理番号');
    const locationNameIndex = headers.indexOf('拠点名');
    
    // テストデータ（最初のデータ行）
    if (data.length > 1) {
      const testRow = data[1];
      const managementNumber = testRow[managementNumberIndex];
      const locationName = testRow[locationNameIndex];
      const locationCode = managementNumber ? managementNumber.split('_')[0] : '';
      
      console.log('\nテストデータ:');
      console.log('  拠点管理番号:', managementNumber);
      console.log('  拠点コード:', locationCode);
      console.log('  拠点名:', locationName);
      
      // 各種フィルター条件でのテスト
      const testFilters = ['osaka', 'OSAKA', '大阪本社', 'Osaka'];
      
      console.log('\nフィルター条件テスト:');
      testFilters.forEach(filter => {
        console.log(`\n  フィルター: "${filter}"`);
        
        // 拠点名での比較
        const nameMatch = locationName === filter;
        console.log(`    拠点名比較: ${locationName} === ${filter} → ${nameMatch}`);
        
        // 拠点コードでの比較（大文字小文字を区別しない）
        const codeMatch = locationCode.toLowerCase() === filter.toLowerCase();
        console.log(`    拠点コード比較: ${locationCode.toLowerCase()} === ${filter.toLowerCase()} → ${codeMatch}`);
        
        // マッピングでの比較
        const locationMapping = {
          'osaka': ['OSAKA', '大阪本社'],
          'kobe': ['KOBE', '神戸支社'],
          'himeji': ['HIMEJI', '姫路営業所'],
          'kyoto': ['KYOTO', '京都支店']
        };
        
        const mappingKey = filter.toLowerCase();
        const mappingMatch = locationMapping[mappingKey] && 
                           (locationMapping[mappingKey].includes(locationCode) || 
                            locationMapping[mappingKey].includes(locationName));
        console.log(`    マッピング比較: ${mappingMatch}`);
        
        const wouldPass = nameMatch || codeMatch || mappingMatch;
        console.log(`    → このフィルターを通過: ${wouldPass}`);
      });
    }
    
  } catch (error) {
    console.error('デバッグエラー:', error);
  }
}