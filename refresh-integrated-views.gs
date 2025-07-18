// 統合ビューを更新して利用可能な拠点を確認
function refreshAndCheckIntegratedViews() {
  console.log('=== 統合ビュー更新と確認 ===');
  
  try {
    // 1. 統合ビューを更新
    console.log('\n1. 統合ビューを更新中...');
    const updateResult = updateIntegratedViewTerminal();
    console.log('更新結果:', updateResult);
    
    // 2. 更新後の拠点分布を確認
    console.log('\n2. 更新後の拠点分布:');
    const sheet = getIntegratedViewTerminalSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const managementNumberIndex = headers.indexOf('拠点管理番号');
    const locationCount = {};
    
    for (let i = 1; i < data.length; i++) {
      const managementNumber = data[i][managementNumberIndex];
      if (managementNumber) {
        const locationCode = managementNumber.split('_')[0];
        locationCount[locationCode] = (locationCount[locationCode] || 0) + 1;
      }
    }
    
    console.log('拠点コード別データ数:');
    Object.entries(locationCount).forEach(([code, count]) => {
      console.log(`  ${code}: ${count}件`);
    });
    
    // 3. 各拠点でのフィルタリングテスト
    console.log('\n3. 各拠点でのフィルタリングテスト:');
    const testLocations = ['osaka', 'kobe', 'himeji', 'kyoto', 'tokyo'];
    
    testLocations.forEach(location => {
      const result = getSpreadsheetData(location, 'all', 'INTEGRATED_VIEW_TERMINAL');
      console.log(`  ${location}: ${result.success ? (result.data.length - 1) + '件' : 'エラー'}`);
    });
    
    // 4. 問題のある tokyo を詳しく調査
    console.log('\n4. tokyo の詳細調査:');
    
    // 東京のデータが存在するか確認
    let tokyoCount = 0;
    let tokyoSamples = [];
    
    for (let i = 1; i < data.length; i++) {
      const managementNumber = data[i][managementNumberIndex];
      if (managementNumber) {
        const locationCode = managementNumber.split('_')[0];
        if (locationCode.toUpperCase() === 'TOKYO') {
          tokyoCount++;
          if (tokyoSamples.length < 3) {
            tokyoSamples.push({
              managementNumber: managementNumber,
              locationName: data[i][headers.indexOf('拠点名')] || ''
            });
          }
        }
      }
    }
    
    console.log('  TOKYOコードのデータ数:', tokyoCount);
    if (tokyoSamples.length > 0) {
      console.log('  サンプルデータ:');
      tokyoSamples.forEach(sample => {
        console.log(`    ${sample.managementNumber} (${sample.locationName})`);
      });
    }
    
  } catch (error) {
    console.error('エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}