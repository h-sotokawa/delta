// 利用可能な拠点を確認する関数
function checkAvailableLocations() {
  console.log('=== 利用可能な拠点確認 ===');
  
  try {
    // 1. 統合ビューシートのデータを確認
    const sheet = getIntegratedViewTerminalSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const managementNumberIndex = headers.indexOf('拠点管理番号');
    const locationNameIndex = headers.indexOf('拠点名');
    
    if (managementNumberIndex < 0) {
      console.error('拠点管理番号列が見つかりません');
      return;
    }
    
    // 拠点情報を集計
    const locationInfo = {};
    
    for (let i = 1; i < data.length; i++) {
      const managementNumber = data[i][managementNumberIndex];
      const locationName = data[i][locationNameIndex] || '';
      
      if (managementNumber) {
        const locationCode = managementNumber.split('_')[0];
        
        if (!locationInfo[locationCode]) {
          locationInfo[locationCode] = {
            code: locationCode,
            names: new Set(),
            count: 0,
            sampleNumbers: []
          };
        }
        
        locationInfo[locationCode].count++;
        if (locationName) {
          locationInfo[locationCode].names.add(locationName);
        }
        if (locationInfo[locationCode].sampleNumbers.length < 3) {
          locationInfo[locationCode].sampleNumbers.push(managementNumber);
        }
      }
    }
    
    console.log('\n統合ビューに含まれる拠点:');
    console.log('-------------------------------');
    
    Object.entries(locationInfo).forEach(([code, info]) => {
      console.log(`\n拠点コード: ${code}`);
      console.log(`  データ件数: ${info.count}`);
      console.log(`  拠点名: ${Array.from(info.names).join(', ') || '(なし)'}`);
      console.log(`  サンプル管理番号:`);
      info.sampleNumbers.forEach(num => {
        console.log(`    - ${num}`);
      });
    });
    
    // 2. 端末マスタの拠点分布も確認
    console.log('\n\n端末マスタの拠点分布:');
    console.log('-------------------------------');
    
    const terminalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('端末マスタ');
    if (terminalSheet) {
      const terminalData = terminalSheet.getDataRange().getValues();
      const terminalHeaders = terminalData[0];
      const terminalMgmtIndex = terminalHeaders.indexOf('拠点管理番号');
      
      const terminalLocationCount = {};
      
      for (let i = 1; i < terminalData.length; i++) {
        const mgmtNumber = terminalData[i][terminalMgmtIndex];
        if (mgmtNumber) {
          const locCode = mgmtNumber.split('_')[0];
          terminalLocationCount[locCode] = (terminalLocationCount[locCode] || 0) + 1;
        }
      }
      
      Object.entries(terminalLocationCount).forEach(([code, count]) => {
        console.log(`  ${code}: ${count}件`);
      });
    }
    
    // 3. 拠点マスタの内容を確認
    console.log('\n\n拠点マスタの内容:');
    console.log('-------------------------------');
    
    const locationMaster = getLocationMaster();
    locationMaster.forEach(loc => {
      console.log(`  ${loc['拠点コード']} - ${loc['拠点名']} (${loc['管轄']})`);
    });
    
  } catch (error) {
    console.error('エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

// 特定の拠点でフィルタリングテスト
function testLocationFilter(locationId) {
  console.log(`\n=== ${locationId} でのフィルタリングテスト ===`);
  
  try {
    const result = getSpreadsheetData(locationId, 'all', 'INTEGRATED_VIEW_TERMINAL');
    
    if (result.success) {
      console.log('取得成功');
      console.log('データ行数（ヘッダー除く）:', result.data.length - 1);
      
      if (result.data.length > 1) {
        const headers = result.data[0];
        const managementNumberIndex = headers.indexOf('拠点管理番号');
        const locationNameIndex = headers.indexOf('拠点名');
        
        console.log('\n最初の3件:');
        for (let i = 1; i <= Math.min(3, result.data.length - 1); i++) {
          const row = result.data[i];
          console.log(`  ${i}. ${row[managementNumberIndex]} (${row[locationNameIndex]})`);
        }
      }
    } else {
      console.log('取得失敗:', result.error);
    }
    
  } catch (error) {
    console.error('エラー:', error);
  }
}