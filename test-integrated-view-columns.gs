// test-integrated-view-columns.gs
function testUpdateIntegratedViews() {
  console.log('=== 統合ビュー列数テスト ===');
  
  try {
    // 実際のデータを取得
    const terminalData = getTerminalMasterData();
    const statusData = getLatestStatusCollectionData();
    const locationData = getLocationMasterData();
    
    console.log('データ取得結果:');
    console.log('  端末マスタデータ件数:', terminalData.length);
    console.log('  ステータスデータ件数:', Object.keys(statusData).length);
    console.log('  拠点マスタデータ件数:', Object.keys(locationData).length);
    
    if (terminalData.length > 0) {
      const testDevice = terminalData[0];
      const managementNumber = testDevice['拠点管理番号'];
      
      console.log('\n対象デバイス:', managementNumber);
      
      // 修正された関数をテスト
      const result = integrateDeviceDataForView([testDevice], statusData, locationData, 'terminal');
      
      console.log('\n統合結果:');
      console.log('  行数:', result.length);
      if (result.length > 0) {
        console.log('  列数:', result[0].length);
        console.log('  期待値: 46列');
        console.log('  結果:', result[0].length === 46 ? '✅ 正常' : '❌ 異常');
        
        // 最初の10列と最後の10列を表示
        const row = result[0];
        console.log('\n最初の10列:');
        for (let i = 0; i < Math.min(10, row.length); i++) {
          console.log(`  ${i+1}: ${row[i] || '(空)'}`)
        }
        
        console.log('\n最後の10列:');
        for (let i = Math.max(0, row.length - 10); i < row.length; i++) {
          console.log(`  ${i+1}: ${row[i] || '(空)'}`)
        }
      }
    }
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

// 実行
// testUpdateIntegratedViews(); // コメントアウト：手動実行時のみコメントを外してください
