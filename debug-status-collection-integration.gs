// ステータス収集データの統合をデバッグする関数
function debugStatusCollectionIntegration() {
  console.log('=== ステータス収集データ統合デバッグ ===');
  
  try {
    // 1. ステータス収集データの取得を確認
    console.log('\n1. ステータス収集データ取得:');
    const statusData = getLatestStatusCollectionData();
    const statusKeys = Object.keys(statusData);
    console.log('  取得したステータスデータ数:', statusKeys.length);
    
    if (statusKeys.length > 0) {
      // 最初のレコードを詳細に確認
      const firstKey = statusKeys[0];
      const firstRecord = statusData[firstKey];
      console.log('\n  最初のレコード (管理番号: ' + firstKey + '):');
      
      // 2-* 列の存在を確認
      const section2Columns = [
        '2-1.預り機返却の有無',
        '2-2.依頼者',
        '2-3.備考',
        '2-3.修理の必要性',  // プリンタ・その他用
        '2-4.備考'  // プリンタ・その他用
      ];
      
      console.log('\n  セクション2の列の値:');
      section2Columns.forEach(col => {
        if (firstRecord[col] !== undefined) {
          console.log(`    ${col}: "${firstRecord[col]}"`);
        }
      });
      
      // すべての列を確認
      console.log('\n  全列一覧:');
      Object.keys(firstRecord).forEach(key => {
        if (key.startsWith('2-')) {
          console.log(`    ${key}: "${firstRecord[key]}"`);
        }
      });
    }
    
    // 2. 統合処理のテスト
    console.log('\n2. 統合処理テスト:');
    const terminalData = getTerminalMasterData();
    const locationData = getLocationMasterData();
    
    if (terminalData.length > 0) {
      const firstDevice = terminalData[0];
      console.log('  テスト端末: ' + firstDevice['拠点管理番号']);
      
      // integrateDeviceDataForView を呼び出し
      const integratedData = integrateDeviceDataForView([firstDevice], statusData, locationData, 'terminal');
      
      if (integratedData.length > 0) {
        const integratedRow = integratedData[0];
        console.log('  統合後の列数:', integratedRow.length);
        
        // 2-* 列に相当する位置を確認（W-Y列、22-24）
        console.log('\n  セクション2相当の列（W-Y列、インデックス22-24）:');
        console.log('    W列 (2-1.預り機返却の有無): "' + integratedRow[22] + '"');
        console.log('    X列 (2-2.依頼者): "' + integratedRow[23] + '"');
        console.log('    Y列 (2-3.備考): "' + integratedRow[24] + '"');
        
        // 実際のシートでの確認
        console.log('\n3. 統合ビューシートの実データ確認:');
        const sheet = getIntegratedViewTerminalSheet();
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        // W-Y列のヘッダーを確認
        console.log('  W列ヘッダー (23列目): ' + headers[22]);
        console.log('  X列ヘッダー (24列目): ' + headers[23]);
        console.log('  Y列ヘッダー (25列目): ' + headers[24]);
        
        // 実際のデータ行を確認
        if (sheet.getLastRow() > 1) {
          const firstDataRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
          console.log('\n  最初のデータ行のW-Y列:');
          console.log('    W列: "' + firstDataRow[22] + '"');
          console.log('    X列: "' + firstDataRow[23] + '"');
          console.log('    Y列: "' + firstDataRow[24] + '"');
          
          // 全体的な列の内容確認
          console.log('\n  データがある列の確認:');
          let emptyColumns = [];
          let filledColumns = [];
          
          for (let i = 0; i < firstDataRow.length; i++) {
            if (firstDataRow[i]) {
              filledColumns.push(`${i+1}(${headers[i]})`);
            } else {
              emptyColumns.push(`${i+1}(${headers[i]})`);
            }
          }
          
          console.log('  データがある列数:', filledColumns.length);
          console.log('  空の列数:', emptyColumns.length);
          
          // セクション別に確認
          console.log('\n  セクション別データ確認:');
          console.log('    セクション1 (A-G): ' + firstDataRow.slice(0, 7).filter(v => v).length + '/7列にデータあり');
          console.log('    セクション2 (H-AN): ' + firstDataRow.slice(7, 40).filter(v => v).length + '/33列にデータあり');
          console.log('    セクション3 (AO-AT): ' + firstDataRow.slice(40, 46).filter(v => v).length + '/6列にデータあり');
        }
      }
    }
    
  } catch (error) {
    console.error('デバッグエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}