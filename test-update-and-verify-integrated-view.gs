// 統合ビューを更新して検証するテスト関数
function testUpdateAndVerifyIntegratedView() {
  console.log('=== 統合ビュー更新・検証テスト ===');
  
  try {
    // 1. 更新前の状態を確認
    console.log('\n1. 更新前の統合ビューシート状態:');
    const sheet = getIntegratedViewTerminalSheet();
    const beforeUpdate = {
      rows: sheet.getLastRow(),
      cols: sheet.getLastColumn()
    };
    console.log('  行数:', beforeUpdate.rows);
    console.log('  列数:', beforeUpdate.cols);
    
    if (beforeUpdate.rows > 1) {
      const firstRow = sheet.getRange(2, 1, 1, beforeUpdate.cols).getValues()[0];
      const headers = sheet.getRange(1, 1, 1, beforeUpdate.cols).getValues()[0];
      
      // データがある列を数える
      let dataCount = 0;
      for (let i = 0; i < firstRow.length; i++) {
        if (firstRow[i]) dataCount++;
      }
      console.log('  最初の行でデータがある列数:', dataCount + '/' + firstRow.length);
    }
    
    // 2. 統合ビューを更新
    console.log('\n2. 統合ビューを更新中...');
    const updateResult = updateIntegratedViewTerminal();
    console.log('  更新結果:', updateResult);
    
    // 3. 更新後の状態を確認
    console.log('\n3. 更新後の統合ビューシート状態:');
    const afterUpdate = {
      rows: sheet.getLastRow(),
      cols: sheet.getLastColumn()
    };
    console.log('  行数:', afterUpdate.rows);
    console.log('  列数:', afterUpdate.cols);
    
    if (afterUpdate.rows > 1) {
      const headers = sheet.getRange(1, 1, 1, afterUpdate.cols).getValues()[0];
      const firstRow = sheet.getRange(2, 1, 1, afterUpdate.cols).getValues()[0];
      
      console.log('\n  列の詳細分析:');
      
      // セクションごとに分析
      const sections = [
        { name: 'マスタシートデータ (A-G)', start: 0, end: 7 },
        { name: '収集シートデータ前半 (H-T)', start: 7, end: 20 },
        { name: '収集シートデータ2-* (U-Y)', start: 20, end: 25 },
        { name: '収集シートデータ3-* (Z-AN)', start: 25, end: 40 },
        { name: '計算・参照データ (AO-AT)', start: 40, end: 46 }
      ];
      
      sections.forEach(section => {
        const sectionData = firstRow.slice(section.start, section.end);
        const filledCount = sectionData.filter(v => v).length;
        console.log(`\n  ${section.name}:`);
        console.log(`    データあり: ${filledCount}/${section.end - section.start}列`);
        
        // 各列の詳細
        for (let i = section.start; i < section.end && i < headers.length; i++) {
          const value = firstRow[i];
          if (section.name.includes('2-*') || !value) {  // 2-*セクションまたは空の値を表示
            console.log(`    ${i+1}. ${headers[i]}: "${value || '(空)'}"`);
          }
        }
      });
      
      // 問題のある2-*列を特定
      console.log('\n  2-*列の詳細確認:');
      const problemColumns = [
        'タイムスタンプ',
        '1-7.預りユーザー機のシリアルNo.(製造番号)',
        '1-8.お預かり証No.',
        '2-1.預り機返却の有無',
        '2-2.依頼者',
        '2-3.備考'
      ];
      
      problemColumns.forEach(colName => {
        const colIndex = headers.indexOf(colName);
        if (colIndex >= 0) {
          console.log(`    ${colName} (列${colIndex+1}): "${firstRow[colIndex] || '(空)'}"`);
        } else {
          console.log(`    ${colName}: ヘッダーに見つかりません`);
        }
      });
    }
    
    // 4. ステータス収集データの確認
    console.log('\n4. ステータス収集データの存在確認:');
    const statusData = getLatestStatusCollectionData();
    const statusKeys = Object.keys(statusData);
    console.log('  ステータスデータレコード数:', statusKeys.length);
    
    if (statusKeys.length > 0) {
      // 統合ビューの最初のデータと対応するステータスデータを確認
      if (afterUpdate.rows > 1) {
        const firstRow = sheet.getRange(2, 1, 1, afterUpdate.cols).getValues()[0];
        const managementNumber = firstRow[0];  // A列が拠点管理番号
        
        console.log('\n  対応確認:');
        console.log('    統合ビューの拠点管理番号:', managementNumber);
        
        if (statusData[managementNumber]) {
          console.log('    対応するステータスデータ: あり');
          const status = statusData[managementNumber];
          
          // 2-*列のデータを確認
          console.log('    ステータスデータの2-*列:');
          Object.keys(status).forEach(key => {
            if (key.startsWith('2-')) {
              console.log(`      ${key}: "${status[key]}"`);
            }
          });
        } else {
          console.log('    対応するステータスデータ: なし');
        }
      }
    }
    
    return {
      success: true,
      beforeUpdate: beforeUpdate,
      afterUpdate: afterUpdate,
      message: '統合ビュー更新・検証完了'
    };
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタックトレース:', error.stack);
    return {
      success: false,
      error: error.toString()
    };
  }
}