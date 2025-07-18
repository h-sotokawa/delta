// ステータス収集シートの内容を確認する関数
function verifyStatusCollectionSheet() {
  console.log('=== ステータス収集シート確認 ===');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // 端末ステータス収集シートを確認
    console.log('\n1. 端末ステータス収集シート:');
    const terminalStatusSheet = spreadsheet.getSheetByName('端末ステータス収集');
    
    if (terminalStatusSheet) {
      const lastRow = terminalStatusSheet.getLastRow();
      const lastCol = terminalStatusSheet.getLastColumn();
      console.log('  行数:', lastRow);
      console.log('  列数:', lastCol);
      
      if (lastRow > 1) {
        const headers = terminalStatusSheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const firstDataRow = terminalStatusSheet.getRange(2, 1, 1, lastCol).getValues()[0];
        
        // ヘッダーを確認
        console.log('\n  ヘッダー確認:');
        headers.forEach((header, index) => {
          if (header.startsWith('2-')) {
            console.log(`    列${index+1}: ${header}`);
          }
        });
        
        // 最初のデータ行の2-*列を確認
        console.log('\n  最初のデータ行の2-*列:');
        headers.forEach((header, index) => {
          if (header.startsWith('2-')) {
            console.log(`    ${header}: "${firstDataRow[index] || '(空)'}"`);
          }
        });
        
        // データが入っている行数を確認
        console.log('\n  データ分析:');
        let rowsWithData = 0;
        let rowsWith2Data = 0;
        
        if (lastRow > 1) {
          const allData = terminalStatusSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
          
          allData.forEach((row, rowIndex) => {
            const hasAnyData = row.some(cell => cell);
            if (hasAnyData) rowsWithData++;
            
            // 2-*列にデータがあるか確認
            let has2Data = false;
            headers.forEach((header, colIndex) => {
              if (header.startsWith('2-') && row[colIndex]) {
                has2Data = true;
              }
            });
            if (has2Data) rowsWith2Data++;
          });
        }
        
        console.log('  データがある行数:', rowsWithData);
        console.log('  2-*列にデータがある行数:', rowsWith2Data);
        
        // 管理番号の例を表示
        console.log('\n  拠点管理番号の例（最初の5件）:');
        const managementNumCol = headers.indexOf('0-0.拠点管理番号');
        if (managementNumCol >= 0 && lastRow > 1) {
          const managementNumbers = terminalStatusSheet.getRange(2, managementNumCol + 1, Math.min(5, lastRow - 1), 1).getValues();
          managementNumbers.forEach((num, index) => {
            console.log(`    ${index + 1}. ${num[0]}`);
          });
        }
      } else {
        console.log('  データ行がありません');
      }
    } else {
      console.log('  シートが見つかりません');
    }
    
    // getLatestStatusCollectionData の動作確認
    console.log('\n2. getLatestStatusCollectionData関数の確認:');
    const statusData = getLatestStatusCollectionData();
    const keys = Object.keys(statusData);
    console.log('  取得したレコード数:', keys.length);
    
    if (keys.length > 0) {
      console.log('\n  最初のレコードの詳細:');
      const firstKey = keys[0];
      const firstRecord = statusData[firstKey];
      console.log('  管理番号:', firstKey);
      console.log('  全フィールド数:', Object.keys(firstRecord).length);
      
      // 2-*フィールドを表示
      console.log('\n  2-*フィールド:');
      Object.keys(firstRecord).forEach(key => {
        if (key.startsWith('2-')) {
          console.log(`    ${key}: "${firstRecord[key]}"`);
        }
      });
    }
    
  } catch (error) {
    console.error('確認エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}