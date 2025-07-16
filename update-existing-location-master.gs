/**
 * 既存の拠点マスタシートに不足している列を追加する関数
 * この関数を一度実行することで、既存のシートを新しい構造に更新できます
 */
function updateExistingLocationMasterStructure() {
  try {
    console.log('=== 拠点マスタシート構造更新開始 ===');
    
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('拠点マスタ');
    
    if (!sheet) {
      console.log('拠点マスタシートが存在しません');
      return;
    }
    
    // 現在のヘッダーを取得
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log('現在のヘッダー:', headers);
    
    // 必要な列の定義
    const requiredHeaders = {
      '拠点ID': 100,
      '拠点名': 150,
      '拠点コード': 120,
      '管轄': 100,
      'グループメール': 200,
      '作成日時': 150,
      'ステータス変更通知': 150,
      'ステータス': 100
    };
    
    // 不足している列を特定
    const missingHeaders = [];
    for (const header in requiredHeaders) {
      if (!headers.includes(header)) {
        missingHeaders.push(header);
      }
    }
    
    if (missingHeaders.length === 0) {
      console.log('すべての必要な列が既に存在します');
      return;
    }
    
    console.log('不足している列:', missingHeaders);
    
    // 新しい列を追加
    let currentColumn = sheet.getLastColumn() + 1;
    missingHeaders.forEach(header => {
      sheet.getRange(1, currentColumn).setValue(header);
      sheet.setColumnWidth(currentColumn, requiredHeaders[header]);
      
      // ヘッダー行のフォーマット
      const headerCell = sheet.getRange(1, currentColumn);
      headerCell.setBackground('#4472C4');
      headerCell.setFontColor('#FFFFFF');
      headerCell.setFontWeight('bold');
      
      // デフォルト値を設定（必要に応じて）
      if (header === 'ステータス') {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          for (let row = 2; row <= lastRow; row++) {
            sheet.getRange(row, currentColumn).setValue('active');
          }
        }
      }
      
      currentColumn++;
    });
    
    console.log('拠点マスタシート構造更新完了');
    console.log('追加された列:', missingHeaders);
    
    return {
      success: true,
      addedColumns: missingHeaders
    };
    
  } catch (error) {
    console.error('エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}