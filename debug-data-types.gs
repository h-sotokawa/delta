/**
 * データタイプマスタのデバッグ関数
 */

/**
 * データタイプマスタの内容を確認
 */
function debugDataTypeMaster() {
  try {
    console.log('=== データタイプマスタデバッグ開始 ===');
    
    // スプレッドシートIDを確認
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_MAIN');
    console.log('SPREADSHEET_ID_MAIN:', spreadsheetId);
    
    if (!spreadsheetId) {
      console.error('SPREADSHEET_ID_MAINが設定されていません');
      return;
    }
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log('スプレッドシート名:', spreadsheet.getName());
    
    // データタイプマスタシートを確認
    const sheet = spreadsheet.getSheetByName('データタイプマスタ');
    if (!sheet) {
      console.error('データタイプマスタシートが見つかりません');
      return;
    }
    
    console.log('シート名:', sheet.getName());
    console.log('最終行:', sheet.getLastRow());
    console.log('最終列:', sheet.getLastColumn());
    
    // ヘッダー行を確認
    if (sheet.getLastRow() >= 1) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      console.log('ヘッダー:', headers);
    }
    
    // データ行を確認
    if (sheet.getLastRow() >= 2) {
      const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      const data = dataRange.getValues();
      
      console.log('データ行数:', data.length);
      
      data.forEach((row, index) => {
        console.log(`行${index + 2}:`, {
          dataTypeId: row[0],
          dataTypeName: row[1],
          description: row[2],
          displayOrder: row[3],
          filterCondition: row[4],
          dataSourceConfig: row[5],
          displayColumnConfig: row[6],
          status: row[7],
          createdAt: row[8],
          updatedAt: row[9]
        });
      });
    }
    
    // getDataTypeMaster関数をテスト
    console.log('\n=== getDataTypeMaster(true) テスト ===');
    const activeResult = getDataTypeMaster(true);
    console.log('成功:', activeResult.success);
    console.log('データタイプ数:', activeResult.dataTypes.length);
    activeResult.dataTypes.forEach(dt => {
      console.log('- ' + dt.dataTypeId + ': ' + dt.dataTypeName);
    });
    
    console.log('\n=== getDataTypeMaster(false) テスト ===');
    const allResult = getDataTypeMaster(false);
    console.log('成功:', allResult.success);
    console.log('データタイプ数:', allResult.dataTypes.length);
    allResult.dataTypes.forEach(dt => {
      console.log('- ' + dt.dataTypeId + ': ' + dt.dataTypeName + ' (status: ' + dt.status + ')');
    });
    
    console.log('=== デバッグ終了 ===');
    
  } catch (error) {
    console.error('デバッグエラー:', error);
  }
}

/**
 * データタイプマスタを再初期化（既存データを削除して再作成）
 */
function reinitializeDataTypeMaster() {
  try {
    console.log('=== データタイプマスタ再初期化開始 ===');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('データタイプマスタ');
    
    if (!sheet) {
      console.error('データタイプマスタシートが見つかりません');
      return;
    }
    
    // 既存データをクリア（ヘッダー行は残す）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
      console.log('既存データをクリアしました');
    }
    
    // 初期データを再投入
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const initialData = [
      ['AUDIT', '監査データ', '機器の監査用データ表示', 1, '', '{}', '{}', 'active', now, now],
      ['SUMMARY', 'サマリーデータ', '拠点別サマリー表示', 2, '', '{}', '{}', 'active', now, now],
      ['INTEGRATED_VIEW_TERMINAL', '統合ビュー（端末系）', 'Server、Desktop、Laptop、Tabletの統合表示', 3, '', '{}', '{}', 'active', now, now],
      ['INTEGRATED_VIEW_PRINTER_OTHER', '統合ビュー（プリンタ・その他系）', 'Printer、Router、Hub、Otherの統合表示', 4, '', '{}', '{}', 'active', now, now],
      ['INTEGRATED_VIEW', '統合ビュー（旧）', '全機器の統合表示（非推奨）', 5, '', '{}', '{}', 'active', now, now]
    ];
    
    const range = sheet.getRange(2, 1, initialData.length, initialData[0].length);
    range.setValues(initialData);
    
    console.log('初期データを投入しました:', initialData.length + '件');
    console.log('=== データタイプマスタ再初期化完了 ===');
    
    // 確認のためデバッグ実行
    debugDataTypeMaster();
    
  } catch (error) {
    console.error('再初期化エラー:', error);
  }
}