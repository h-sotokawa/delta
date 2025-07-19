/**
 * 端末ビューで null が返される問題のデバッグ用関数
 */
function debugTerminalViewIssue() {
  console.log('=== 端末ビューnull問題デバッグ開始 ===');
  
  try {
    console.log('1. getIntegratedViewData("", "terminal")を呼び出します...');
    
    // 実際の関数呼び出し
    const result = getIntegratedViewData('', 'terminal');
    
    console.log('2. 呼び出し結果:');
    console.log('   - 結果オブジェクト:', result);
    console.log('   - success:', result ? result.success : 'undefined');
    console.log('   - error:', result ? result.error : 'undefined');
    console.log('   - data:', result && result.data ? '配列（長さ: ' + result.data.length + '）' : 'null/undefined');
    console.log('   - metadata:', result ? result.metadata : 'undefined');
    
    // もしnullが返された場合の詳細分析
    if (result === null || result === undefined) {
      console.log('3. NULLが返されました - 詳細分析を開始...');
      
      // getIntegratedViewTerminalSheet()を直接テスト
      console.log('3.1. getIntegratedViewTerminalSheet()を直接テスト:');
      try {
        const sheet = getIntegratedViewTerminalSheet();
        console.log('   - シート取得成功:', sheet ? sheet.getName() : 'null');
        
        if (sheet) {
          console.log('   - 最終行:', sheet.getLastRow());
          console.log('   - 最終列:', sheet.getLastColumn());
        }
      } catch (sheetError) {
        console.error('   - シート取得エラー:', sheetError.toString());
      }
      
      // スプレッドシートID確認
      console.log('3.2. スプレッドシートID確認:');
      try {
        const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
        console.log('   - SPREADSHEET_ID_DESTINATION:', spreadsheetId ? '設定済み' : '未設定');
        
        if (spreadsheetId) {
          const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
          console.log('   - スプレッドシート取得成功:', spreadsheet.getName());
          
          const terminalSheetName = VIEW_SHEET_TYPES.INTEGRATED_TERMINAL;
          console.log('   - 端末シート名:', terminalSheetName);
          
          const sheet = spreadsheet.getSheetByName(terminalSheetName);
          console.log('   - 端末シート存在確認:', sheet ? '存在' : '存在しない');
        }
      } catch (propError) {
        console.error('   - プロパティ/スプレッドシート確認エラー:', propError.toString());
      }
    }
    
    // プリンタ・その他ビューとの比較
    console.log('4. プリンタ・その他ビューとの比較:');
    try {
      const printerResult = getIntegratedViewData('', 'printer_other');
      console.log('   - プリンタ・その他ビュー結果:');
      console.log('     - success:', printerResult ? printerResult.success : 'undefined');
      console.log('     - data存在:', printerResult && printerResult.data ? 'あり（長さ: ' + printerResult.data.length + '）' : 'なし');
    } catch (printerError) {
      console.error('   - プリンタ・その他ビューエラー:', printerError.toString());
    }
    
    console.log('5. ログシステムの確認:');
    // ログを確認（最新の5件）
    try {
      // getLogEntries関数が存在する場合のみ実行
      if (typeof getLogEntries === 'function') {
        const logs = getLogEntries(5);
        console.log('   - 最新ログ（5件）:');
        logs.forEach((log, index) => {
          console.log(`     ${index + 1}. ${log.timestamp}: ${log.message} - ${JSON.stringify(log.details || {})}`);
        });
      } else {
        console.log('   - ログ取得関数が見つかりません');
      }
    } catch (logError) {
      console.error('   - ログ取得エラー:', logError.toString());
    }
    
  } catch (error) {
    console.error('デバッグ中にエラーが発生しました:', error.toString());
    console.error('スタックトレース:', error.stack);
  }
  
  console.log('=== 端末ビューnull問題デバッグ完了 ===');
}

/**
 * getIntegratedViewData関数の詳細トレース
 */
function traceGetIntegratedViewDataExecution() {
  console.log('=== getIntegratedViewData実行トレース開始 ===');
  
  const location = '';
  const viewType = 'terminal';
  
  console.log('パラメータ:', { location, viewType });
  
  try {
    console.log('1. startPerformanceTimer呼び出し...');
    const startTime = startPerformanceTimer();
    console.log('   - startTime:', startTime);
    
    console.log('2. addLog呼び出し...');
    addLog('統合ビューデータ取得開始', { location, viewType });
    
    console.log('3. ビューシート取得...');
    let sheet;
    if (viewType === 'terminal') {
      console.log('3.1. getIntegratedViewTerminalSheet()呼び出し...');
      sheet = getIntegratedViewTerminalSheet();
      console.log('   - シート:', sheet ? sheet.getName() : 'null');
      
      addLog('端末系統合ビューシート取得', { sheetName: sheet ? sheet.getName() : 'null' });
    }
    
    if (!sheet) {
      throw new Error('統合ビューシートが見つかりません');
    }
    
    console.log('4. シート情報取得...');
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    console.log('   - lastRow:', lastRow);
    console.log('   - lastColumn:', lastColumn);
    
    addLog('シート情報', { 
      sheetName: sheet.getName(),
      lastRow: lastRow,
      lastColumn: lastColumn 
    });
    
    console.log('5. データ取得処理...');
    
    if (lastRow === 0) {
      console.log('   - シートが空（行なし）');
      const result = {
        success: true,
        data: [],
        metadata: {
          viewType: viewType,
          location: location,
          totalRows: 0
        }
      };
      console.log('   - 返却結果:', result);
      return result;
    }
    
    if (lastRow === 1) {
      console.log('   - ヘッダーのみ');
      const headers = sheet.getRange(1, 1, 1, lastColumn).getValues();
      const result = {
        success: true,
        data: headers,
        metadata: {
          viewType: viewType,
          location: location,
          totalRows: 0
        }
      };
      console.log('   - 返却結果:', result);
      return result;
    }
    
    console.log('6. 全データ取得...');
    const allData = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    console.log('   - データ取得完了:', allData.length, '行');
    
    addLog('全データ取得完了', { 
      totalRows: allData.length,
      headerRow: allData[0] ? allData[0].slice(0, 5) : 'no header'
    });
    
    console.log('7. フィルタリング処理...');
    let filteredData = allData;
    if (location) {
      console.log('   - 拠点フィルタリング実行');
      // フィルタリング処理（省略）
    } else {
      console.log('   - 拠点フィルタリングなし（全拠点）');
    }
    
    console.log('8. 結果作成...');
    const result = {
      success: true,
      data: filteredData,
      metadata: {
        viewType: viewType,
        location: location || '全拠点',
        totalRows: filteredData.length - 1
      }
    };
    
    console.log('9. 結果返却:', {
      success: result.success,
      dataLength: result.data ? result.data.length : 'null',
      metadata: result.metadata
    });
    
    endPerformanceTimer(startTime, '統合ビューデータ取得完了');
    return result;
    
  } catch (error) {
    console.error('トレース中にエラー:', error.toString());
    console.error('スタックトレース:', error.stack);
    
    const errorResult = {
      success: false,
      error: error.toString(),
      data: null
    };
    
    console.log('エラー結果返却:', errorResult);
    return errorResult;
  }
  
  console.log('=== getIntegratedViewData実行トレース完了 ===');
}

/**
 * シート存在確認とデータ状況チェック
 */
function checkSheetExistenceAndData() {
  console.log('=== シート存在確認とデータ状況チェック ===');
  
  try {
    // スプレッドシートID確認
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    console.log('1. SPREADSHEET_ID_DESTINATION:', spreadsheetId ? 'OK' : 'NG - 未設定');
    
    if (!spreadsheetId) {
      console.error('スプレッドシートIDが設定されていません');
      return;
    }
    
    // スプレッドシート取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log('2. スプレッドシート取得: OK -', spreadsheet.getName());
    
    // 全シート一覧
    const allSheets = spreadsheet.getSheets();
    console.log('3. 全シート一覧:');
    allSheets.forEach((sheet, index) => {
      console.log(`   ${index + 1}. ${sheet.getName()} (${sheet.getLastRow()}行, ${sheet.getLastColumn()}列)`);
    });
    
    // 端末系統合ビューシート確認
    const terminalSheetName = VIEW_SHEET_TYPES.INTEGRATED_TERMINAL;
    console.log('4. 端末系統合ビューシート確認:');
    console.log('   - 期待シート名:', terminalSheetName);
    
    const terminalSheet = spreadsheet.getSheetByName(terminalSheetName);
    if (terminalSheet) {
      console.log('   - 存在: YES');
      console.log('   - 行数:', terminalSheet.getLastRow());
      console.log('   - 列数:', terminalSheet.getLastColumn());
      
      if (terminalSheet.getLastRow() > 0) {
        const headers = terminalSheet.getRange(1, 1, 1, Math.min(terminalSheet.getLastColumn(), 10)).getValues()[0];
        console.log('   - ヘッダー（最初の10列）:', headers);
      }
      
      if (terminalSheet.getLastRow() > 1) {
        const firstRow = terminalSheet.getRange(2, 1, 1, Math.min(terminalSheet.getLastColumn(), 5)).getValues()[0];
        console.log('   - 最初のデータ行（最初の5列）:', firstRow);
      }
    } else {
      console.log('   - 存在: NO');
    }
    
    // プリンタ・その他系統合ビューシート確認
    const printerSheetName = VIEW_SHEET_TYPES.INTEGRATED_PRINTER_OTHER;
    console.log('5. プリンタ・その他系統合ビューシート確認:');
    console.log('   - 期待シート名:', printerSheetName);
    
    const printerSheet = spreadsheet.getSheetByName(printerSheetName);
    if (printerSheet) {
      console.log('   - 存在: YES');
      console.log('   - 行数:', printerSheet.getLastRow());
      console.log('   - 列数:', printerSheet.getLastColumn());
    } else {
      console.log('   - 存在: NO');
    }
    
  } catch (error) {
    console.error('チェック中にエラー:', error.toString());
    console.error('スタックトレース:', error.stack);
  }
  
  console.log('=== シート存在確認とデータ状況チェック完了 ===');
}