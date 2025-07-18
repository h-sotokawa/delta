/**
 * 統合ビュー再構築関数
 * 日次バッチ処理で全統合ビューを再構築
 */

/**
 * 端末系統合ビューの再構築
 */
function rebuildTerminalIntegratedView() {
  console.log('端末系統合ビュー再構築開始...');
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const integratedSheet = spreadsheet.getSheetByName('integrated_view_terminal');
  
  if (!integratedSheet) {
    console.error('端末系統合ビューシートが見つかりません');
    return;
  }
  
  try {
    // ヘッダー行以外をクリア
    const lastRow = integratedSheet.getLastRow();
    if (lastRow > 1) {
      integratedSheet.getRange(2, 1, lastRow - 1, integratedSheet.getLastColumn()).clear();
    }
    
    // 端末マスタから全データ取得
    const terminalMaster = spreadsheet.getSheetByName('端末マスタ');
    if (!terminalMaster) return;
    
    const masterData = terminalMaster.getDataRange().getValues();
    const allIntegratedData = [];
    
    // ヘッダー行を取得
    const terminalHeaders = masterData[0];
    
    // 各端末のデータを収集（ヘッダー行をスキップ）
    for (let i = 1; i < masterData.length; i++) {
      const locationNumber = getValueByColumnName(masterData[i], terminalHeaders, '拠点管理番号');
      if (!locationNumber) continue;
      
      const integratedData = collectIntegratedData(locationNumber, 'terminal');
      if (integratedData && integratedData.length > 0) {
        allIntegratedData.push(integratedData);
      }
    }
    
    // データを一括で書き込み
    if (allIntegratedData.length > 0) {
      const range = integratedSheet.getRange(2, 1, allIntegratedData.length, allIntegratedData[0].length);
      range.setValues(allIntegratedData);
      
      // 書式設定
      formatIntegratedView(integratedSheet, 'terminal');
    }
    
    console.log(`端末系統合ビュー再構築完了: ${allIntegratedData.length}件`);
    
  } catch (error) {
    console.error('端末系統合ビュー再構築エラー:', error);
    throw error;
  }
}

/**
 * プリンタ・その他系統合ビューの再構築
 */
function rebuildPrinterOtherIntegratedView() {
  console.log('プリンタ・その他系統合ビュー再構築開始...');
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const integratedSheet = spreadsheet.getSheetByName('integrated_view_printer_other');
  
  if (!integratedSheet) {
    console.error('プリンタ・その他系統合ビューシートが見つかりません');
    return;
  }
  
  try {
    // ヘッダー行以外をクリア
    const lastRow = integratedSheet.getLastRow();
    if (lastRow > 1) {
      integratedSheet.getRange(2, 1, lastRow - 1, integratedSheet.getLastColumn()).clear();
    }
    
    const allIntegratedData = [];
    
    // プリンタマスタからデータ収集
    const printerMaster = spreadsheet.getSheetByName('プリンタマスタ');
    if (printerMaster) {
      const printerData = printerMaster.getDataRange().getValues();
      const printerHeaders = printerData[0];
      for (let i = 1; i < printerData.length; i++) {
        const locationNumber = getValueByColumnName(printerData[i], printerHeaders, '拠点管理番号');
        if (!locationNumber) continue;
        
        const integratedData = collectIntegratedData(locationNumber, 'printer_other');
        if (integratedData && integratedData.length > 0) {
          allIntegratedData.push(integratedData);
        }
      }
    }
    
    // その他マスタからデータ収集
    const otherMaster = spreadsheet.getSheetByName('その他マスタ');
    if (otherMaster) {
      const otherData = otherMaster.getDataRange().getValues();
      const otherHeaders = otherData[0];
      for (let i = 1; i < otherData.length; i++) {
        const locationNumber = getValueByColumnName(otherData[i], otherHeaders, '拠点管理番号');
        if (!locationNumber) continue;
        
        const integratedData = collectIntegratedData(locationNumber, 'printer_other');
        if (integratedData && integratedData.length > 0) {
          allIntegratedData.push(integratedData);
        }
      }
    }
    
    // データを一括で書き込み
    if (allIntegratedData.length > 0) {
      const range = integratedSheet.getRange(2, 1, allIntegratedData.length, allIntegratedData[0].length);
      range.setValues(allIntegratedData);
      
      // 書式設定
      formatIntegratedView(integratedSheet, 'printer_other');
    }
    
    console.log(`プリンタ・その他系統合ビュー再構築完了: ${allIntegratedData.length}件`);
    
  } catch (error) {
    console.error('プリンタ・その他系統合ビュー再構築エラー:', error);
    throw error;
  }
}

/**
 * 検索インデックスの再構築
 */
function rebuildSearchIndex() {
  console.log('検索インデックス再構築開始...');
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = spreadsheet.getSheetByName('search_index');
  
  if (!indexSheet) {
    console.error('検索インデックスシートが見つかりません');
    return;
  }
  
  try {
    // ヘッダー行以外をクリア
    const lastRow = indexSheet.getLastRow();
    if (lastRow > 1) {
      indexSheet.getRange(2, 1, lastRow - 1, indexSheet.getLastColumn()).clear();
    }
    
    const allIndexData = [];
    
    // 端末系統合ビューからインデックス作成
    const terminalView = spreadsheet.getSheetByName('integrated_view_terminal');
    if (terminalView && terminalView.getLastRow() > 1) {
      const terminalData = terminalView.getDataRange().getValues();
      const terminalHeaders = terminalData[0];
      for (let i = 1; i < terminalData.length; i++) {
        const locationNumber = getValueByColumnName(terminalData[i], terminalHeaders, '拠点管理番号');
        if (!locationNumber) continue;
        
        const searchKey = generateSearchKey(terminalData[i], locationNumber);
        allIndexData.push([
          locationNumber,
          searchKey,
          getValueByColumnName(terminalData[i], terminalHeaders, 'カテゴリ') || '', // カテゴリ
          getValueByColumnName(terminalData[i], terminalHeaders, '拠点') || '', // 拠点
          getValueByColumnName(terminalData[i], terminalHeaders, '状態') || '', // 状態
          new Date() // 最終更新日時
        ]);
      }
    }
    
    // プリンタ・その他系統合ビューからインデックス作成
    const printerView = spreadsheet.getSheetByName('integrated_view_printer_other');
    if (printerView && printerView.getLastRow() > 1) {
      const printerData = printerView.getDataRange().getValues();
      const printerHeaders = printerData[0];
      for (let i = 1; i < printerData.length; i++) {
        const locationNumber = getValueByColumnName(printerData[i], printerHeaders, '拠点管理番号');
        if (!locationNumber) continue;
        
        const searchKey = generateSearchKey(printerData[i], locationNumber);
        allIndexData.push([
          locationNumber,
          searchKey,
          getValueByColumnName(printerData[i], printerHeaders, 'カテゴリ') || '', // カテゴリ
          getValueByColumnName(printerData[i], printerHeaders, '拠点') || '', // 拠点
          getValueByColumnName(printerData[i], printerHeaders, '状態') || '', // 状態
          new Date() // 最終更新日時
        ]);
      }
    }
    
    // データを一括で書き込み
    if (allIndexData.length > 0) {
      const range = indexSheet.getRange(2, 1, allIndexData.length, allIndexData[0].length);
      range.setValues(allIndexData);
    }
    
    console.log(`検索インデックス再構築完了: ${allIndexData.length}件`);
    
  } catch (error) {
    console.error('検索インデックス再構築エラー:', error);
    throw error;
  }
}

/**
 * 統合ビューの書式設定
 */
function formatIntegratedView(sheet, viewType) {
  if (!sheet) return;
  
  try {
    // 全体の書式設定
    const dataRange = sheet.getDataRange();
    dataRange.setFontFamily('Noto Sans JP');
    dataRange.setFontSize(10);
    
    // 条件付き書式設定
    sheet.clearConditionalFormatRules();
    const rules = sheet.getConditionalFormatRules();
    
    // 状態列の条件付き書式（端末系：I列、プリンタ系：G列）
    const statusColumn = viewType === 'terminal' ? 9 : 7;
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      // 貸出中（赤色背景）
      const loanRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('貸出中')
        .setBackground('#FFE6E6')
        .setRanges([sheet.getRange(2, statusColumn, lastRow - 1, 1)])
        .build();
      rules.push(loanRule);
      
      // 返却済み（緑色背景）
      const returnRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('返却済み')
        .setBackground('#E6FFE6')
        .setRanges([sheet.getRange(2, statusColumn, lastRow - 1, 1)])
        .build();
      rules.push(returnRule);
      
      // 修理中（黄色背景）
      const repairRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('修理中')
        .setBackground('#FFFFE6')
        .setRanges([sheet.getRange(2, statusColumn, lastRow - 1, 1)])
        .build();
      rules.push(repairRule);
      
      // 要注意フラグ（90日以上貸出）
      if (viewType === 'terminal') {
        const warningRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(90)
          .setBackground('#FFB6C1')
          .setRanges([sheet.getRange(2, 12, lastRow - 1, 1)]) // L列：貸出日数
          .build();
        rules.push(warningRule);
      }
    }
    
    sheet.setConditionalFormatRules(rules);
    
    // 列幅の自動調整
    for (let col = 1; col <= sheet.getLastColumn(); col++) {
      sheet.autoResizeColumn(col);
    }
    
    // 日付列の書式設定
    const dateColumns = viewType === 'terminal' ? 
      [7, 8, 10, 11, 17, 18, 25, 26, 31, 35, 39, 43] : 
      [5, 6, 10, 15, 16, 20, 24, 28, 32, 36, 40];
    
    dateColumns.forEach(col => {
      if (col <= sheet.getLastColumn() && lastRow > 1) {
        const dateRange = sheet.getRange(2, col, lastRow - 1, 1);
        dateRange.setNumberFormat('yyyy/mm/dd');
      }
    });
    
    console.log('書式設定完了');
    
  } catch (error) {
    console.error('書式設定エラー:', error);
  }
}

/**
 * 統合ビュー手動再構築（管理者用）
 * メニューから実行可能
 */
function manualRebuildIntegratedViews() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '統合ビュー再構築',
    '全ての統合ビューを再構築します。\n処理には数分かかる場合があります。\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    try {
      rebuildAllIntegratedViews();
      ui.alert('完了', '統合ビューの再構築が完了しました。', ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('エラー', '再構築中にエラーが発生しました。\n' + error.toString(), ui.ButtonSet.OK);
    }
  }
}

/**
 * カスタムメニューの追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('管理機能')
    .addItem('スプレッドシート初期化', 'initializeSpreadsheet')
    .addItem('設定確認', 'checkSpreadsheetSetup')
    .addSeparator()
    .addItem('統合ビュー手動再構築', 'manualRebuildIntegratedViews')
    .addItem('検索インデックス再構築', 'rebuildSearchIndex')
    .addSeparator()
    .addItem('トリガー設定', 'setupTriggers')
    .addToUi();
}