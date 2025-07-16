/**
 * ビューシート機能のセットアップをテスト実行する関数
 * 実際にセットアップを行う前に、この関数で動作確認できます
 */
function testViewSheetSetup() {
  console.log('=== ビューシート機能テスト開始 ===');
  
  try {
    // 1. 現在の状態を確認
    console.log('1. 現在のデータタイプを確認...');
    const currentTypes = checkViewSheetDataTypes();
    
    if (currentTypes && currentTypes.length > 0) {
      console.log('既にビューシートデータタイプが登録されています');
      return;
    }
    
    // 2. データタイプのみ登録（実際のビューシート作成はしない）
    console.log('2. データタイプ登録テスト...');
    const dataTypeResult = initializeViewSheetDataTypes();
    console.log('データタイプ登録結果:', dataTypeResult);
    
    // 3. ビューシートの存在確認（作成はしない）
    console.log('3. ビューシート存在確認...');
    try {
      const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      
      const integratedExists = !!spreadsheet.getSheetByName('integrated_view');
      const searchExists = !!spreadsheet.getSheetByName('search_index');
      const summaryExists = !!spreadsheet.getSheetByName('summary_view');
      
      console.log('- 統合ビュー:', integratedExists ? '存在' : '未作成');
      console.log('- 検索インデックス:', searchExists ? '存在' : '未作成');
      console.log('- サマリービュー:', summaryExists ? '存在' : '未作成');
      
    } catch (error) {
      console.error('ビューシート確認エラー:', error);
    }
    
    // 4. マスタデータの状態確認
    console.log('4. マスタデータ確認...');
    const terminalMaster = getMasterData('端末マスタ');
    const printerMaster = getMasterData('プリンタマスタ');
    const otherMaster = getMasterData('その他マスタ');
    
    console.log('- 端末マスタ:', terminalMaster.length, '件');
    console.log('- プリンタマスタ:', printerMaster.length, '件');
    console.log('- その他マスタ:', otherMaster.length, '件');
    
    console.log('=== ビューシート機能テスト完了 ===');
    console.log('実際にセットアップを実行するには setupViewSheetFeatureComplete() を実行してください');
    
  } catch (error) {
    console.error('テストエラー:', error);
  }
}

/**
 * ビューシート機能を段階的にセットアップ（確認しながら実行）
 */
function setupViewSheetStepByStep() {
  console.log('=== ビューシート機能段階的セットアップ開始 ===');
  
  const ui = SpreadsheetApp.getUi();
  
  // ステップ1: データタイプ登録
  const step1Result = ui.alert(
    'ステップ1: データタイプ登録',
    'データタイプマスタにビューシート用のデータタイプを登録します。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (step1Result === ui.Button.YES) {
    const dataTypeResult = initializeViewSheetDataTypes();
    ui.alert('データタイプ登録完了', 
      `追加: ${dataTypeResult.added}件, スキップ: ${dataTypeResult.skipped}件`, 
      ui.ButtonSet.OK);
  } else {
    ui.alert('セットアップを中止しました');
    return;
  }
  
  // ステップ2: ビューシート構築
  const step2Result = ui.alert(
    'ステップ2: ビューシート構築',
    'ビューシートを作成し、初期データを構築します。この処理には時間がかかる場合があります。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (step2Result === ui.Button.YES) {
    const viewResult = initializeBuildViewSheets();
    if (viewResult.success) {
      ui.alert('ビューシート構築完了', 
        `統合ビュー: ${viewResult.integrated}行\n` +
        `検索インデックス: ${viewResult.searchIndex}行\n` +
        `サマリービュー: ${viewResult.summary}行`, 
        ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', `ビューシート構築に失敗しました: ${viewResult.error}`, ui.ButtonSet.OK);
      return;
    }
  } else {
    ui.alert('セットアップを中止しました');
    return;
  }
  
  // ステップ3: 自動更新トリガー
  const step3Result = ui.alert(
    'ステップ3: 自動更新トリガー設定',
    '5分ごとに検索インデックスを自動更新するトリガーを設定します。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (step3Result === ui.Button.YES) {
    const triggerResult = setupViewSheetUpdateTrigger();
    if (triggerResult.success) {
      ui.alert('トリガー設定完了', triggerResult.message, ui.ButtonSet.OK);
    } else {
      ui.alert('警告', `トリガー設定に失敗しました: ${triggerResult.error}\n手動更新は可能です。`, ui.ButtonSet.OK);
    }
  }
  
  ui.alert('セットアップ完了', 
    'ビューシート機能のセットアップが完了しました。\n' +
    'Webアプリケーションからデータタイプを選択して使用できます。',
    ui.ButtonSet.OK);
}

/**
 * ビューシートデータを手動で更新
 */
function updateAllViewSheets() {
  console.log('=== 全ビューシート更新開始 ===');
  
  const results = {
    integrated: null,
    searchIndex: null,
    summary: null
  };
  
  try {
    // 1. 統合ビュー更新
    console.log('統合ビュー更新中...');
    results.integrated = buildIntegratedViewData();
    console.log(`統合ビュー: ${results.integrated.success ? '成功' : '失敗'} (${results.integrated.rowsCreated || 0}行)`);
    
    // 2. 検索インデックス更新
    console.log('検索インデックス更新中...');
    results.searchIndex = rebuildSearchIndex();
    console.log(`検索インデックス: ${results.searchIndex.success ? '成功' : '失敗'} (${results.searchIndex.rowsIndexed || 0}行)`);
    
    // 3. サマリービュー更新
    console.log('サマリービュー更新中...');
    results.summary = buildSummaryViewData();
    console.log(`サマリービュー: ${results.summary.success ? '成功' : '失敗'} (${results.summary.rowsCreated || 0}行)`);
    
    console.log('=== 全ビューシート更新完了 ===');
    
    return results;
    
  } catch (error) {
    console.error('ビューシート更新エラー:', error);
    return {
      success: false,
      error: error.toString(),
      results: results
    };
  }
}

/**
 * ビューシート機能の削除（データタイプとシートを削除）
 * 注意：この関数を実行すると、ビューシート機能が完全に削除されます
 */
function removeViewSheetFeature() {
  const ui = SpreadsheetApp.getUi();
  
  const confirmResult = ui.alert(
    '警告',
    'ビューシート機能を完全に削除します。この操作は取り消せません。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResult !== ui.Button.YES) {
    ui.alert('削除を中止しました');
    return;
  }
  
  console.log('=== ビューシート機能削除開始 ===');
  
  try {
    // 1. トリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'updateSearchIndexTrigger') {
        ScriptApp.deleteTrigger(trigger);
        console.log('自動更新トリガーを削除しました');
      }
    });
    
    // 2. ビューシートを削除
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    const sheetsToDelete = ['integrated_view', 'search_index', 'summary_view'];
    sheetsToDelete.forEach(sheetName => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        spreadsheet.deleteSheet(sheet);
        console.log(`${sheetName}シートを削除しました`);
      }
    });
    
    // 3. データタイプを削除
    const viewTypeIds = ['INTEGRATED_VIEW', 'SEARCH_INDEX', 'SUMMARY_VIEW'];
    viewTypeIds.forEach(typeId => {
      // データタイプ削除関数が必要（data-type-master.gsに実装が必要）
      console.log(`データタイプ ${typeId} の削除をスキップ（手動で削除してください）`);
    });
    
    console.log('=== ビューシート機能削除完了 ===');
    ui.alert('削除完了', 'ビューシート機能を削除しました', ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('削除エラー:', error);
    ui.alert('エラー', `削除中にエラーが発生しました: ${error}`, ui.ButtonSet.OK);
  }
}