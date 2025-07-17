/**
 * データタイプマスタにビューシート用のデータタイプを初期登録する関数
 * この関数を一度実行することで、統合ビュー、検索インデックス、サマリービューが使えるようになります
 */
function initializeViewSheetDataTypes() {
  console.log('=== ビューシートデータタイプ初期化開始 ===');
  
  try {
    // 追加するデータタイプの定義
    const viewSheetDataTypes = [
      {
        dataTypeId: 'INTEGRATED_VIEW',
        dataTypeName: '統合ビュー',
        targetSheet: 'INTEGRATED_VIEW',
        sheetName: 'integrated_view',
        queryType: 'all',
        description: 'マスタシートとステータス収集シートを統合した最新状態を表示',
        displayOptions: {
          showLocation: true,
          showDeviceType: false,
          showAuditSheet: false,
          showQueryType: false,
          showJurisdiction: true,
          showRefreshButton: true,
          showSearchBox: true  // 統合ビューに検索機能を統合
        },
        sortOrder: 20,
        isActive: true
      },
      {
        dataTypeId: 'SUMMARY_VIEW',
        dataTypeName: 'サマリービュー',
        targetSheet: 'SUMMARY_VIEW',
        sheetName: 'summary_view',
        queryType: 'summary',
        description: '管轄・拠点・カテゴリ別の集計データを表示',
        displayOptions: {
          showLocation: false,
          showDeviceType: false,
          showAuditSheet: false,
          showQueryType: false,
          showJurisdiction: true,
          showRefreshButton: true
        },
        sortOrder: 21,
        isActive: true
      }
    ];
    
    // 各データタイプを追加
    let addedCount = 0;
    let skippedCount = 0;
    
    for (const dataType of viewSheetDataTypes) {
      const result = addDataType(dataType);
      if (result.success) {
        addedCount++;
        console.log(`✅ 追加成功: ${dataType.dataTypeName}`);
      } else {
        skippedCount++;
        console.log(`⏭️ スキップ: ${dataType.dataTypeName} - ${result.error}`);
      }
    }
    
    console.log('=== ビューシートデータタイプ初期化完了 ===');
    console.log(`追加: ${addedCount}件, スキップ: ${skippedCount}件`);
    console.log('注意: 検索インデックスは内部使用のため、データタイプには登録されません');
    
    return {
      success: true,
      added: addedCount,
      skipped: skippedCount
    };
    
  } catch (error) {
    console.error('ビューシートデータタイプ初期化エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 統合ビューと検索インデックスの初期構築を実行
 * データタイプ登録後に実行することで、実際のビューシートにデータが入ります
 */
function initializeBuildViewSheets() {
  console.log('=== ビューシート初期構築開始 ===');
  
  try {
    // 1. ビューシートの作成（存在しない場合）
    console.log('1. ビューシートの作成確認...');
    const integratedSheet = getIntegratedViewSheet();
    const searchSheet = getSearchIndexSheet();
    const summarySheet = getSummaryViewSheet();
    
    console.log('✅ ビューシート確認完了');
    console.log(`  - 統合ビュー: ${integratedSheet.getName()}`);
    console.log(`  - 検索インデックス: ${searchSheet.getName()}`);
    console.log(`  - サマリービュー: ${summarySheet.getName()}`);
    
    // 2. 統合ビューの構築
    console.log('2. 統合ビューデータ構築中...');
    const viewResult = buildIntegratedViewData();
    
    if (viewResult.success) {
      console.log(`✅ 統合ビュー構築成功: ${viewResult.rowsCreated}行`);
    } else {
      console.error('❌ 統合ビュー構築失敗:', viewResult.error);
      return viewResult;
    }
    
    // 3. 検索インデックスの構築
    console.log('3. 検索インデックス構築中...');
    const indexResult = rebuildSearchIndex();
    
    if (indexResult.success) {
      console.log(`✅ 検索インデックス構築成功: ${indexResult.rowsIndexed}行`);
    } else {
      console.error('❌ 検索インデックス構築失敗:', indexResult.error);
      return indexResult;
    }
    
    // 4. サマリービューの構築
    console.log('4. サマリービュー構築中...');
    const summaryResult = buildSummaryViewData();
    
    if (summaryResult.success) {
      console.log(`✅ サマリービュー構築成功: ${summaryResult.rowsCreated}行`);
    } else {
      console.error('❌ サマリービュー構築失敗:', summaryResult.error);
      return summaryResult;
    }
    
    console.log('=== ビューシート初期構築完了 ===');
    
    return {
      success: true,
      integrated: viewResult.rowsCreated,
      searchIndex: indexResult.rowsIndexed,
      summary: summaryResult.rowsCreated
    };
    
  } catch (error) {
    console.error('ビューシート初期構築エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 定期更新トリガーを設定
 * 5分ごとに検索インデックスを更新するトリガーを作成します
 */
function setupViewSheetUpdateTrigger() {
  console.log('=== ビューシート更新トリガー設定開始 ===');
  
  try {
    // 既存のトリガーを確認
    const triggers = ScriptApp.getProjectTriggers();
    const existingTrigger = triggers.find(trigger => 
      trigger.getHandlerFunction() === 'updateSearchIndexTrigger'
    );
    
    if (existingTrigger) {
      console.log('既存のトリガーが見つかりました。削除して再作成します。');
      ScriptApp.deleteTrigger(existingTrigger);
    }
    
    // 新しいトリガーを作成（5分ごと）
    ScriptApp.newTrigger('updateSearchIndexTrigger')
      .timeBased()
      .everyMinutes(5)
      .create();
    
    console.log('✅ 更新トリガー設定完了（5分ごと）');
    
    return {
      success: true,
      message: '検索インデックス更新トリガーを設定しました（5分ごと）'
    };
    
  } catch (error) {
    console.error('トリガー設定エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ビューシート機能の完全セットアップ
 * この関数を実行すると、データタイプ登録、ビューシート作成、初期データ構築、トリガー設定まで全て行います
 */
function setupViewSheetFeatureComplete() {
  console.log('=== ビューシート機能完全セットアップ開始 ===');
  
  const results = {
    dataTypes: null,
    viewSheets: null,
    trigger: null
  };
  
  try {
    // 1. データタイプ登録
    console.log('ステップ1: データタイプ登録...');
    results.dataTypes = initializeViewSheetDataTypes();
    
    if (!results.dataTypes.success) {
      throw new Error('データタイプ登録に失敗しました');
    }
    
    // 2. ビューシート構築
    console.log('ステップ2: ビューシート構築...');
    results.viewSheets = initializeBuildViewSheets();
    
    if (!results.viewSheets.success) {
      throw new Error('ビューシート構築に失敗しました');
    }
    
    // 3. トリガー設定
    console.log('ステップ3: 自動更新トリガー設定...');
    results.trigger = setupViewSheetUpdateTrigger();
    
    if (!results.trigger.success) {
      console.warn('トリガー設定に失敗しましたが、手動更新は可能です');
    }
    
    console.log('=== ビューシート機能完全セットアップ完了 ===');
    console.log('セットアップ結果:');
    console.log(`- データタイプ: 追加${results.dataTypes.added}件, スキップ${results.dataTypes.skipped}件`);
    console.log(`- 統合ビュー: ${results.viewSheets.integrated}行`);
    console.log(`- 検索インデックス: ${results.viewSheets.searchIndex}行`);
    console.log(`- サマリービュー: ${results.viewSheets.summary}行`);
    console.log(`- 自動更新: ${results.trigger.success ? '有効' : '無効'}`);
    
    return {
      success: true,
      results: results
    };
    
  } catch (error) {
    console.error('セットアップエラー:', error);
    return {
      success: false,
      error: error.toString(),
      results: results
    };
  }
}

/**
 * ビューシートデータタイプの状態を確認
 */
function checkViewSheetDataTypes() {
  console.log('=== ビューシートデータタイプ確認 ===');
  
  const response = getDataTypeMaster(true);
  if (!response.success) {
    console.error('データタイプマスタ取得失敗:', response.error);
    return;
  }
  
  const viewSheetTypes = response.dataTypes.filter(dt => 
    ['INTEGRATED_VIEW', 'SEARCH_INDEX', 'SUMMARY_VIEW'].includes(dt.dataTypeId)
  );
  
  console.log(`ビューシートデータタイプ: ${viewSheetTypes.length}件`);
  viewSheetTypes.forEach(dt => {
    console.log(`- ${dt.dataTypeName} (${dt.dataTypeId}): ${dt.isActive ? '有効' : '無効'}`);
  });
  
  return viewSheetTypes;
}