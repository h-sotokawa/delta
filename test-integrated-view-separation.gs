/**
 * 統合ビュー分離のテスト
 * 端末系とプリンタ・その他系の統合ビューが正しく動作するかを確認
 */

function testIntegratedViewSeparation() {
  console.log('\n===== 統合ビュー分離テスト開始 =====\n');
  
  try {
    // 1. データタイプマスタの確認
    console.log('1. データタイプマスタの確認');
    const dataTypeMaster = getDataTypeMaster(true);
    
    if (dataTypeMaster.success) {
      console.log(`✅ データタイプマスタ取得成功: ${dataTypeMaster.dataTypes.length}件`);
      
      // 新しいデータタイプが存在するか確認
      const terminalView = dataTypeMaster.dataTypes.find(dt => dt.dataTypeId === 'INTEGRATED_VIEW_TERMINAL');
      const printerOtherView = dataTypeMaster.dataTypes.find(dt => dt.dataTypeId === 'INTEGRATED_VIEW_PRINTER_OTHER');
      const oldView = dataTypeMaster.dataTypes.find(dt => dt.dataTypeId === 'INTEGRATED_VIEW');
      
      console.log(`  - 端末系統合ビュー: ${terminalView ? '✅ 存在' : '❌ 見つかりません'}`);
      console.log(`  - プリンタ・その他系統合ビュー: ${printerOtherView ? '✅ 存在' : '❌ 見つかりません'}`);
      console.log(`  - 旧統合ビュー: ${oldView ? '✅ 存在（deprecated）' : '❌ 見つかりません'}`);
      
      if (terminalView) {
        console.log(`    - 端末系設定: ${JSON.stringify(terminalView.dataSourceConfig)}`);
      }
      if (printerOtherView) {
        console.log(`    - プリンタ系設定: ${JSON.stringify(printerOtherView.dataSourceConfig)}`);
      }
    } else {
      console.error('❌ データタイプマスタ取得失敗:', dataTypeMaster.error);
    }
    
    // 2. ビューシートの存在確認
    console.log('\n2. ビューシートの存在確認');
    
    try {
      const terminalSheet = getIntegratedViewTerminalSheet();
      console.log(`✅ 端末系統合ビューシート: ${terminalSheet.getName()}`);
      console.log(`  - 行数: ${terminalSheet.getLastRow()}`);
      console.log(`  - 列数: ${terminalSheet.getLastColumn()}`);
    } catch (error) {
      console.error('❌ 端末系統合ビューシート取得エラー:', error.toString());
    }
    
    try {
      const printerOtherSheet = getIntegratedViewPrinterOtherSheet();
      console.log(`✅ プリンタ・その他系統合ビューシート: ${printerOtherSheet.getName()}`);
      console.log(`  - 行数: ${printerOtherSheet.getLastRow()}`);
      console.log(`  - 列数: ${printerOtherSheet.getLastColumn()}`);
    } catch (error) {
      console.error('❌ プリンタ・その他系統合ビューシート取得エラー:', error.toString());
    }
    
    // 3. データ統合機能のテスト
    console.log('\n3. データ統合機能のテスト');
    
    const terminalUpdateResult = updateIntegratedViewTerminal();
    if (terminalUpdateResult.success) {
      console.log(`✅ 端末系統合ビュー更新成功: ${terminalUpdateResult.rowsUpdated}行`);
    } else {
      console.error('❌ 端末系統合ビュー更新失敗:', terminalUpdateResult.error);
    }
    
    const printerOtherUpdateResult = updateIntegratedViewPrinterOther();
    if (printerOtherUpdateResult.success) {
      console.log(`✅ プリンタ・その他系統合ビュー更新成功: ${printerOtherUpdateResult.rowsUpdated}行`);
    } else {
      console.error('❌ プリンタ・その他系統合ビュー更新失敗:', printerOtherUpdateResult.error);
    }
    
    // 4. 検索インデックスの確認
    console.log('\n4. 検索インデックスの確認');
    
    const searchIndexResult = rebuildSearchIndex();
    if (searchIndexResult.success) {
      console.log(`✅ 検索インデックス再構築成功: ${searchIndexResult.rowsIndexed}行`);
      console.log('  - 両方の統合ビューからデータを結合して検索インデックスを構築');
    } else {
      console.error('❌ 検索インデックス再構築失敗:', searchIndexResult.error);
    }
    
    // 5. データサンプルの確認
    console.log('\n5. データサンプルの確認');
    
    const terminalSheet = getIntegratedViewTerminalSheet();
    if (terminalSheet.getLastRow() > 1) {
      const terminalHeaders = terminalSheet.getRange(1, 1, 1, terminalSheet.getLastColumn()).getValues()[0];
      const terminalSample = terminalSheet.getRange(2, 1, Math.min(3, terminalSheet.getLastRow() - 1), terminalSheet.getLastColumn()).getValues();
      
      console.log('端末系データサンプル:');
      terminalSample.forEach((row, index) => {
        const categoryIndex = terminalHeaders.indexOf('カテゴリ');
        const modelIndex = terminalHeaders.indexOf('機種名');
        if (categoryIndex >= 0 && modelIndex >= 0) {
          console.log(`  ${index + 1}. カテゴリ: ${row[categoryIndex]}, 機種名: ${row[modelIndex]}`);
        }
      });
    }
    
    const printerOtherSheet = getIntegratedViewPrinterOtherSheet();
    if (printerOtherSheet.getLastRow() > 1) {
      const printerHeaders = printerOtherSheet.getRange(1, 1, 1, printerOtherSheet.getLastColumn()).getValues()[0];
      const printerSample = printerOtherSheet.getRange(2, 1, Math.min(3, printerOtherSheet.getLastRow() - 1), printerOtherSheet.getLastColumn()).getValues();
      
      console.log('\nプリンタ・その他系データサンプル:');
      printerSample.forEach((row, index) => {
        const categoryIndex = printerHeaders.indexOf('カテゴリ');
        const modelIndex = printerHeaders.indexOf('機種名');
        if (categoryIndex >= 0 && modelIndex >= 0) {
          console.log(`  ${index + 1}. カテゴリ: ${row[categoryIndex]}, 機種名: ${row[modelIndex]}`);
        }
      });
    }
    
    console.log('\n===== 統合ビュー分離テスト完了 =====');
    
    return {
      success: true,
      terminalViewExists: !!terminalView,
      printerOtherViewExists: !!printerOtherView,
      terminalRowsUpdated: terminalUpdateResult.success ? terminalUpdateResult.rowsUpdated : 0,
      printerOtherRowsUpdated: printerOtherUpdateResult.success ? printerOtherUpdateResult.rowsUpdated : 0
    };
    
  } catch (error) {
    console.error('\n❌ テスト実行中にエラーが発生しました:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * データタイプマスタの初期化（新しいビューを追加）
 */
function initializeNewIntegratedViews() {
  console.log('\n===== 新しい統合ビューの初期化 =====\n');
  
  try {
    // 端末系統合ビューを追加
    const terminalViewResult = addDataType({
      dataTypeId: 'INTEGRATED_VIEW_TERMINAL',
      dataTypeName: '統合ビュー（端末系）',
      description: '端末系機器（サーバー、デスクトップ、ラップトップ、タブレット）の統合ビュー表示',
      displayOrder: 5,
      dataSourceConfig: {
        sourceSheets: ['integrated_view_terminal'],
        deviceTypes: ['Server', 'Desktop', 'Laptop', 'Tablet'],
        viewType: 'integrated'
      },
      displayColumnConfig: {
        columns: ['拠点管理番号', 'カテゴリ', '機種名', '製造番号', '資産管理番号', 'ソフトウェア', 'OS', '現在ステータス', '最終更新日時', '貸出日数', '要注意フラグ']
      }
    });
    
    if (terminalViewResult.success) {
      console.log('✅ 端末系統合ビューをデータタイプマスタに追加しました');
    } else {
      console.error('❌ 端末系統合ビューの追加に失敗:', terminalViewResult.error);
    }
    
    // プリンタ・その他系統合ビューを追加
    const printerOtherViewResult = addDataType({
      dataTypeId: 'INTEGRATED_VIEW_PRINTER_OTHER',
      dataTypeName: '統合ビュー（プリンタ・その他系）',
      description: 'プリンタ・その他系機器（プリンタ、ルーター、ハブ、その他）の統合ビュー表示',
      displayOrder: 6,
      dataSourceConfig: {
        sourceSheets: ['integrated_view_printer_other'],
        deviceTypes: ['Printer', 'Router', 'Hub', 'Other'],
        viewType: 'integrated'
      },
      displayColumnConfig: {
        columns: ['拠点管理番号', 'カテゴリ', '機種名', '製造番号', '現在ステータス', '最終更新日時', '貸出日数', '要注意フラグ']
      }
    });
    
    if (printerOtherViewResult.success) {
      console.log('✅ プリンタ・その他系統合ビューをデータタイプマスタに追加しました');
    } else {
      console.error('❌ プリンタ・その他系統合ビューの追加に失敗:', printerOtherViewResult.error);
    }
    
    // 旧統合ビューをdeprecatedとしてマーク
    const deprecateResult = updateDataType('INTEGRATED_VIEW', {
      description: '全機器の統合ビュー表示（非推奨 - 端末系/プリンタ・その他系に分離されました）',
      dataSourceConfig: {
        sourceSheets: ['integrated_view'],
        deprecated: true,
        migrateTo: ['INTEGRATED_VIEW_TERMINAL', 'INTEGRATED_VIEW_PRINTER_OTHER']
      }
    });
    
    if (deprecateResult.success) {
      console.log('✅ 旧統合ビューをdeprecatedとしてマークしました');
    } else {
      console.log('⚠️ 旧統合ビューの更新に失敗（存在しない可能性があります）');
    }
    
    console.log('\n===== 初期化完了 =====');
    
  } catch (error) {
    console.error('初期化中にエラーが発生しました:', error.toString());
  }
}