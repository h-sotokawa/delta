// 統合ビューのデータ問題を修正する関数
function fixIntegratedViewData() {
  console.log('=== 統合ビューデータ修正開始 ===');
  
  try {
    // 1. 現在の統合ビューの状態を確認
    console.log('\n1. 現在の統合ビュー状態確認:');
    const terminalSheet = getIntegratedViewTerminalSheet();
    const currentData = terminalSheet.getDataRange().getValues();
    console.log('  現在の行数:', currentData.length - 1);
    
    // 2. 統合ビューを強制的に再構築
    console.log('\n2. 統合ビューを強制再構築中...');
    const rebuildResult = rebuildTerminalIntegratedView();
    console.log('  再構築結果:', rebuildResult);
    
    // 3. 再構築後のデータを確認
    console.log('\n3. 再構築後の統合ビュー確認:');
    const updatedData = terminalSheet.getDataRange().getValues();
    const headers = updatedData[0];
    console.log('  更新後の行数:', updatedData.length - 1);
    
    // 拠点管理番号のインデックスを取得
    const managementNumberIndex = headers.indexOf('拠点管理番号');
    const locationNameIndex = headers.indexOf('拠点名');
    
    // 拠点別のデータ数を集計
    const locationStats = {};
    for (let i = 1; i < updatedData.length; i++) {
      const managementNumber = updatedData[i][managementNumberIndex];
      const locationName = updatedData[i][locationNameIndex];
      
      if (managementNumber) {
        const locationCode = managementNumber.split('_')[0].toUpperCase();
        if (!locationStats[locationCode]) {
          locationStats[locationCode] = {
            count: 0,
            locationName: locationName,
            samples: []
          };
        }
        locationStats[locationCode].count++;
        if (locationStats[locationCode].samples.length < 2) {
          locationStats[locationCode].samples.push(managementNumber);
        }
      }
    }
    
    console.log('\n  拠点別データ数:');
    Object.entries(locationStats).forEach(([code, stats]) => {
      console.log(`    ${code} (${stats.locationName}): ${stats.count}件`);
      stats.samples.forEach(sample => {
        console.log(`      - ${sample}`);
      });
    });
    
    // 4. 各拠点でフィルタリングテスト
    console.log('\n4. 拠点フィルタリングテスト:');
    const testLocations = ['all', 'osaka', 'kobe', 'himeji', 'kyoto', 'tokyo'];
    
    testLocations.forEach(location => {
      const result = getSpreadsheetData(location, 'all', 'INTEGRATED_VIEW_TERMINAL');
      console.log(`  ${location}: ${result.success ? (result.data.length - 1) + '件' : 'エラー: ' + result.error}`);
    });
    
    // 5. 東京のデータがない場合、テストデータを追加
    if (!locationStats['TOKYO'] || locationStats['TOKYO'].count === 0) {
      console.log('\n5. 東京のテストデータを追加:');
      addTokyoTestData();
      
      // 再度統合ビューを更新
      console.log('  統合ビューを再更新中...');
      const updateResult = updateIntegratedViewTerminal();
      console.log('  更新結果:', updateResult);
      
      // 東京データを再確認
      const tokyoResult = getSpreadsheetData('tokyo', 'all', 'INTEGRATED_VIEW_TERMINAL');
      console.log('  東京データ再確認:', tokyoResult.success ? (tokyoResult.data.length - 1) + '件' : 'エラー');
    }
    
  } catch (error) {
    console.error('修正エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

// 東京のテストデータを追加する関数
function addTokyoTestData() {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // 端末マスタシートにテストデータを追加
    const terminalSheet = spreadsheet.getSheetByName('端末マスタ');
    if (!terminalSheet) {
      console.error('端末マスタシートが見つかりません');
      return;
    }
    
    // テストデータ
    const testData = [
      ['TOKYO_desktop_ThinkPad_TK001_001', 'desktop', 'ThinkPad X1 Carbon', 'TK001', 'ASSET_TK001', 'Windows 11 Pro', 'Windows', formatDateFast(new Date()), 'MGT_TK001', 'https://forms.gle/terminal_tk001', 'https://qr.codes/terminal_tk001'],
      ['TOKYO_laptop_MacBook_TK002_001', 'laptop', 'MacBook Pro 14', 'TK002', 'ASSET_TK002', 'macOS Sonoma', 'macOS', formatDateFast(new Date()), 'MGT_TK002', 'https://forms.gle/terminal_tk002', 'https://qr.codes/terminal_tk002'],
      ['TOKYO_server_PowerEdge_TK003_001', 'server', 'Dell PowerEdge R750', 'TK003', 'ASSET_TK003', 'Windows Server 2022', 'Windows Server', formatDateFast(new Date()), 'MGT_TK003', 'https://forms.gle/terminal_tk003', 'https://qr.codes/terminal_tk003']
    ];
    
    // データを追加
    const lastRow = terminalSheet.getLastRow();
    testData.forEach((row, index) => {
      terminalSheet.getRange(lastRow + index + 1, 1, 1, row.length).setValues([row]);
    });
    
    console.log(`  ${testData.length}件の東京テストデータを追加しました`);
    
    // 拠点マスタに東京が存在するか確認
    const locationSheet = spreadsheet.getSheetByName('拠点マスタ');
    if (locationSheet) {
      const locationData = locationSheet.getDataRange().getValues();
      const headers = locationData[0];
      const codeIndex = headers.indexOf('拠点コード');
      
      let tokyoExists = false;
      for (let i = 1; i < locationData.length; i++) {
        if (locationData[i][codeIndex] === 'TOKYO') {
          tokyoExists = true;
          break;
        }
      }
      
      if (!tokyoExists) {
        // 東京を拠点マスタに追加
        const newLocation = ['TOKYO', '東京支店', '関東', true];
        locationSheet.getRange(locationSheet.getLastRow() + 1, 1, 1, newLocation.length).setValues([newLocation]);
        console.log('  拠点マスタに東京を追加しました');
      }
    }
    
    // ステータス収集シートにもテストデータを追加
    const statusSheet = spreadsheet.getSheetByName('端末ステータス収集');
    if (statusSheet) {
      const statusHeaders = statusSheet.getRange(1, 1, 1, statusSheet.getLastColumn()).getValues()[0];
      const statusData = [
        formatDateFast(new Date()),  // タイムスタンプ
        'MGT_TK001',  // 9999.管理ID
        'TOKYO_desktop_ThinkPad_TK001_001',  // 0-0.拠点管理番号
        '山田太郎',  // 0-1.担当者
        'はい',  // 0-2.EMシステムズの社員ですか？
        'EMシステムズ',  // 0-3.所属会社
        '3.社内にて保管中',  // 0-4.ステータス
        '',  // 1-1.顧客名または貸出先
        '',  // 1-2.顧客番号
        '',  // 1-3.住所
        '',  // 1-4.ユーザー機の預り有無
        '',  // 1-5.依頼者
        '',  // 1-6.備考
        '',  // 1-7.預りユーザー機のシリアルNo.
        '',  // 1-8.お預かり証No.
        '',  // 2-1.預り機返却の有無
        '',  // 2-2.依頼者
        '',  // 2-3.備考
        '社内保管',  // 3-0.社内ステータス
        '',  // 3-0-1.棚卸しフラグ
        'TOKYO',  // 3-0-2.拠点
        '',  // 3-1-1.ソフト
        '',  // 3-1-2.備考
        '',  // 3-2-1.端末初期化の引継ぎ
        '',  // 3-2-2.備考
        '',  // 3-2-3.引継ぎ担当者
        '',  // 3-2-4.初期化作業の引継ぎ
        '',  // 4-1.所在
        '',  // 4-2.持ち出し理由
        '',  // 4-3.備考
        '',  // 5-1.内容
        '',  // 5-2.所在
        ''   // 5-3.備考
      ];
      
      statusSheet.getRange(statusSheet.getLastRow() + 1, 1, 1, statusData.length).setValues([statusData]);
      console.log('  ステータス収集シートにテストデータを追加しました');
    }
    
  } catch (error) {
    console.error('テストデータ追加エラー:', error);
  }
}

// 統合ビューを完全に再構築する関数
function rebuildTerminalIntegratedView() {
  try {
    const sheet = getIntegratedViewTerminalSheet();
    
    // 既存データをクリア（ヘッダーは残す）
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    
    // データを再構築
    const terminalData = getTerminalMasterData();
    const statusData = getLatestStatusCollectionData();
    const locationMasterData = getLocationMasterData();
    
    const integratedData = integrateDeviceDataForView(terminalData, statusData, locationMasterData, 'terminal');
    
    if (integratedData.length > 0) {
      sheet.getRange(2, 1, integratedData.length, integratedData[0].length).setValues(integratedData);
    }
    
    return {
      success: true,
      rowsCreated: integratedData.length
    };
    
  } catch (error) {
    console.error('再構築エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}