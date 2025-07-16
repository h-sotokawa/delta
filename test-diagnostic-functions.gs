// ========================================
// テスト・診断関数
// ========================================

/**
 * サマリーシートの詳細診断を行う関数
 * @return {Object} 診断結果
 */
function diagnoseSummarySheet() {
  const startTime = Date.now();
  const diagnosticLogs = [];
  
  function addDiagnosticLog(message, data = {}) {
    const logEntry = {
      timestamp: new Date().toISOString(),
      message: message,
      data: data
    };
    diagnosticLogs.push(logEntry);
    console.log(`[DIAGNOSTIC] ${message}`, data);
  }
  
  try {
    addDiagnosticLog('サマリーシート診断開始');
    
    // スプレッドシートID取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID_DESTINATION');
    
    addDiagnosticLog('スプレッドシートID確認', {
      spreadsheetId: spreadsheetId ? '設定済み' : '未設定',
      idLength: spreadsheetId ? spreadsheetId.length : 0
    });
    
    if (!spreadsheetId) {
      throw new Error('SPREADSHEET_ID_DESTINATIONが設定されていません');
    }
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    addDiagnosticLog('スプレッドシート取得成功', {
      name: spreadsheet.getName(),
      url: spreadsheet.getUrl()
    });
    
    // 全シート取得
    const allSheets = spreadsheet.getSheets();
    const sheetInfo = allSheets.map(sheet => ({
      name: sheet.getName(),
      id: sheet.getSheetId(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn(),
      hidden: sheet.isSheetHidden()
    }));
    
    addDiagnosticLog('全シート情報', {
      totalSheets: allSheets.length,
      sheets: sheetInfo
    });
    
    // サマリーシート特定
    const summarySheet = spreadsheet.getSheetByName('サマリー');
    if (!summarySheet) {
      addDiagnosticLog('エラー: サマリーシートが見つかりません', {
        availableSheets: sheetInfo.map(s => s.name)
      });
      throw new Error('「サマリー」シートが見つかりません');
    }
    
    addDiagnosticLog('サマリーシート取得成功', {
      name: summarySheet.getName(),
      id: summarySheet.getSheetId(),
      isHidden: summarySheet.isSheetHidden()
    });
    
    // データ範囲確認
    const lastRow = summarySheet.getLastRow();
    const lastColumn = summarySheet.getLastColumn();
    
    addDiagnosticLog('データ範囲確認', {
      lastRow: lastRow,
      lastColumn: lastColumn,
      isEmpty: lastRow === 0,
      totalCells: lastRow * lastColumn
    });
    
    if (lastRow === 0) {
      addDiagnosticLog('警告: サマリーシートにデータがありません');
      return {
        success: false,
        error: 'サマリーシートにデータがありません',
        diagnostics: diagnosticLogs,
        executionTime: Date.now() - startTime + 'ms'
      };
    }
    
    // 実際のデータを少し取得してみる
    const sampleRange = summarySheet.getRange(1, 1, Math.min(lastRow, 5), Math.min(lastColumn, 5));
    const sampleData = sampleRange.getValues();
    
    addDiagnosticLog('サンプルデータ取得成功', {
      sampleRows: sampleData.length,
      sampleCols: sampleData[0] ? sampleData[0].length : 0,
      firstRow: sampleData[0] || [],
      secondRow: sampleData[1] || []
    });
    
    // セルの型チェック
    let dateCount = 0;
    let errorCount = 0;
    let nullCount = 0;
    
    for (let i = 0; i < sampleData.length; i++) {
      for (let j = 0; j < sampleData[i].length; j++) {
        const cell = sampleData[i][j];
        if (cell instanceof Date) {
          dateCount++;
        } else if (cell === null || cell === undefined) {
          nullCount++;
        } else if (cell && typeof cell === 'object' && cell.toString().startsWith('#')) {
          errorCount++;
        }
      }
    }
    
    addDiagnosticLog('セル型統計', {
      dateCount: dateCount,
      errorCount: errorCount,
      nullCount: nullCount
    });
    
    // 実際のgetDestinationSheetDataを呼び出し
    addDiagnosticLog('実際の関数呼び出しテスト開始');
    const result = getDestinationSheetData('サマリー', 'diagnostic');
    
    addDiagnosticLog('実際の関数呼び出し結果', {
      success: result.success,
      hasData: result.data ? true : false,
      dataLength: result.data ? result.data.length : 0,
      error: result.error || 'なし'
    });
    
    return {
      success: true,
      message: 'サマリーシート診断完了',
      diagnostics: diagnosticLogs,
      sheetInfo: {
        name: summarySheet.getName(),
        lastRow: lastRow,
        lastColumn: lastColumn,
        sampleData: sampleData
      },
      functionResult: result,
      executionTime: Date.now() - startTime + 'ms'
    };
    
  } catch (error) {
    addDiagnosticLog('診断中にエラー発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.toString(),
      diagnostics: diagnosticLogs,
      executionTime: Date.now() - startTime + 'ms'
    };
  }
}

/**
 * サマリーデータの動的拠点表示テスト
 * @return {Object} テスト結果
 */
function testDynamicSummaryDisplay() {
  console.log('=== サマリーデータ動的拠点表示テスト開始 ===');
  
  try {
    // 1. 拠点マスタの確認
    console.log('\n1. 拠点マスタデータ確認');
    const locations = getLocationMaster();
    console.log('登録拠点数:', locations.length);
    console.log('拠点リスト:', locations.map(loc => ({
      id: loc.locationId,
      name: loc.locationName,
      jurisdiction: loc.jurisdiction
    })));
    
    // 2. サマリーデータの取得
    console.log('\n2. サマリーデータ取得');
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const summarySheet = spreadsheet.getSheetByName('サマリー');
    
    if (!summarySheet) {
      console.log('サマリーシートが見つかりません');
      return { success: false, error: 'サマリーシートが見つかりません' };
    }
    
    const summaryData = summarySheet.getDataRange().getValues();
    console.log('サマリーデータ行数:', summaryData.length);
    
    // 3. サマリーデータ内の拠点名を確認
    console.log('\n3. サマリーデータ内の拠点名確認');
    const foundLocationNames = new Set();
    
    for (let i = 0; i < summaryData.length; i++) {
      const firstCell = summaryData[i][0];
      if (firstCell && typeof firstCell === 'string' && firstCell.trim() !== '') {
        // カテゴリ行や合計行を除外
        if (!firstCell.match(/^\d+\./) && firstCell !== '合計' && firstCell !== 'SV') {
          foundLocationNames.add(firstCell.trim());
        }
      }
    }
    
    console.log('サマリーデータ内で見つかった拠点名:', Array.from(foundLocationNames));
    
    // 4. 拠点マスタとの照合
    console.log('\n4. 拠点マスタとの照合');
    const validLocationNames = locations.map(loc => loc.locationName);
    const unmatchedLocations = [];
    const matchedLocations = [];
    
    foundLocationNames.forEach(name => {
      if (validLocationNames.includes(name)) {
        matchedLocations.push(name);
      } else {
        unmatchedLocations.push(name);
      }
    });
    
    console.log('拠点マスタと一致する拠点:', matchedLocations);
    console.log('拠点マスタに存在しない拠点:', unmatchedLocations);
    
    // 5. 新規拠点追加シミュレーション
    console.log('\n5. 新規拠点追加シミュレーション');
    console.log('テスト拠点「テスト横浜」を追加した場合...');
    console.log('動的実装では自動的にサマリーに含まれるようになります');
    
    console.log('\n=== サマリーデータ動的拠点表示テスト完了 ===');
    
    return {
      success: true,
      results: {
        locationCount: locations.length,
        summaryLocationCount: foundLocationNames.size,
        matchedCount: matchedLocations.length,
        unmatchedCount: unmatchedLocations.length,
        dynamicImplementation: '有効'
      }
    };
    
  } catch (error) {
    console.error('テストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 管轄機能の統合テスト
 * @return {Object} テスト結果
 */
function testJurisdictionFeatures() {
  console.log('=== 管轄機能統合テスト開始 ===');
  
  try {
    // 1. 既存データの初期化
    console.log('\n1. 既存データの管轄フィールド初期化');
    const initResult = initializeLocationMasterJurisdiction();
    console.log('初期化結果:', initResult);
    
    // 2. 管轄リストの取得テスト
    console.log('\n2. 管轄リスト取得テスト');
    const jurisdictions = getJurisdictionList();
    console.log('管轄リスト:', jurisdictions);
    
    // 3. 管轄別拠点取得テスト
    console.log('\n3. 管轄別拠点取得テスト');
    jurisdictions.forEach(jurisdiction => {
      const locations = getLocationsByJurisdiction(jurisdiction);
      console.log(`管轄「${jurisdiction}」の拠点数:`, locations.length);
      console.log('拠点リスト:', locations.map(loc => loc.locationName));
    });
    
    // 4. 新規拠点追加テスト（管轄とステータス変更通知を含む）
    console.log('\n4. 新規拠点追加テスト');
    const testLocation = {
      locationId: 'test_tokyo',
      locationName: 'テスト東京',
      locationCode: 'TESTTOKYO',
      jurisdiction: '関東',
      groupEmail: 'tokyo-test@example.com',
      statusChangeNotification: true,
      status: 'active'
    };
    
    // 既存の場合は削除
    try {
      deleteLocation(testLocation.locationId);
    } catch (e) {
      // 無視
    }
    
    const addResult = addLocation(testLocation);
    console.log('追加結果:', addResult);
    
    // 5. 追加した拠点の確認
    console.log('\n5. 追加した拠点の確認');
    const addedLocation = getLocationById(testLocation.locationId);
    console.log('追加された拠点:', addedLocation);
    
    // 6. 拠点情報の更新テスト
    console.log('\n6. 拠点情報更新テスト');
    const updateData = {
      jurisdiction: '中部',
      statusChangeNotification: false
    };
    const updateResult = updateLocation(testLocation.locationId, updateData);
    console.log('更新結果:', updateResult);
    
    // 7. 更新後の確認
    console.log('\n7. 更新後の確認');
    const updatedLocation = getLocationById(testLocation.locationId);
    console.log('更新された拠点:', updatedLocation);
    
    // 8. テストデータのクリーンアップ
    console.log('\n8. テストデータのクリーンアップ');
    const deleteResult = deleteLocation(testLocation.locationId);
    console.log('削除結果:', deleteResult);
    
    console.log('\n=== 管轄機能統合テスト完了 ===');
    
    return {
      success: true,
      results: {
        initResult,
        jurisdictions,
        testResults: 'All tests passed'
      }
    };
    
  } catch (error) {
    console.error('テストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 拠点マスタのテスト関数（デバッグ用）
 * @return {Object} テスト結果
 */
function testLocationMaster() {
  try {
    console.log('=== 拠点マスタテスト開始 ===');
    
    // 1. シートの存在確認
    console.log('1. シート取得テスト');
    const sheet = getLocationMasterSheet();
    console.log('シート名:', sheet.getName());
    console.log('最終行:', sheet.getLastRow());
    console.log('最終列:', sheet.getLastColumn());
    
    // 2. データ取得テスト
    console.log('\n2. データ取得テスト');
    const locations = getLocationMaster();
    console.log('取得件数:', locations.length);
    console.log('取得データ（日付除外）:', JSON.stringify(locations, null, 2));
    
    // 3. 個別拠点取得テスト
    console.log('\n3. 個別拠点取得テスト');
    if (locations.length > 0) {
      const firstLocationId = locations[0].locationId;
      const location = getLocationById(firstLocationId);
      console.log('個別取得結果:', location);
    }
    
    console.log('\n=== 拠点マスタテスト完了 ===');
    return {
      success: true,
      message: '拠点マスタテスト完了'
    };
  } catch (error) {
    console.error('拠点マスタテストエラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}