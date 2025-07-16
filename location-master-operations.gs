// ========================================
// 拠点マスタ操作関連
// ========================================

/**
 * 拠点IDから拠点名を取得（拠点マスタベース）
 * @param {string} locationId - 拠点ID
 * @return {string} 拠点名
 */
function getLocationNameById(locationId) {
  return safeExecute(() => {
    const location = getLocationById(locationId);
    return location ? location.locationName : locationId;
  }, 'getLocationNameById', locationId);
}

/**
 * 拠点IDから拠点コードを取得（拠点マスタベース）
 * @param {string} locationId - 拠点ID
 * @return {string} 拠点コード
 */
function getLocationCodeById(locationId) {
  return safeExecute(() => {
    const location = getLocationById(locationId);
    return location ? location.locationCode : locationId;
  }, 'getLocationCodeById', locationId);
}

/**
 * 全拠点の名前マッピングを取得（互換性維持用）
 * @return {Object} 拠点IDと拠点名のマッピング
 */
function getLocationNamesMapping() {
  return safeExecute(() => {
    const locations = getLocationMaster();
    const mapping = {};
    locations.forEach(location => {
      mapping[location.locationId] = location.locationName;
    });
    return mapping;
  }, 'getLocationNamesMapping', {});
}

/**
 * 拠点マスタシートを取得または作成
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 拠点マスタシート
 */
function getLocationMasterSheet() {
  const startTime = startPerformanceTimer();
  addLog('拠点マスタシート取得開始');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName('拠点マスタ');
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = spreadsheet.insertSheet('拠点マスタ');
      
      // ヘッダー行を設定
      const headers = [
        '拠点ID',          // A列
        '拠点名',          // B列
        '拠点コード',      // C列
        '管轄',            // D列
        '作成日時',        // E列
        'ステータス変更通知' // F列
      ];
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダー行のフォーマット
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4472C4');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      
      // 列幅の調整
      sheet.setColumnWidth(1, 100); // 拠点ID
      sheet.setColumnWidth(2, 150); // 拠点名
      sheet.setColumnWidth(3, 120); // 拠点コード
      sheet.setColumnWidth(4, 100); // 管轄
      sheet.setColumnWidth(5, 150); // 作成日時
      sheet.setColumnWidth(6, 150); // ステータス変更通知
      
      // 日付列のフォーマットをテキストに設定
      const dateColumn = sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1);
      dateColumn.setNumberFormat('@');
      
      addLog('拠点マスタシート作成完了');
    }
    
    endPerformanceTimer(startTime, '拠点マスタシート取得');
    return sheet;
  } catch (error) {
    endPerformanceTimer(startTime, '拠点マスタシート取得エラー');
    addLog('拠点マスタシート取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * 拠点マスタデータを取得
 * @return {Array<Object>} 拠点データの配列
 */
function getLocationMaster() {
  const startTime = startPerformanceTimer();
  addLog('拠点マスタ取得開始');
  
  try {
    const sheet = getLocationMasterSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      addLog('拠点マスタデータなし');
      return [];
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const locations = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      // 拠点IDが空の行はスキップ
      if (!row[0] || row[0] === '') continue;
      
      locations.push({
        locationId: row[0] || '',
        locationName: row[1] || '',
        locationCode: row[2] || '',
        jurisdiction: row[3] || '',
        createdDate: row[4] || '',
        statusNotification: row[5] === true || row[5] === 'TRUE' || row[5] === 'true'
      });
    }
    
    endPerformanceTimer(startTime, '拠点マスタ取得');
    addLog('拠点マスタ取得完了', { count: locations.length });
    
    return locations;
  } catch (error) {
    endPerformanceTimer(startTime, '拠点マスタ取得エラー');
    addLog('拠点マスタ取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * 拠点IDから拠点情報を取得
 * @param {string} locationId - 拠点ID
 * @return {Object|null} 拠点情報またはnull
 */
function getLocationById(locationId) {
  const startTime = startPerformanceTimer();
  addLog('拠点ID検索開始', { locationId });
  
  try {
    const locations = getLocationMaster();
    const location = locations.find(loc => loc.locationId === locationId);
    
    endPerformanceTimer(startTime, '拠点ID検索');
    addLog('拠点ID検索完了', { found: !!location });
    
    return location || null;
  } catch (error) {
    endPerformanceTimer(startTime, '拠点ID検索エラー');
    addLog('拠点ID検索エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * 新規拠点を追加
 * @param {Object} locationData - 拠点データ
 * @return {Object} 処理結果
 */
function addLocation(locationData) {
  const startTime = startPerformanceTimer();
  addLog('拠点追加開始', locationData);
  
  try {
    const sheet = getLocationMasterSheet();
    
    // 既存の拠点IDをチェック
    const existingLocations = getLocationMaster();
    const exists = existingLocations.some(loc => loc.locationId === locationData.locationId);
    
    if (exists) {
      throw new Error('この拠点IDは既に存在します');
    }
    
    // 新規行を追加
    const newRow = sheet.getLastRow() + 1;
    const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd');
    
    sheet.getRange(newRow, 1, 1, 6).setValues([[
      locationData.locationId,
      locationData.locationName,
      locationData.locationCode,
      locationData.jurisdiction || '関西',
      today,
      locationData.statusNotification || false
    ]]);
    
    // 日付列の書式をテキストに設定
    sheet.getRange(newRow, 5).setNumberFormat('@');
    
    endPerformanceTimer(startTime, '拠点追加');
    addLog('拠点追加完了');
    
    return { success: true, message: '拠点が追加されました' };
  } catch (error) {
    endPerformanceTimer(startTime, '拠点追加エラー');
    addLog('拠点追加エラー', { error: error.toString() });
    return { success: false, error: error.toString() };
  }
}

/**
 * 拠点情報を更新
 * @param {string} locationId - 更新対象の拠点ID
 * @param {Object} updateData - 更新データ
 * @return {Object} 処理結果
 */
function updateLocation(locationId, updateData) {
  const startTime = startPerformanceTimer();
  addLog('拠点更新開始', { locationId, updateData });
  
  try {
    const sheet = getLocationMasterSheet();
    const lastRow = sheet.getLastRow();
    
    // 該当行を検索
    for (let row = 2; row <= lastRow; row++) {
      const currentId = sheet.getRange(row, 1).getValue();
      if (currentId === locationId) {
        // 更新データを設定
        if (updateData.locationName !== undefined) {
          sheet.getRange(row, 2).setValue(updateData.locationName);
        }
        if (updateData.locationCode !== undefined) {
          sheet.getRange(row, 3).setValue(updateData.locationCode);
        }
        if (updateData.jurisdiction !== undefined) {
          sheet.getRange(row, 4).setValue(updateData.jurisdiction);
        }
        if (updateData.statusNotification !== undefined) {
          sheet.getRange(row, 6).setValue(updateData.statusNotification);
        }
        
        endPerformanceTimer(startTime, '拠点更新');
        addLog('拠点更新完了');
        
        return { success: true, message: '拠点が更新されました' };
      }
    }
    
    throw new Error('指定された拠点が見つかりません');
  } catch (error) {
    endPerformanceTimer(startTime, '拠点更新エラー');
    addLog('拠点更新エラー', { error: error.toString() });
    return { success: false, error: error.toString() };
  }
}

/**
 * 拠点を削除
 * @param {string} locationId - 削除対象の拠点ID
 * @return {Object} 処理結果
 */
function deleteLocation(locationId) {
  const startTime = startPerformanceTimer();
  addLog('拠点削除開始', { locationId });
  
  try {
    const sheet = getLocationMasterSheet();
    const lastRow = sheet.getLastRow();
    
    // 該当行を検索
    for (let row = 2; row <= lastRow; row++) {
      const currentId = sheet.getRange(row, 1).getValue();
      if (currentId === locationId) {
        sheet.deleteRow(row);
        
        endPerformanceTimer(startTime, '拠点削除');
        addLog('拠点削除完了');
        
        return { success: true, message: '拠点が削除されました' };
      }
    }
    
    throw new Error('指定された拠点が見つかりません');
  } catch (error) {
    endPerformanceTimer(startTime, '拠点削除エラー');
    addLog('拠点削除エラー', { error: error.toString() });
    return { success: false, error: error.toString() };
  }
}

/**
 * 既存の拠点マスタに管轄フィールドを追加する初期化関数
 * @return {Object} 処理結果
 */
function initializeLocationMasterJurisdiction() {
  const startTime = startPerformanceTimer();
  addLog('拠点マスタ管轄フィールド初期化開始');
  
  try {
    const sheet = getLocationMasterSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      addLog('データがないため初期化をスキップ');
      return { success: true, message: 'データがないため初期化をスキップしました' };
    }
    
    // ヘッダー行を確認
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hasJurisdiction = headers.includes('管轄');
    const hasStatusNotification = headers.includes('ステータス変更通知');
    
    if (hasJurisdiction && hasStatusNotification) {
      addLog('既に管轄フィールドが存在します');
      return { success: true, message: '既に管轄フィールドが存在します' };
    }
    
    // 既存データの管轄情報を設定
    for (let row = 2; row <= lastRow; row++) {
      const locationId = sheet.getRange(row, 1).getValue();
      
      // 既存拠点の管轄を設定（デフォルトは関西）
      if (!hasJurisdiction) {
        const currentJurisdiction = sheet.getRange(row, 4).getValue();
        if (!currentJurisdiction || currentJurisdiction === '') {
          sheet.getRange(row, 4).setValue('関西');
        }
      }
      
      // ステータス変更通知のデフォルト値を設定
      if (!hasStatusNotification) {
        const currentNotification = sheet.getRange(row, 6).getValue();
        if (currentNotification === '' || currentNotification === null) {
          sheet.getRange(row, 6).setValue(false);
        }
      }
    }
    
    endPerformanceTimer(startTime, '拠点マスタ管轄フィールド初期化');
    addLog('拠点マスタ管轄フィールド初期化完了');
    
    return { success: true, message: '管轄フィールドの初期化が完了しました' };
  } catch (error) {
    endPerformanceTimer(startTime, '拠点マスタ管轄フィールド初期化エラー');
    addLog('拠点マスタ管轄フィールド初期化エラー', { error: error.toString() });
    return { success: false, error: error.toString() };
  }
}

/**
 * 管轄に基づいて拠点を取得
 * @param {string} jurisdiction - 管轄名
 * @return {Array<Object>} フィルタリングされた拠点データ
 */
function getLocationsByJurisdiction(jurisdiction) {
  const startTime = startPerformanceTimer();
  addLog('管轄別拠点取得開始', { jurisdiction });
  
  try {
    const allLocations = getLocationMaster();
    
    if (!jurisdiction || jurisdiction === '') {
      // 管轄指定なしの場合は全拠点を返す
      endPerformanceTimer(startTime, '管轄別拠点取得（全件）');
      return allLocations;
    }
    
    // 指定された管轄の拠点のみをフィルタリング
    const filteredLocations = allLocations.filter(location => 
      location.jurisdiction === jurisdiction
    );
    
    endPerformanceTimer(startTime, '管轄別拠点取得');
    addLog('管轄別拠点取得完了', { 
      jurisdiction,
      totalLocations: allLocations.length,
      filteredLocations: filteredLocations.length 
    });
    
    return filteredLocations;
  } catch (error) {
    endPerformanceTimer(startTime, '管轄別拠点取得エラー');
    addLog('管轄別拠点取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * 管轄リストを取得（重複なし）
 * @return {Array<string>} 管轄名の配列（ソート済み）
 */
function getJurisdictionList() {
  const startTime = startPerformanceTimer();
  addLog('管轄リスト取得開始');
  
  try {
    const locations = getLocationMaster();
    const jurisdictions = [...new Set(locations.map(loc => loc.jurisdiction).filter(j => j && j !== ''))];
    
    endPerformanceTimer(startTime, '管轄リスト取得');
    addLog('管轄リスト取得完了', { jurisdictions });
    
    return jurisdictions.sort();
  } catch (error) {
    endPerformanceTimer(startTime, '管轄リスト取得エラー');
    addLog('管轄リスト取得エラー', { error: error.toString() });
    throw error;
  }
}

/**
 * 拠点マスタのテスト関数（デバッグ用）
 */
function testLocationMaster() {
  console.log('=== 拠点マスタテスト開始 ===');
  
  // シート取得テスト
  const sheet = getLocationMasterSheet();
  console.log('シート名:', sheet.getName());
  console.log('最終行:', sheet.getLastRow());
  console.log('最終列:', sheet.getLastColumn());
  
  // データ取得テスト
  const locations = getLocationMaster();
  console.log('拠点数:', locations.length);
  locations.forEach(location => {
    console.log('拠点:', location);
  });
  
  // 特定拠点取得テスト
  if (locations.length > 0) {
    const firstLocationId = locations[0].locationId;
    const location = getLocationById(firstLocationId);
    console.log('特定拠点取得:', location);
  }
  
  console.log('=== 拠点マスタテスト完了 ===');
}