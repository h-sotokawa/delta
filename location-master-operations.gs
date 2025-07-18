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
        'グループメール',  // E列
        '作成日時',        // F列
        'ステータス変更通知', // G列
        'ステータス'      // H列
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
      sheet.setColumnWidth(5, 200); // グループメール
      sheet.setColumnWidth(6, 150); // 作成日時
      sheet.setColumnWidth(7, 150); // ステータス変更通知
      sheet.setColumnWidth(8, 100); // ステータス
      
      // 日付列のフォーマットをテキストに設定
      const dateColumn = sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1);
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
    
    // ヘッダー行を取得して列のインデックスを動的に特定
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnMap = {};
    headers.forEach((header, index) => {
      columnMap[header] = index;
    });
    
    // 列インデックスを取得（存在しない場合は-1）
    const idCol = columnMap['拠点ID'] !== undefined ? columnMap['拠点ID'] : -1;
    const nameCol = columnMap['拠点名'] !== undefined ? columnMap['拠点名'] : -1;
    const codeCol = columnMap['拠点コード'] !== undefined ? columnMap['拠点コード'] : -1;
    const jurisdictionCol = columnMap['管轄'] !== undefined ? columnMap['管轄'] : -1;
    const groupEmailCol = columnMap['グループメール'] !== undefined ? columnMap['グループメール'] : columnMap['グループメールアドレス'] !== undefined ? columnMap['グループメールアドレス'] : -1;
    const createdDateCol = columnMap['作成日時'] !== undefined ? columnMap['作成日時'] : columnMap['作成日'] !== undefined ? columnMap['作成日'] : -1;
    const statusNotificationCol = columnMap['ステータス変更通知'] !== undefined ? columnMap['ステータス変更通知'] : -1;
    const statusCol = columnMap['ステータス'] !== undefined ? columnMap['ステータス'] : -1;
    
    // 必須列の確認
    if (idCol === -1 || nameCol === -1) {
      throw new Error('拠点マスタシートの必須列（拠点ID、拠点名）が見つかりません');
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const locations = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      // 拠点IDが空の行はスキップ
      const locationId = getValueByColumnName(row, headers, '拠点ID');
      if (!locationId || locationId === '') continue;
      
      locations.push({
        locationId: locationId,
        locationName: getValueByColumnName(row, headers, '拠点名'),
        locationCode: getValueByColumnName(row, headers, '拠点コード'),
        jurisdiction: getValueByColumnName(row, headers, '管轄'),
        groupEmail: getValueByColumnName(row, headers, 'グループメール') || getValueByColumnName(row, headers, 'グループメールアドレス'),
        createdDate: getValueByColumnName(row, headers, '作成日時') || getValueByColumnName(row, headers, '作成日'),
        statusNotification: statusNotificationCol >= 0 ? (row[statusNotificationCol] === true || row[statusNotificationCol] === 'TRUE' || row[statusNotificationCol] === 'true') : false,
        status: getValueByColumnName(row, headers, 'ステータス') || 'active'
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
    
    // ヘッダー行を取得して列のインデックスを動的に特定
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnMap = {};
    headers.forEach((header, index) => {
      columnMap[header] = index + 1; // 1ベースの列番号
    });
    
    // 新規行を追加
    const newRow = sheet.getLastRow() + 1;
    const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd');
    
    // 各列に値を設定
    if (columnMap['拠点ID']) {
      sheet.getRange(newRow, columnMap['拠点ID']).setValue(locationData.locationId);
    }
    if (columnMap['拠点名']) {
      sheet.getRange(newRow, columnMap['拠点名']).setValue(locationData.locationName);
    }
    if (columnMap['拠点コード']) {
      sheet.getRange(newRow, columnMap['拠点コード']).setValue(locationData.locationCode || '');
    }
    if (columnMap['管轄']) {
      sheet.getRange(newRow, columnMap['管轄']).setValue(locationData.jurisdiction || '関西');
    }
    if (columnMap['グループメール'] || columnMap['グループメールアドレス']) {
      const emailCol = columnMap['グループメール'] || columnMap['グループメールアドレス'];
      sheet.getRange(newRow, emailCol).setValue(locationData.groupEmail || '');
    }
    if (columnMap['作成日時'] || columnMap['作成日']) {
      const dateCol = columnMap['作成日時'] || columnMap['作成日'];
      sheet.getRange(newRow, dateCol).setValue(today);
      // 日付列の書式をテキストに設定
      sheet.getRange(newRow, dateCol).setNumberFormat('@');
    }
    if (columnMap['ステータス変更通知']) {
      sheet.getRange(newRow, columnMap['ステータス変更通知']).setValue(locationData.statusNotification || false);
    }
    if (columnMap['ステータス']) {
      sheet.getRange(newRow, columnMap['ステータス']).setValue(locationData.status || 'active');
    }
    
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
    
    // ヘッダー行を取得して列のインデックスを動的に特定
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnMap = {};
    headers.forEach((header, index) => {
      columnMap[header] = index + 1; // 1ベースの列番号
    });
    
    // 拠点ID列を特定
    const idCol = columnMap['拠点ID'];
    if (!idCol) {
      throw new Error('拠点ID列が見つかりません');
    }
    
    // 該当行を検索
    for (let row = 2; row <= lastRow; row++) {
      const currentId = sheet.getRange(row, idCol).getValue();
      if (currentId === locationId) {
        // 更新データを設定
        if (updateData.locationName !== undefined && columnMap['拠点名']) {
          sheet.getRange(row, columnMap['拠点名']).setValue(updateData.locationName);
        }
        if (updateData.locationCode !== undefined && columnMap['拠点コード']) {
          sheet.getRange(row, columnMap['拠点コード']).setValue(updateData.locationCode);
        }
        if (updateData.jurisdiction !== undefined && columnMap['管轄']) {
          sheet.getRange(row, columnMap['管轄']).setValue(updateData.jurisdiction);
        }
        if (updateData.groupEmail !== undefined && (columnMap['グループメール'] || columnMap['グループメールアドレス'])) {
          const emailCol = columnMap['グループメール'] || columnMap['グループメールアドレス'];
          sheet.getRange(row, emailCol).setValue(updateData.groupEmail);
        }
        if (updateData.statusNotification !== undefined && columnMap['ステータス変更通知']) {
          sheet.getRange(row, columnMap['ステータス変更通知']).setValue(updateData.statusNotification);
        }
        if (updateData.status !== undefined && columnMap['ステータス']) {
          sheet.getRange(row, columnMap['ステータス']).setValue(updateData.status);
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
    
    // ヘッダー行を取得して拠点ID列を特定
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idCol = getColumnNumber(headers, '拠点ID');
    
    if (idCol === 0) {
      throw new Error('拠点ID列が見つかりません');
    }
    
    // 該当行を検索
    for (let row = 2; row <= lastRow; row++) {
      const currentId = sheet.getRange(row, idCol).getValue();
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
    
    // 列インデックスを取得
    const jurisdictionCol = headers.indexOf('管轄') + 1;
    const statusNotificationCol = headers.indexOf('ステータス変更通知') + 1;
    
    // 既存データの管轄情報を設定
    for (let row = 2; row <= lastRow; row++) {
      // 既存拠点の管轄を設定（デフォルトは関西）
      if (!hasJurisdiction && jurisdictionCol > 0) {
        const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
        const currentJurisdiction = getValueByColumnName(rowData, headers, '管轄');
        if (!currentJurisdiction || currentJurisdiction === '') {
          sheet.getRange(row, jurisdictionCol).setValue('関西');
        }
      }
      
      // ステータス変更通知のデフォルト値を設定
      if (!hasStatusNotification && statusNotificationCol > 0) {
        const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
        const currentNotification = getValueByColumnName(rowData, headers, 'ステータス変更通知');
        if (currentNotification === '' || currentNotification === null) {
          sheet.getRange(row, statusNotificationCol).setValue(false);
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