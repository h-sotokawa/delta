// ========================================
// 検索インデックス構築関連
// 統合ビューから検索用インデックスを構築
// ========================================

/**
 * 検索インデックスを再構築
 * @return {Object} 構築結果
 */
function rebuildSearchIndex() {
  const startTime = startPerformanceTimer();
  addLog('検索インデックス再構築開始');
  
  try {
    // 両方の統合ビューシートからデータを取得
    const terminalSheet = getIntegratedViewTerminalSheet();
    const printerOtherSheet = getIntegratedViewPrinterOtherSheet();
    const searchSheet = getSearchIndexSheet();
    
    // 端末系データを取得
    const terminalLastRow = terminalSheet.getLastRow();
    let terminalRows = [];
    if (terminalLastRow > 1) {
      const terminalData = terminalSheet.getRange(1, 1, terminalLastRow, terminalSheet.getLastColumn()).getValues();
      terminalRows = terminalData.slice(1); // ヘッダーを除外
    }
    
    // プリンタ・その他系データを取得
    const printerOtherLastRow = printerOtherSheet.getLastRow();
    let printerOtherRows = [];
    if (printerOtherLastRow > 1) {
      const printerOtherData = printerOtherSheet.getRange(1, 1, printerOtherLastRow, printerOtherSheet.getLastColumn()).getValues();
      printerOtherRows = printerOtherData.slice(1); // ヘッダーを除外
    }
    
    // データを結合
    const rows = [...terminalRows, ...printerOtherRows];
    
    if (rows.length === 0) {
      addLog('統合ビューにデータがありません');
      return {
        success: true,
        rowsIndexed: 0,
        message: '統合ビューにデータがありません'
      };
    }
    
    // ヘッダーはどちらのシートも同じ構造なので、端末系から取得
    const headers = terminalSheet.getRange(1, 1, 1, terminalSheet.getLastColumn()).getValues()[0];
    
    // 必要な列のインデックスを動的に取得
    const columnIndices = {
      managementNumber: getColumnIndex(headers, '拠点管理番号'),
      category: getColumnIndex(headers, 'カテゴリ'),
      modelName: getColumnIndex(headers, '機種名'),
      status: getColumnIndex(headers, '現在ステータス') || getColumnIndex(headers, '0-4.ステータス'),
      lastUpdate: getColumnIndex(headers, '最終更新日時') || getColumnIndex(headers, 'タイムスタンプ'),
      location: getColumnIndex(headers, '拠点名'),
      jurisdiction: getColumnIndex(headers, '管轄')
    };
    
    // インデックスデータを構築
    const indexData = [];
    for (const row of rows) {
      // 拠点管理番号が空の行はスキップ
      if (!row[columnIndices.managementNumber]) continue;
      
      // 拠点コードを拠点管理番号から抽出
      const managementNumber = row[columnIndices.managementNumber];
      const locationCode = extractLocationCode(managementNumber);
      
      // 検索キーを生成
      const searchKey = generateSearchKey(row, columnIndices);
      
      indexData.push([
        managementNumber,                    // A: 拠点管理番号
        row[columnIndices.category] || '',   // B: カテゴリ
        row[columnIndices.modelName] || '',  // C: 機種名
        row[columnIndices.status] || '',     // D: ステータス
        locationCode,                        // E: 拠点コード
        row[columnIndices.jurisdiction] || '', // F: 管轄
        row[columnIndices.lastUpdate] || '', // G: 最終更新
        searchKey                            // H: 検索キー
      ]);
    }
    
    // 検索インデックスシートをクリアして新しいデータを書き込み
    searchSheet.getRange(2, 1, searchSheet.getMaxRows() - 1, 8).clearContent();
    
    if (indexData.length > 0) {
      searchSheet.getRange(2, 1, indexData.length, 8).setValues(indexData);
    }
    
    endPerformanceTimer(startTime, '検索インデックス再構築');
    
    const result = {
      success: true,
      rowsIndexed: indexData.length,
      timestamp: new Date()
    };
    
    addLog('検索インデックス再構築完了', result);
    return result;
    
  } catch (error) {
    endPerformanceTimer(startTime, '検索インデックス再構築エラー');
    addLog('検索インデックス再構築エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 拠点管理番号から拠点コードを抽出
 * @param {string} managementNumber - 拠点管理番号
 * @return {string} 拠点コード
 */
function extractLocationCode(managementNumber) {
  if (!managementNumber) return '';
  
  // 拠点管理番号のフォーマット: 拠点_カテゴリ_モデル_製造番号_連番
  const parts = managementNumber.split('_');
  return parts.length > 0 ? parts[0] : '';
}

/**
 * 検索キーを生成
 * @param {Array} row - データ行
 * @param {Object} columnIndices - 列インデックス
 * @return {string} 検索キー
 */
function generateSearchKey(row, columnIndices) {
  const parts = [];
  
  // 重要な検索対象フィールドを結合
  if (row[columnIndices.managementNumber]) {
    parts.push(row[columnIndices.managementNumber]);
  }
  if (row[columnIndices.category]) {
    parts.push(row[columnIndices.category]);
  }
  if (row[columnIndices.modelName]) {
    parts.push(row[columnIndices.modelName]);
  }
  if (row[columnIndices.status]) {
    parts.push(row[columnIndices.status]);
  }
  if (row[columnIndices.location]) {
    parts.push(row[columnIndices.location]);
  }
  if (row[columnIndices.jurisdiction]) {
    parts.push(row[columnIndices.jurisdiction]);
  }
  
  // スペースで結合して検索キーを作成
  return parts.join(' ');
}

/**
 * 検索インデックスを使用した高速検索
 * @param {string} keyword - 検索キーワード
 * @param {Object} filters - 追加フィルター条件
 * @return {Object} 検索結果
 */
function searchWithIndex(keyword, filters = {}) {
  const startTime = startPerformanceTimer();
  addLog('インデックス検索開始', { keyword, filters });
  
  try {
    const searchSheet = getSearchIndexSheet();
    const lastRow = searchSheet.getLastRow();
    
    if (lastRow <= 1) {
      return {
        success: true,
        results: [],
        totalMatches: 0
      };
    }
    
    // 検索インデックスのヘッダーを取得
    const headers = searchSheet.getRange(1, 1, 1, searchSheet.getLastColumn()).getValues()[0];
    const indexData = searchSheet.getRange(2, 1, lastRow - 1, searchSheet.getLastColumn()).getValues();
    
    // キーワード検索
    const results = [];
    const keywordLower = keyword ? keyword.toLowerCase() : '';
    
    for (const row of indexData) {
      // 検索キーでの検索
      const searchKey = getValueByColumnName(row, headers, '検索キー');
      if (keywordLower && !searchKey.toLowerCase().includes(keywordLower)) {
        continue;
      }
      
      // 追加フィルターの適用
      const category = getValueByColumnName(row, headers, 'カテゴリ');
      const status = getValueByColumnName(row, headers, '状態');
      const jurisdiction = getValueByColumnName(row, headers, '管轄');
      const locationCode = getValueByColumnName(row, headers, '拠点コード');
      
      if (filters.category && category !== filters.category) continue;
      if (filters.status && status !== filters.status) continue;
      if (filters.jurisdiction && jurisdiction !== filters.jurisdiction) continue;
      if (filters.locationCode && locationCode !== filters.locationCode) continue;
      
      results.push({
        managementNumber: getValueByColumnName(row, headers, '拠点管理番号'),
        category: category,
        modelName: getValueByColumnName(row, headers, '機種名'),
        status: status,
        locationCode: locationCode,
        jurisdiction: jurisdiction,
        lastUpdate: getValueByColumnName(row, headers, '最終更新日時')
      });
    }
    
    endPerformanceTimer(startTime, 'インデックス検索');
    
    return {
      success: true,
      results: results,
      totalMatches: results.length
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'インデックス検索エラー');
    addLog('インデックス検索エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 統合ビューデータを構築（マスタシートから）
 * @deprecated 両方の統合ビューを構築するためのbuildAllIntegratedViewsを使用
 * @return {Object} 構築結果
 */
function buildIntegratedViewData() {
  return buildAllIntegratedViews();
}

/**
 * すべての統合ビューデータを構築
 * @return {Object} 構築結果
 */
function buildAllIntegratedViews() {
  const startTime = startPerformanceTimer();
  addLog('すべての統合ビューデータ構築開始');
  
  try {
    // 端末系統合ビューを構築
    const terminalResult = buildIntegratedViewTerminal();
    if (!terminalResult.success) {
      throw new Error('端末系統合ビュー構築失敗: ' + terminalResult.error);
    }
    
    // プリンタ・その他系統合ビューを構築
    const printerOtherResult = buildIntegratedViewPrinterOther();
    if (!printerOtherResult.success) {
      throw new Error('プリンタ・その他系統合ビュー構築失敗: ' + printerOtherResult.error);
    }
    
    endPerformanceTimer(startTime, 'すべての統合ビューデータ構築');
    
    const result = {
      success: true,
      terminal: terminalResult,
      printerOther: printerOtherResult,
      totalRowsCreated: terminalResult.rowsCreated + printerOtherResult.rowsCreated,
      timestamp: new Date()
    };
    
    addLog('すべての統合ビューデータ構築完了', result);
    return result;
    
  } catch (error) {
    endPerformanceTimer(startTime, 'すべての統合ビューデータ構築エラー');
    addLog('すべての統合ビューデータ構築エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 端末系統合ビューデータを構築
 * @return {Object} 構築結果
 */
function buildIntegratedViewTerminal() {
  const startTime = startPerformanceTimer();
  addLog('端末系統合ビューデータ構築開始');
  
  try {
    // 端末マスタシートからデータを取得
    const terminalData = getTerminalMasterData();
    
    // ステータス収集シートから最新データを取得
    const statusData = getLatestStatusCollectionData();
    
    // 拠点マスタデータを取得
    const locationMaster = getLocationMaster();
    const locationMap = {};
    locationMaster.forEach(loc => {
      locationMap[loc.locationId] = loc;
    });
    
    // 端末データを統合
    const integratedData = integrateDeviceData(terminalData, statusData, locationMap, 'terminal');
    
    // 端末系統合ビューシートに書き込み
    const viewSheet = getIntegratedViewTerminalSheet();
    
    // 既存データをクリア
    if (viewSheet.getLastRow() > 1) {
      viewSheet.getRange(2, 1, viewSheet.getLastRow() - 1, viewSheet.getLastColumn()).clearContent();
    }
    
    // 新しいデータを書き込み
    if (integratedData.length > 0) {
      viewSheet.getRange(2, 1, integratedData.length, integratedData[0].length).setValues(integratedData);
    }
    
    endPerformanceTimer(startTime, '端末系統合ビューデータ構築');
    
    const result = {
      success: true,
      rowsCreated: integratedData.length,
      timestamp: new Date()
    };
    
    addLog('端末系統合ビューデータ構築完了', result);
    return result;
    
  } catch (error) {
    endPerformanceTimer(startTime, '端末系統合ビューデータ構築エラー');
    addLog('端末系統合ビューデータ構築エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * プリンタ・その他系統合ビューデータを構築
 * @return {Object} 構築結果
 */
function buildIntegratedViewPrinterOther() {
  const startTime = startPerformanceTimer();
  addLog('プリンタ・その他系統合ビューデータ構築開始');
  
  try {
    // プリンタとその他マスタシートからデータを取得
    const printerData = getPrinterMasterData();
    const otherData = getOtherMasterData();
    
    // ステータス収集シートから最新データを取得
    const statusData = getLatestStatusCollectionData();
    
    // 拠点マスタデータを取得
    const locationMaster = getLocationMaster();
    const locationMap = {};
    locationMaster.forEach(loc => {
      locationMap[loc.locationId] = loc;
    });
    
    // プリンタとその他データを統合
    const printerIntegratedData = integrateDeviceData(printerData, statusData, locationMap, 'printer');
    const otherIntegratedData = integrateDeviceData(otherData, statusData, locationMap, 'other');
    const integratedData = [...printerIntegratedData, ...otherIntegratedData];
    
    // プリンタ・その他系統合ビューシートに書き込み
    const viewSheet = getIntegratedViewPrinterOtherSheet();
    
    // 既存データをクリア
    if (viewSheet.getLastRow() > 1) {
      viewSheet.getRange(2, 1, viewSheet.getLastRow() - 1, viewSheet.getLastColumn()).clearContent();
    }
    
    // 新しいデータを書き込み
    if (integratedData.length > 0) {
      viewSheet.getRange(2, 1, integratedData.length, integratedData[0].length).setValues(integratedData);
    }
    
    endPerformanceTimer(startTime, 'プリンタ・その他系統合ビューデータ構築');
    
    const result = {
      success: true,
      rowsCreated: integratedData.length,
      timestamp: new Date()
    };
    
    addLog('プリンタ・その他系統合ビューデータ構築完了', result);
    return result;
    
  } catch (error) {
    endPerformanceTimer(startTime, 'プリンタ・その他系統合ビューデータ構築エラー');
    addLog('プリンタ・その他系統合ビューデータ構築エラー', { error: error.toString() });
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * デバイスデータを統合形式に変換
 * @param {Array} deviceData - デバイスマスタデータ
 * @param {Object} statusData - ステータスデータ
 * @param {Object} locationMap - 拠点マスタマップ
 * @param {string} deviceType - デバイスタイプ
 * @return {Array} 統合されたデータ
 */
function integrateDeviceData(deviceData, statusData, locationMap, deviceType) {
  const integratedRows = [];
  
  for (const device of deviceData) {
    const managementNumber = device.managementNumber || device['拠点管理番号'];
    if (!managementNumber) continue;
    
    // ステータスデータから最新情報を取得
    const latestStatus = statusData[managementNumber] || {};
    
    // 拠点情報を取得
    const locationCode = extractLocationCode(managementNumber);
    const locationInfo = locationMap[locationCode] || {};
    
    // 貸出日数を計算
    const loanDays = calculateLoanDays(latestStatus);
    
    // 要注意フラグを判定
    const cautionFlag = determineCautionFlag(latestStatus, loanDays);
    
    integratedRows.push([
      managementNumber,                          // A: 拠点管理番号
      device.category || device['カテゴリ'] || '', // B: カテゴリ
      device.modelName || device['機種名'] || '', // C: 機種名
      device.serialNumber || device['製造番号'] || '', // D: 製造番号
      device.assetNumber || device['資産管理番号'] || '', // E: 資産管理番号
      device.software || device['ソフトウェア'] || '', // F: ソフトウェア
      device.os || device['OS'] || '',           // G: OS
      latestStatus.timestamp || '',               // H: 最終更新日時
      latestStatus.status || '',                  // I: 現在ステータス
      latestStatus.customerName || '',            // J: 顧客名
      latestStatus.customerNumber || '',          // K: 顧客番号
      latestStatus.address || '',                 // L: 住所
      latestStatus.userMachineFlag || '',         // M: ユーザー機預り有無
      latestStatus.userMachineSerial || '',       // N: 預りユーザー機シリアル
      latestStatus.receiptNumber || '',           // O: お預かり証No
      latestStatus.internalStatus || '',          // P: 社内ステータス
      latestStatus.inventoryFlag || '',           // Q: 棚卸フラグ
      latestStatus.currentLocation || locationCode, // R: 現在拠点
      latestStatus.remarks || '',                 // S: 備考
      loanDays,                                   // T: 貸出日数
      cautionFlag,                                // U: 要注意フラグ
      locationInfo.locationName || locationCode,  // V: 拠点名
      locationInfo.jurisdiction || '',            // W: 管轄
      device.formURL || '',                       // X: formURL
      device.qrCodeURL || ''                      // Y: QRコードURL
    ]);
  }
  
  return integratedRows;
}

/**
 * 貸出日数を計算
 * @param {Object} statusData - ステータスデータ
 * @return {number} 貸出日数
 */
function calculateLoanDays(statusData) {
  if (!statusData.status || statusData.status !== '1.貸出中') {
    return 0;
  }
  
  if (!statusData.loanDate) {
    return 0;
  }
  
  try {
    const loanDate = new Date(statusData.loanDate);
    const today = new Date();
    const diffTime = today - loanDate;
    const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
    return diffDays >= 0 ? diffDays : 0;
  } catch (error) {
    return 0;
  }
}

/**
 * 要注意フラグを判定
 * @param {Object} statusData - ステータスデータ
 * @param {number} loanDays - 貸出日数
 * @return {boolean} 要注意フラグ
 */
function determineCautionFlag(statusData, loanDays) {
  // 90日以上貸出中の場合は要注意
  if (loanDays >= 90) {
    return true;
  }
  
  // 修理中で30日以上経過している場合は要注意
  if (statusData.internalStatus === '1.修理中' && statusData.repairDays > 30) {
    return true;
  }
  
  return false;
}

/**
 * 定期的な検索インデックス更新用トリガー関数
 */
function updateSearchIndexTrigger() {
  try {
    // すべての統合ビューを更新
    const viewResult = buildAllIntegratedViews();
    if (!viewResult.success) {
      console.error('統合ビュー更新失敗:', viewResult.error);
      return;
    }
    
    // 検索インデックスを再構築
    const indexResult = rebuildSearchIndex();
    if (!indexResult.success) {
      console.error('検索インデックス更新失敗:', indexResult.error);
      return;
    }
    
    console.log('検索インデックス更新完了:', {
      terminalRows: viewResult.terminal.rowsCreated,
      printerOtherRows: viewResult.printerOther.rowsCreated,
      totalViewRows: viewResult.totalRowsCreated,
      indexRows: indexResult.rowsIndexed,
      timestamp: new Date()
    });
    
  } catch (error) {
    console.error('検索インデックス更新エラー:', error);
  }
}

/**
 * 検索インデックスのテスト関数
 */
function testSearchIndex() {
  console.log('=== 検索インデックステスト開始 ===');
  
  // すべての統合ビューデータ構築
  const viewResult = buildAllIntegratedViews();
  console.log('統合ビュー構築結果:', viewResult);
  
  // 検索インデックス再構築
  const indexResult = rebuildSearchIndex();
  console.log('インデックス構築結果:', indexResult);
  
  // 検索テスト
  const searchResult = searchWithIndex('ThinkPad', { jurisdiction: '関西' });
  console.log('検索結果:', searchResult);
  
  console.log('=== 検索インデックステスト完了 ===');
}