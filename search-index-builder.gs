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
    // 統合ビューシートからデータを取得
    const viewSheet = getIntegratedViewSheet();
    const searchSheet = getSearchIndexSheet();
    
    const lastRow = viewSheet.getLastRow();
    if (lastRow <= 1) {
      addLog('統合ビューにデータがありません');
      return {
        success: true,
        rowsIndexed: 0,
        message: '統合ビューにデータがありません'
      };
    }
    
    // 統合ビューのデータを取得
    const viewData = viewSheet.getRange(1, 1, lastRow, viewSheet.getLastColumn()).getValues();
    const headers = viewData[0];
    const rows = viewData.slice(1);
    
    // 必要な列のインデックスを取得
    const columnIndices = {
      managementNumber: headers.indexOf('拠点管理番号'),
      category: headers.indexOf('カテゴリ'),
      modelName: headers.indexOf('機種名'),
      status: headers.indexOf('現在ステータス'),
      lastUpdate: headers.indexOf('最終更新日時'),
      location: headers.indexOf('拠点名'),
      jurisdiction: headers.indexOf('管轄')
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
    
    // 検索インデックスデータを取得
    const indexData = searchSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    // キーワード検索
    const results = [];
    const keywordLower = keyword ? keyword.toLowerCase() : '';
    
    for (const row of indexData) {
      // 検索キーでの検索
      const searchKey = row[7] || ''; // H列: 検索キー
      if (keywordLower && !searchKey.toLowerCase().includes(keywordLower)) {
        continue;
      }
      
      // 追加フィルターの適用
      if (filters.category && row[1] !== filters.category) continue;
      if (filters.status && row[3] !== filters.status) continue;
      if (filters.jurisdiction && row[5] !== filters.jurisdiction) continue;
      if (filters.locationCode && row[4] !== filters.locationCode) continue;
      
      results.push({
        managementNumber: row[0],
        category: row[1],
        modelName: row[2],
        status: row[3],
        locationCode: row[4],
        jurisdiction: row[5],
        lastUpdate: row[6]
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
 * @return {Object} 構築結果
 */
function buildIntegratedViewData() {
  const startTime = startPerformanceTimer();
  addLog('統合ビューデータ構築開始');
  
  try {
    // 各マスタシートからデータを取得
    const terminalData = getTerminalMasterData();
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
    
    // データを統合
    const integratedData = [];
    
    // 端末データの統合
    integratedData.push(...integrateDeviceData(terminalData, statusData, locationMap, 'terminal'));
    
    // プリンタデータの統合
    integratedData.push(...integrateDeviceData(printerData, statusData, locationMap, 'printer'));
    
    // その他データの統合
    integratedData.push(...integrateDeviceData(otherData, statusData, locationMap, 'other'));
    
    // 統合ビューシートに書き込み
    const viewSheet = getIntegratedViewSheet();
    
    // 既存データをクリア
    if (viewSheet.getLastRow() > 1) {
      viewSheet.getRange(2, 1, viewSheet.getLastRow() - 1, viewSheet.getLastColumn()).clearContent();
    }
    
    // 新しいデータを書き込み
    if (integratedData.length > 0) {
      viewSheet.getRange(2, 1, integratedData.length, integratedData[0].length).setValues(integratedData);
    }
    
    endPerformanceTimer(startTime, '統合ビューデータ構築');
    
    const result = {
      success: true,
      rowsCreated: integratedData.length,
      timestamp: new Date()
    };
    
    addLog('統合ビューデータ構築完了', result);
    return result;
    
  } catch (error) {
    endPerformanceTimer(startTime, '統合ビューデータ構築エラー');
    addLog('統合ビューデータ構築エラー', { error: error.toString() });
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
    // 統合ビューを更新
    const viewResult = buildIntegratedViewData();
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
      viewRows: viewResult.rowsCreated,
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
  
  // 統合ビューデータ構築
  const viewResult = buildIntegratedViewData();
  console.log('統合ビュー構築結果:', viewResult);
  
  // 検索インデックス再構築
  const indexResult = rebuildSearchIndex();
  console.log('インデックス構築結果:', indexResult);
  
  // 検索テスト
  const searchResult = searchWithIndex('ThinkPad', { jurisdiction: '関西' });
  console.log('検索結果:', searchResult);
  
  console.log('=== 検索インデックステスト完了 ===');
}