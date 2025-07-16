// ========================================
// スプレッドシート操作関連
// ========================================

/**
 * 拠点ID変換関数（拠点マスタベース・互換性維持）
 * @param {string} location - 拠点ID
 * @return {string} 変換後の拠点ID
 */
function convertLegacyLocation(location) {
  // 拠点マスタベースでは直接IDを使用
  return location;
}

/**
 * 統一スプレッドシートIDをスクリプトプロパティから取得
 * @param {string} location - 拠点ID
 * @return {string} スプレッドシートID
 */
function getSpreadsheetIdFromProperty(location) {
  // 旧拠点名を新拠点名に変換
  const normalizedLocation = convertLegacyLocation(location);
  
  // 全拠点で統一スプレッドシートを使用
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
}

/**
 * 統一スプレッドシートIDを取得（getDestinationSheetData用）
 * @return {string} スプレッドシートID
 */
function getUnifiedSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
}

/**
 * 指定された拠点とデバイスタイプのスプレッドシートデータを取得
 * @param {string} location - 拠点ID
 * @param {string} queryType - クエリタイプ
 * @param {string} deviceType - デバイスタイプ（デフォルト: 'terminal'）
 * @return {Object} データ取得結果
 */
function getSpreadsheetData(location, queryType, deviceType = 'terminal') {
  const startTime = startPerformanceTimer();
  addLog('getSpreadsheetData関数が呼び出されました', { location, queryType, deviceType });
  
  try {
    // 旧拠点名を新拠点名に変換
    const normalizedLocation = convertLegacyLocation(location);
    
    // 統一スプレッドシートIDを取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません: ' + location);
    }

    addLog('使用するスプレッドシートID', spreadsheetId);
    
    // デバイスタイプに応じてマスタシートを決定
    const targetSheetName = getTargetSheetName(deviceType);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(targetSheetName);
    
    if (!sheet) {
      throw new Error('シート「' + targetSheetName + '」が見つかりません。');
    }

    // データを取得
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow === 0) {
      throw new Error('シートにデータがありません。');
    }
    
    const range = sheet.getRange(1, 1, lastRow, lastColumn);
    const data = range.getValues();
    
    // 日付データの処理
    const dateProcessingStart = Date.now();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        if (row[j] instanceof Date) {
          row[j] = formatDateFast(row[j]);
        }
      }
    }
    const dateProcessingTime = Date.now() - dateProcessingStart;
    addLog('日付処理時間', dateProcessingTime + 'ms');
    
    const responseTime = endPerformanceTimer(startTime, 'スプレッドシート取得');
    
    const response = {
      success: true,
      data: data,
      logs: DEBUG ? serverLogs : [],
      metadata: {
        location: getLocationNameById(location),
        queryType: queryType,
        spreadsheetName: spreadsheet.getName(),
        sheetName: sheet.getName(),
        lastRow: lastRow,
        lastColumn: lastColumn,
        responseTime: responseTime,
        dateProcessingTime: dateProcessingTime,
        dataSize: data.length
      }
    };
    
    addLog('返却するレスポンス準備完了');
    return response;
    
  } catch (error) {
    endPerformanceTimer(startTime, 'エラー処理');
    addLog('エラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack,
      name: error.name
    });
    
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack,
        name: error.name
      },
      logs: DEBUG ? serverLogs : []
    };
  }
}

/**
 * スプレッドシートIDが設定されているかチェック
 * @param {string} location - 拠点ID
 * @return {Object} チェック結果
 */
function checkSpreadsheetIdExists(location) {
  try {
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    return {
      success: true,
      exists: !!spreadsheetId,
      location: location,
      locationName: getLocationNameById(location),
      spreadsheetId: spreadsheetId
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      location: location,
      locationName: getLocationNameById(location)
    };
  }
}

/**
 * ページネーション付きでスプレッドシートデータを取得
 * @param {string} location - 拠点ID
 * @param {string} queryType - クエリタイプ
 * @param {number} startRow - 開始行（デフォルト: 1）
 * @param {number} maxRows - 最大行数（デフォルト: 100）
 * @param {string} deviceType - デバイスタイプ（デフォルト: 'terminal'）
 * @return {Object} ページネーション付きデータ
 */
function getSpreadsheetDataPaginated(location, queryType, startRow = 1, maxRows = 100, deviceType = 'terminal') {
  addLog('getSpreadsheetDataPaginated関数が呼び出されました', { location, queryType, startRow, maxRows, deviceType });
  
  try {
    // 旧拠点名を新拠点名に変換
    const normalizedLocation = convertLegacyLocation(location);
    
    // 統一スプレッドシートIDを取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません: ' + location);
    }

    // デバイスタイプに応じてマスタシートを決定
    const targetSheetName = getTargetSheetName(deviceType);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(targetSheetName);
    
    if (!sheet) {
      throw new Error('シート「' + targetSheetName + '」が見つかりません。');
    }

    const totalRows = sheet.getLastRow();
    const totalColumns = sheet.getLastColumn();
    
    if (totalRows === 0) {
      throw new Error('シートにデータがありません。');
    }
    
    // ヘッダー行を取得
    const headerRange = sheet.getRange(1, 1, 1, totalColumns);
    const headers = headerRange.getValues()[0];
    
    // 指定された範囲のデータを取得
    const endRow = Math.min(startRow + maxRows - 1, totalRows);
    const actualStartRow = Math.max(startRow, 2); // ヘッダー行をスキップ
    
    let data = [headers]; // ヘッダーを最初に追加
    
    if (actualStartRow <= totalRows) {
      const dataRange = sheet.getRange(actualStartRow, 1, endRow - actualStartRow + 1, totalColumns);
      const rowData = dataRange.getValues();
      
      // 日付データの処理
      for (let i = 0; i < rowData.length; i++) {
        const row = rowData[i];
        for (let j = 0; j < row.length; j++) {
          if (row[j] instanceof Date) {
            row[j] = formatDateFast(row[j]);
          }
        }
      }
      
      data = data.concat(rowData);
    }
    
    const response = {
      success: true,
      data: data,
      pagination: {
        startRow: startRow,
        endRow: endRow,
        maxRows: maxRows,
        totalRows: totalRows,
        hasMore: endRow < totalRows
      },
      metadata: {
        location: getLocationNameById(location),
        queryType: queryType,
        spreadsheetName: spreadsheet.getName(),
        sheetName: sheet.getName(),
        totalRows: totalRows,
        dataSize: data.length
      },
      logs: DEBUG ? serverLogs : []
    };
    
    return response;
    
  } catch (error) {
    addLog('エラーが発生', {
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.toString(),
      logs: DEBUG ? serverLogs : []
    };
  }
}

/**
 * 機器のステータスを更新
 * @param {number} rowIndex - 更新する行番号
 * @param {string} newStatus - 新しいステータス
 * @param {string} location - 拠点ID
 * @param {string} deviceType - デバイスタイプ（デフォルト: 'terminal'）
 * @return {Object} 更新結果
 */
function updateMachineStatus(rowIndex, newStatus, location, deviceType = 'terminal') {
  const startTime = startPerformanceTimer();
  
  try {
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません: ' + location);
    }
    
    // デバイスタイプに応じてマスタシートを決定
    const targetSheetName = getTargetSheetName(deviceType);
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(targetSheetName);
    
    if (!sheet) {
      throw new Error('シート「' + targetSheetName + '」が見つかりません。');
    }
    
    // ステータス列と更新日時列を取得（ヘッダー行から）
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol = headers.indexOf('ステータス') + 1;
    const updateDateCol = headers.indexOf('更新日時') + 1;
    
    if (statusCol === 0) {
      throw new Error('ステータス列が見つかりません。');
    }
    
    // 実際の行番号（ヘッダーを考慮）
    const actualRow = rowIndex + 1;
    
    // ステータスを更新
    sheet.getRange(actualRow, statusCol).setValue(newStatus);
    
    // 更新日時があれば更新
    if (updateDateCol > 0) {
      const now = new Date();
      const formattedDate = formatDateFast(now);
      sheet.getRange(actualRow, updateDateCol).setValue(formattedDate);
    }
    
    SpreadsheetApp.flush();
    
    const responseTime = endPerformanceTimer(startTime, 'ステータス更新');
    
    return {
      success: true,
      message: 'ステータスが更新されました',
      rowIndex: rowIndex,
      newStatus: newStatus,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'ステータス更新エラー');
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 複数の機器のステータスを一括更新
 * @param {Array} updates - 更新情報の配列 [{rowIndex, newStatus}]
 * @param {string} location - 拠点ID
 * @param {string} deviceType - デバイスタイプ（デフォルト: 'terminal'）
 * @return {Object} 一括更新結果
 */
function updateMultipleStatuses(updates, location, deviceType = 'terminal') {
  const startTime = startPerformanceTimer();
  
  try {
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません: ' + location);
    }
    
    // デバイスタイプに応じてマスタシートを決定
    const targetSheetName = getTargetSheetName(deviceType);
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(targetSheetName);
    
    if (!sheet) {
      throw new Error('シート「' + targetSheetName + '」が見つかりません。');
    }
    
    // ステータス列と更新日時列を取得
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol = headers.indexOf('ステータス') + 1;
    const updateDateCol = headers.indexOf('更新日時') + 1;
    
    if (statusCol === 0) {
      throw new Error('ステータス列が見つかりません。');
    }
    
    const now = new Date();
    const formattedDate = formatDateFast(now);
    
    // バッチ処理で更新
    const updatedRows = [];
    updates.forEach(update => {
      const actualRow = update.rowIndex + 1;
      sheet.getRange(actualRow, statusCol).setValue(update.newStatus);
      
      if (updateDateCol > 0) {
        sheet.getRange(actualRow, updateDateCol).setValue(formattedDate);
      }
      
      updatedRows.push(actualRow);
    });
    
    SpreadsheetApp.flush();
    
    const responseTime = endPerformanceTimer(startTime, '一括ステータス更新');
    
    return {
      success: true,
      message: `${updates.length}件のステータスが更新されました`,
      updatedRows: updatedRows,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '一括ステータス更新エラー');
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * データの整合性をチェック
 * @param {string} location - 拠点ID
 * @param {string} deviceType - デバイスタイプ（デフォルト: 'terminal'）
 * @return {Object} 整合性チェック結果
 */
function checkDataConsistency(location, deviceType = 'terminal') {
  const startTime = startPerformanceTimer();
  
  try {
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません: ' + location);
    }
    
    // デバイスタイプに応じてマスタシートを決定
    const targetSheetName = getTargetSheetName(deviceType);
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(targetSheetName);
    
    if (!sheet) {
      throw new Error('シート「' + targetSheetName + '」が見つかりません。');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const issues = [];
    const stats = {
      totalRows: rows.length,
      emptyRows: 0,
      duplicateIds: [],
      invalidDates: [],
      missingRequiredFields: []
    };
    
    // 必須フィールドのインデックスを取得
    const idCol = headers.indexOf('管理番号');
    const statusCol = headers.indexOf('ステータス');
    
    const idMap = new Map();
    
    rows.forEach((row, index) => {
      const rowNum = index + 2; // 実際の行番号
      
      // 空行チェック
      if (row.every(cell => !cell || cell === '')) {
        stats.emptyRows++;
        return;
      }
      
      // 管理番号の重複チェック
      if (idCol >= 0 && row[idCol]) {
        const id = row[idCol];
        if (idMap.has(id)) {
          stats.duplicateIds.push({
            id: id,
            rows: [idMap.get(id), rowNum]
          });
        } else {
          idMap.set(id, rowNum);
        }
      }
      
      // 必須フィールドチェック
      if (idCol >= 0 && !row[idCol]) {
        stats.missingRequiredFields.push({
          row: rowNum,
          field: '管理番号'
        });
      }
      
      if (statusCol >= 0 && !row[statusCol]) {
        stats.missingRequiredFields.push({
          row: rowNum,
          field: 'ステータス'
        });
      }
    });
    
    // 問題がある場合はissuesに追加
    if (stats.emptyRows > 0) {
      issues.push(`空行が${stats.emptyRows}行あります`);
    }
    
    if (stats.duplicateIds.length > 0) {
      issues.push(`重複する管理番号が${stats.duplicateIds.length}件あります`);
    }
    
    if (stats.missingRequiredFields.length > 0) {
      issues.push(`必須フィールドの欠損が${stats.missingRequiredFields.length}件あります`);
    }
    
    const responseTime = endPerformanceTimer(startTime, 'データ整合性チェック');
    
    return {
      success: true,
      isConsistent: issues.length === 0,
      issues: issues,
      stats: stats,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'データ整合性チェックエラー');
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 拠点別シートのデータを取得
 * @param {string} location - 拠点ID
 * @param {string} locationSheetName - 拠点シート名
 * @param {string} queryType - クエリタイプ
 * @return {Object} シートデータ
 */
function getLocationSheetData(location, locationSheetName, queryType) {
  const startTime = startPerformanceTimer();
  addLog('getLocationSheetData関数が呼び出されました', { location, locationSheetName, queryType });
  
  try {
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません: ' + location);
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(locationSheetName);
    
    if (!sheet) {
      throw new Error(`シート「${locationSheetName}」が見つかりません。`);
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow === 0) {
      throw new Error('シートにデータがありません。');
    }
    
    const range = sheet.getRange(1, 1, lastRow, lastColumn);
    const data = range.getValues();
    
    // 日付データの処理
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      for (let j = 0; j < row.length; j++) {
        if (row[j] instanceof Date) {
          row[j] = formatDateFast(row[j]);
        }
      }
    }
    
    const responseTime = endPerformanceTimer(startTime, '拠点シートデータ取得');
    
    return {
      success: true,
      data: data,
      metadata: {
        location: getLocationNameById(location),
        sheetName: locationSheetName,
        lastRow: lastRow,
        lastColumn: lastColumn,
        responseTime: responseTime
      },
      logs: DEBUG ? serverLogs : []
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '拠点シートデータ取得エラー');
    addLog('エラーが発生', {
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.toString(),
      logs: DEBUG ? serverLogs : []
    };
  }
}

/**
 * 拠点別シートのセルを更新
 * @param {string} location - 拠点ID
 * @param {string} locationSheetName - 拠点シート名
 * @param {number} rowIndex - 行番号
 * @param {number} columnIndex - 列番号
 * @param {*} newValue - 新しい値
 * @return {Object} 更新結果
 */
function updateLocationSheetCell(location, locationSheetName, rowIndex, columnIndex, newValue) {
  const startTime = startPerformanceTimer();
  
  try {
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません: ' + location);
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(locationSheetName);
    
    if (!sheet) {
      throw new Error(`シート「${locationSheetName}」が見つかりません。`);
    }
    
    // 値を更新
    sheet.getRange(rowIndex + 1, columnIndex + 1).setValue(newValue);
    
    // 更新日時列があれば更新
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const updateDateCol = headers.indexOf('更新日時') + 1;
    
    if (updateDateCol > 0) {
      const now = new Date();
      const formattedDate = formatDateFast(now);
      sheet.getRange(rowIndex + 1, updateDateCol).setValue(formattedDate);
    }
    
    SpreadsheetApp.flush();
    
    const responseTime = endPerformanceTimer(startTime, '拠点シートセル更新');
    
    return {
      success: true,
      message: 'セルが更新されました',
      rowIndex: rowIndex,
      columnIndex: columnIndex,
      newValue: newValue,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '拠点シートセル更新エラー');
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 統一スプレッドシートのデータを取得
 * @param {string} sheetName - シート名
 * @param {string} queryType - クエリタイプ
 * @return {Object} シートデータ
 */
function getDestinationSheetData(sheetName, queryType) {
  const startTime = startPerformanceTimer();
  addLog('getDestinationSheetData関数が呼び出されました', { sheetName, queryType });
  
  try {
    const spreadsheetId = getUnifiedSpreadsheetId();
    if (!spreadsheetId) {
      throw new Error('統一スプレッドシートIDが設定されていません');
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません。`);
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow === 0) {
      return {
        success: true,
        data: [],
        metadata: {
          sheetName: sheetName,
          isEmpty: true
        }
      };
    }
    
    const range = sheet.getRange(1, 1, lastRow, lastColumn);
    const data = range.getValues();
    
    // シートタイプを判定
    const sheetType = getSheetTypeFromName(sheetName);
    
    // 日付データの処理とデータ整形
    const processedData = [];
    const headers = data[0];
    processedData.push(headers);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const processedRow = [];
      
      for (let j = 0; j < row.length; j++) {
        let cellValue = row[j];
        
        // 日付の処理
        if (cellValue instanceof Date) {
          cellValue = formatDateFast(cellValue);
        }
        
        // 拠点IDを拠点名に変換（必要に応じて）
        if (sheetType !== 'location' && headers[j] === '拠点' && cellValue) {
          const locationName = getLocationNameById(cellValue);
          if (locationName && locationName !== cellValue) {
            cellValue = locationName;
          }
        }
        
        processedRow.push(cellValue);
      }
      
      processedData.push(processedRow);
    }
    
    const responseTime = endPerformanceTimer(startTime, '統一シートデータ取得');
    
    return {
      success: true,
      data: processedData,
      metadata: {
        sheetName: sheetName,
        sheetType: sheetType,
        lastRow: lastRow,
        lastColumn: lastColumn,
        responseTime: responseTime
      },
      logs: DEBUG ? serverLogs : []
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '統一シートデータ取得エラー');
    addLog('エラーが発生', {
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.toString(),
      logs: DEBUG ? serverLogs : []
    };
  }
}

/**
 * シート名からシートタイプを判定
 * @param {string} sheetName - シート名
 * @return {string} シートタイプ
 */
function getSheetTypeFromName(sheetName) {
  if (sheetName === '拠点マスタ') return 'location';
  if (sheetName === '機種マスタ') return 'model';
  if (sheetName === '端末マスタ') return 'terminal';
  if (sheetName === 'プリンタマスタ') return 'printer';
  if (sheetName === 'その他マスタ') return 'other';
  if (sheetName === 'サマリー') return 'summary';
  if (sheetName === '監査データ') return 'audit';
  if (sheetName === 'integrated_view') return 'integrated_view';
  if (sheetName === 'search_index') return 'search_index';
  if (sheetName === 'summary_view') return 'summary_view';
  return 'unknown';
}

/**
 * 統一スプレッドシートのシート一覧を取得
 * @return {Object} シート一覧
 */
function getDestinationSheets() {
  const startTime = startPerformanceTimer();
  
  try {
    const spreadsheetId = getUnifiedSpreadsheetId();
    if (!spreadsheetId) {
      throw new Error('統一スプレッドシートIDが設定されていません');
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    
    const sheetList = sheets.map(sheet => ({
      name: sheet.getName(),
      index: sheet.getIndex(),
      lastRow: sheet.getLastRow(),
      lastColumn: sheet.getLastColumn()
    }));
    
    const responseTime = endPerformanceTimer(startTime, 'シート一覧取得');
    
    return {
      success: true,
      sheets: sheetList,
      totalSheets: sheets.length,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'シート一覧取得エラー');
    return {
      success: false,
      error: error.toString()
    };
  }
}