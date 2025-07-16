// ========================================
// 機種マスタ操作関連
// ========================================

/**
 * 機種ID自動生成機能
 * @return {string} 次の機種ID（M00001形式）
 */
function generateNextModelId() {
  const startTime = startPerformanceTimer();
  addLog('generateNextModelId関数が呼び出されました');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      // シートが存在しない場合は M00001 から開始
      addLog('機種マスタシートが存在しないため、M00001から開始');
      return 'M00001';
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      // データがない場合は M00001 から開始
      addLog('データが存在しないため、M00001から開始');
      return 'M00001';
    }
    
    // 既存の機種IDを取得して最大値を見つける
    const idRange = sheet.getRange(2, 1, lastRow - 1, 1); // A列の2行目以降
    const idValues = idRange.getValues().flat();
    
    let maxNumber = 0;
    const modelIdPattern = /^M(\d{5})$/;
    
    for (const id of idValues) {
      if (id && typeof id === 'string') {
        const match = id.match(modelIdPattern);
        if (match) {
          const number = parseInt(match[1], 10);
          if (number > maxNumber) {
            maxNumber = number;
          }
        }
      }
    }
    
    // 次の番号を生成
    const nextNumber = maxNumber + 1;
    const nextId = 'M' + nextNumber.toString().padStart(5, '0');
    
    addLog('機種ID生成完了', {
      maxNumber: maxNumber,
      nextNumber: nextNumber,
      nextId: nextId,
      existingIds: idValues.length
    });
    
    const responseTime = endPerformanceTimer(startTime, '機種ID生成');
    
    return nextId;
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種ID生成エラー');
    addLog('機種ID生成でエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    // エラーの場合もデフォルトIDを返す
    return 'M001';
  }
}

/**
 * 機種マスタデータを取得
 * @return {Object} 機種マスタデータの結果オブジェクト
 */
function getModelMasterData() {
  const startTime = startPerformanceTimer();
  addLog('getModelMasterData関数が呼び出されました');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにSPREADSHEET_ID_DESTINATIONが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      addLog('機種マスタシートが存在しないため新規作成します');
      sheet = spreadsheet.insertSheet(MASTER_SHEET_NAMES.model);
      const headers = ['機種ID', '機種名', 'メーカー', 'カテゴリ', '作成日', '更新日', '備考'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // 列フォーマットを設定
      const dateColumn1 = sheet.getRange('E:E'); // 作成日列（機種ID追加により1列ずれる）
      const dateColumn2 = sheet.getRange('F:F'); // 更新日列
      dateColumn1.setNumberFormat('@'); // @は文字列フォーマット
      dateColumn2.setNumberFormat('@');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // デバッグログ
    addLog('機種マスタデータ取得成功', {
      headers: headers,
      rowCount: rows.length
    });
    
    const responseTime = endPerformanceTimer(startTime, '機種マスタ取得');
    
    return {
      success: true,
      headers: headers,
      data: rows,
      rowCount: rows.length,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタ取得エラー');
    addLog('機種マスタデータ取得でエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.toString(),
      headers: [],
      data: []
    };
  }
}

/**
 * 機種マスタにデータを追加
 * @param {Object} modelData - 追加する機種データ
 * @return {Object} 処理結果
 */
function addModelMasterData(modelData) {
  const startTime = startPerformanceTimer();
  addLog('addModelMasterData関数が呼び出されました', { modelData });
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      // シートが存在しない場合は作成
      sheet = spreadsheet.insertSheet(MASTER_SHEET_NAMES.model);
      const headers = ['機種ID', '機種名', 'メーカー', 'カテゴリ', '作成日', '更新日', '備考'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // 列フォーマットを設定
      const dateColumn1 = sheet.getRange('E:E'); // 作成日列
      const dateColumn2 = sheet.getRange('F:F'); // 更新日列
      dateColumn1.setNumberFormat('@'); // @は文字列フォーマット
      dateColumn2.setNumberFormat('@');
    }
    
    // 機種IDの自動生成（指定されていない場合）
    if (!modelData.modelId) {
      modelData.modelId = generateNextModelId();
      addLog('機種ID自動生成', { generatedId: modelData.modelId });
    }
    
    // 現在の日時を文字列フォーマットで取得
    const now = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd');
    
    // 新しい行を追加
    const newRow = sheet.getLastRow() + 1;
    const rowData = [
      modelData.modelId,
      modelData.modelName,
      modelData.manufacturer,
      modelData.category,
      now,  // 作成日（文字列）
      now,  // 更新日（文字列）
      modelData.notes || ''
    ];
    
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    
    addLog('機種マスタデータ追加成功', {
      modelId: modelData.modelId,
      rowNumber: newRow,
      addedData: rowData
    });
    
    const responseTime = endPerformanceTimer(startTime, '機種マスタ追加');
    
    return {
      success: true,
      message: '機種データが正常に追加されました',
      modelId: modelData.modelId,
      rowNumber: newRow,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタ追加エラー');
    addLog('機種マスタデータ追加でエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 機種マスタデータを更新
 * @param {number} rowIndex - 更新する行番号（0ベース、ヘッダーを含む）
 * @param {Object} modelData - 更新データ
 * @return {Object} 処理結果
 */
function updateModelMasterData(rowIndex, modelData) {
  const startTime = startPerformanceTimer();
  addLog('updateModelMasterData関数が呼び出されました', { rowIndex, modelData });
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      throw new Error('機種マスタシートが見つかりません');
    }
    
    // シート上の実際の行番号（0ベースから1ベースに変換）
    const sheetRowNumber = rowIndex + 1;
    const lastRow = sheet.getLastRow();
    
    if (sheetRowNumber < 2 || sheetRowNumber > lastRow) {
      throw new Error('指定された行番号が無効です');
    }
    
    // 現在の行データを取得
    const currentData = sheet.getRange(sheetRowNumber, 1, 1, 7).getValues()[0];
    
    // 現在の日時を文字列フォーマットで取得
    const now = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy/MM/dd');
    
    // 更新するデータを準備（変更がない場合は現在の値を保持）
    const updatedData = [
      modelData.modelId !== undefined ? modelData.modelId : currentData[0],
      modelData.modelName !== undefined ? modelData.modelName : currentData[1],
      modelData.manufacturer !== undefined ? modelData.manufacturer : currentData[2],
      modelData.category !== undefined ? modelData.category : currentData[3],
      currentData[4], // 作成日は変更しない
      now, // 更新日を現在日時に更新
      modelData.notes !== undefined ? modelData.notes : currentData[6]
    ];
    
    // データを更新
    sheet.getRange(sheetRowNumber, 1, 1, updatedData.length).setValues([updatedData]);
    
    addLog('機種マスタデータ更新成功', {
      rowIndex: rowIndex,
      sheetRowNumber: sheetRowNumber,
      updatedData: updatedData
    });
    
    const responseTime = endPerformanceTimer(startTime, '機種マスタ更新');
    
    return {
      success: true,
      message: '機種データが正常に更新されました',
      rowNumber: sheetRowNumber,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタ更新エラー');
    addLog('機種マスタデータ更新でエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 機種マスタデータを削除
 * @param {string} modelName - 削除する機種名
 * @param {string} manufacturer - メーカー名
 * @return {Object} 処理結果
 */
function deleteModelMasterData(modelName, manufacturer) {
  const startTime = startPerformanceTimer();
  addLog('deleteModelMasterData関数が呼び出されました', { modelName, manufacturer });
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      throw new Error('機種マスタシートが見つかりません');
    }
    
    const data = sheet.getDataRange().getValues();
    let rowIndexToDelete = -1;
    
    // 機種IDで削除する場合の処理を追加
    const isModelId = modelName && modelName.match(/^M\d{5}$/);
    
    // 該当する行を検索（ヘッダー行をスキップ）
    for (let i = 1; i < data.length; i++) {
      if (isModelId) {
        // 機種IDで検索
        if (data[i][0] === modelName) {
          rowIndexToDelete = i + 1; // シート上の行番号（1ベース）
          break;
        }
      } else {
        // 機種名とメーカーで検索
        if (data[i][1] === modelName && data[i][2] === manufacturer) {
          rowIndexToDelete = i + 1; // シート上の行番号（1ベース）
          break;
        }
      }
    }
    
    if (rowIndexToDelete === -1) {
      throw new Error('指定された機種データが見つかりません');
    }
    
    // 行を削除
    sheet.deleteRow(rowIndexToDelete);
    
    addLog('機種マスタデータ削除成功', {
      modelName: modelName,
      manufacturer: manufacturer,
      deletedRow: rowIndexToDelete,
      isModelId: isModelId
    });
    
    const responseTime = endPerformanceTimer(startTime, '機種マスタ削除');
    
    return {
      success: true,
      message: '機種データが正常に削除されました',
      deletedRow: rowIndexToDelete,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタ削除エラー');
    addLog('機種マスタデータ削除でエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * フォーム用に機種マスタデータを取得（カテゴリ名を日本語に変換）
 * @return {Object} フォーム用の機種マスタデータ
 */
function getModelMasterForForm() {
  const startTime = startPerformanceTimer();
  addLog('getModelMasterForForm関数が呼び出されました');
  
  try {
    const result = getModelMasterData();
    
    if (!result.success) {
      throw new Error(result.error);
    }
    
    // カテゴリの日本語変換マップ（英語→日本語）
    const categoryMap = {
      'desktop': 'デスクトップPC',
      'laptop': 'ノートPC',
      'server': 'サーバー',
      'printer': 'プリンタ',
      'other': 'その他'
    };
    
    // データを整形
    const models = result.data.map((row, index) => ({
      rowIndex: index,
      modelId: row[0] || '',
      modelName: row[1] || '',
      manufacturer: row[2] || '',
      category: row[3] || '',
      categoryDisplay: categoryMap[row[3]] || row[3] || '',
      createdDate: row[4] || '',
      updatedDate: row[5] || '',
      notes: row[6] || ''
    }));
    
    // カテゴリごとにグループ化
    const modelsByCategory = {};
    models.forEach(model => {
      const cat = model.category;
      if (!modelsByCategory[cat]) {
        modelsByCategory[cat] = [];
      }
      modelsByCategory[cat].push(model);
    });
    
    addLog('機種マスタフォーム用データ取得成功', {
      totalModels: models.length,
      categories: Object.keys(modelsByCategory)
    });
    
    const responseTime = endPerformanceTimer(startTime, '機種マスタフォーム用取得');
    
    return {
      success: true,
      models: models,
      modelsByCategory: modelsByCategory,
      totalCount: models.length,
      responseTime: responseTime
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタフォーム用取得エラー');
    addLog('機種マスタフォーム用データ取得でエラーが発生', {
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.toString(),
      models: [],
      modelsByCategory: {}
    };
  }
}

/**
 * 機種マスタデータを取得（getModelMasterForFormのエイリアス）
 * @return {Object} 機種マスタデータ
 */
function getModelMaster() {
  return getModelMasterForForm();
}

/**
 * カテゴリ別に機種マスタを取得
 * @param {string} category - カテゴリ名
 * @return {Object} フィルタリングされた機種マスタデータ
 */
function getModelMasterByCategory(category) {
  try {
    const result = getModelMasterForForm();
    
    if (!result.success) {
      return result;
    }
    
    const filteredModels = result.models.filter(model => 
      model.category === category
    );
    
    return {
      success: true,
      models: filteredModels,
      totalCount: filteredModels.length,
      category: category
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      models: []
    };
  }
}

/**
 * 機種マスタから利用可能なカテゴリ一覧を取得
 * @return {Object} カテゴリ一覧
 */
function getAvailableCategories() {
  try {
    const result = getModelMasterForForm();
    
    if (!result.success) {
      return {
        success: false,
        error: result.error,
        categories: []
      };
    }
    
    // カテゴリを英語に統一して取得
    const categories = [...new Set(result.models.map(model => model.category))].filter(cat => cat);
    
    // 英語カテゴリのみを返すように設定
    const englishCategories = ['desktop', 'laptop', 'server', 'printer', 'other'];
    
    return {
      success: true,
      categories: englishCategories,
      totalCount: englishCategories.length
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      categories: []
    };
  }
}

/**
 * 機種マスタのカテゴリを日本語から英語に変換
 * @return {Object} 変換結果
 */
function convertModelMasterCategoriesToEnglish() {
  const startTime = startPerformanceTimer();
  addLog('機種マスタカテゴリ英語変換開始');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません');
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      throw new Error('機種マスタシートが見つかりません');
    }
    
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    
    // カテゴリ列のインデックスを取得
    const categoryColIndex = headerRow.indexOf('カテゴリ');
    if (categoryColIndex === -1) {
      throw new Error('カテゴリ列が見つかりません');
    }
    
    // 日本語→英語カテゴリマッピング
    const categoryMapping = {
      'SV': 'server',
      'CL': 'desktop',
      'ノートブック': 'laptop',
      'ノートPC': 'laptop',
      'デスクトップPC': 'desktop',
      'サーバー': 'server',
      'プリンタ': 'printer',
      'その他': 'other'
    };
    
    let updatedCount = 0;
    const updates = [];
    
    // データ行をチェック（1行目はヘッダーなのでスキップ）
    for (let i = 1; i < data.length; i++) {
      const currentCategory = data[i][categoryColIndex];
      const englishCategory = categoryMapping[currentCategory];
      
      if (englishCategory && currentCategory !== englishCategory) {
        updates.push({
          row: i + 1,
          currentCategory: currentCategory,
          newCategory: englishCategory
        });
        data[i][categoryColIndex] = englishCategory;
        updatedCount++;
      }
    }
    
    // 更新を実行
    if (updates.length > 0) {
      // カテゴリ列のみを一括更新
      const categoryValues = data.slice(1).map(row => [row[categoryColIndex]]);
      const range = sheet.getRange(2, categoryColIndex + 1, categoryValues.length, 1);
      range.setValues(categoryValues);
      
      addLog('機種マスタカテゴリ英語変換完了', {
        updatedCount: updatedCount,
        updates: updates
      });
    }
    
    endPerformanceTimer(startTime, '機種マスタカテゴリ英語変換');
    
    return {
      success: true,
      message: `${updatedCount}個のカテゴリを英語に変換しました`,
      updatedCount: updatedCount,
      updates: updates,
      totalRows: data.length - 1
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '機種マスタカテゴリ英語変換エラー');
    addLog('機種マスタカテゴリ英語変換エラー', { error: error.toString() });
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 機種マスタの診断機能
 * @return {Object} 診断結果
 */
function diagnoseModelMaster() {
  const startTime = Date.now();
  const diagnosticLogs = [];
  
  function addDiagnosticLog(message, data = {}) {
    const logEntry = {
      timestamp: new Date().toISOString(),
      message: message,
      data: data
    };
    diagnosticLogs.push(logEntry);
    console.log(`[MODEL-MASTER-DIAGNOSTIC] ${message}`, data);
  }
  
  try {
    addDiagnosticLog('機種マスタ診断開始');
    
    // スプレッドシートID確認
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty('SPREADSHEET_ID_DESTINATION');
    
    addDiagnosticLog('スプレッドシートID確認', {
      spreadsheetId: spreadsheetId ? '設定済み' : '未設定',
      idLength: spreadsheetId ? spreadsheetId.length : 0
    });
    
    if (!spreadsheetId) {
      return {
        success: false,
        error: 'SPREADSHEET_ID_DESTINATIONが設定されていません',
        diagnostics: diagnosticLogs,
        executionTime: Date.now() - startTime + 'ms'
      };
    }
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    addDiagnosticLog('スプレッドシート接続成功', {
      spreadsheetName: spreadsheet.getName()
    });
    
    // 機種マスタシートの確認
    const sheetName = MASTER_SHEET_NAMES.model;
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    addDiagnosticLog('機種マスタシート確認', {
      sheetName: sheetName,
      sheetExists: !!sheet
    });
    
    if (!sheet) {
      return {
        success: false,
        error: `${sheetName} シートが見つかりません`,
        diagnostics: diagnosticLogs,
        executionTime: Date.now() - startTime + 'ms'
      };
    }
    
    // シートのデータ確認
    const data = sheet.getDataRange().getValues();
    const headers = data[0] || [];
    const dataRows = data.slice(1);
    
    addDiagnosticLog('シートデータ分析', {
      totalRows: data.length,
      headerRows: 1,
      dataRows: dataRows.length,
      headers: headers,
      columnCount: headers.length
    });
    
    // ヘッダー列の検証
    const expectedHeaders = ['機種ID', '機種名', 'メーカー', 'カテゴリ', '作成日', '更新日', '備考'];
    const missingHeaders = expectedHeaders.filter(h => !headers.includes(h));
    
    addDiagnosticLog('ヘッダー検証', {
      expectedHeaders: expectedHeaders,
      actualHeaders: headers,
      missingHeaders: missingHeaders,
      isValid: missingHeaders.length === 0
    });
    
    // データ品質チェック
    const dataQuality = {
      totalRows: dataRows.length,
      emptyRows: 0,
      invalidIds: [],
      missingNames: [],
      missingManufacturers: [],
      invalidCategories: [],
      invalidDates: []
    };
    
    const validCategories = ['desktop', 'laptop', 'server', 'printer', 'other'];
    const modelIdPattern = /^M\d{5}$/;
    
    dataRows.forEach((row, index) => {
      const rowNum = index + 2;
      
      // 空行チェック
      if (row.every(cell => !cell || cell === '')) {
        dataQuality.emptyRows++;
        return;
      }
      
      // 機種IDチェック
      if (!row[0] || !modelIdPattern.test(row[0])) {
        dataQuality.invalidIds.push({ row: rowNum, value: row[0] });
      }
      
      // 機種名チェック
      if (!row[1] || row[1] === '') {
        dataQuality.missingNames.push({ row: rowNum });
      }
      
      // メーカーチェック
      if (!row[2] || row[2] === '') {
        dataQuality.missingManufacturers.push({ row: rowNum });
      }
      
      // カテゴリチェック
      if (!validCategories.includes(row[3])) {
        dataQuality.invalidCategories.push({ row: rowNum, value: row[3] });
      }
    });
    
    addDiagnosticLog('データ品質チェック結果', dataQuality);
    
    // 診断結果サマリー
    const summary = {
      success: true,
      sheetStatus: 'OK',
      dataRowCount: dataRows.length - dataQuality.emptyRows,
      issues: []
    };
    
    if (missingHeaders.length > 0) {
      summary.issues.push(`ヘッダー列が不足: ${missingHeaders.join(', ')}`);
    }
    
    if (dataQuality.invalidIds.length > 0) {
      summary.issues.push(`無効な機種ID: ${dataQuality.invalidIds.length}件`);
    }
    
    if (dataQuality.missingNames.length > 0) {
      summary.issues.push(`機種名未入力: ${dataQuality.missingNames.length}件`);
    }
    
    if (dataQuality.invalidCategories.length > 0) {
      summary.issues.push(`無効なカテゴリ: ${dataQuality.invalidCategories.length}件`);
    }
    
    addDiagnosticLog('診断完了', summary);
    
    return {
      success: true,
      summary: summary,
      dataQuality: dataQuality,
      diagnostics: diagnosticLogs,
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
 * 既存の日付フォーマットを修正
 * @return {Object} 修正結果
 */
function fixExistingDateFormats() {
  const startTime = startPerformanceTimer();
  addLog('既存日付フォーマット修正開始');
  
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID_DESTINATION');
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが設定されていません');
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.model);
    
    if (!sheet) {
      throw new Error('機種マスタシートが見つかりません');
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      addLog('データがないため修正をスキップ');
      return {
        success: true,
        message: 'データがないため修正をスキップしました',
        fixedCount: 0
      };
    }
    
    // 日付列のインデックス（0ベース）
    const createdDateCol = 5; // E列
    const updatedDateCol = 6; // F列
    
    // 日付列の範囲を取得
    const dateRange1 = sheet.getRange(2, createdDateCol, lastRow - 1, 1);
    const dateRange2 = sheet.getRange(2, updatedDateCol, lastRow - 1, 1);
    
    // 現在の値を取得
    const createdDates = dateRange1.getValues();
    const updatedDates = dateRange2.getValues();
    
    let fixedCount = 0;
    const fixedData = [];
    
    // 日付値を修正
    for (let i = 0; i < createdDates.length; i++) {
      const createdDate = createdDates[i][0];
      const updatedDate = updatedDates[i][0];
      
      let newCreatedDate = createdDate;
      let newUpdatedDate = updatedDate;
      let needsFix = false;
      
      // 作成日の修正
      if (createdDate instanceof Date) {
        newCreatedDate = Utilities.formatDate(createdDate, TIMEZONE, 'yyyy/MM/dd');
        needsFix = true;
      }
      
      // 更新日の修正
      if (updatedDate instanceof Date) {
        newUpdatedDate = Utilities.formatDate(updatedDate, TIMEZONE, 'yyyy/MM/dd');
        needsFix = true;
      }
      
      if (needsFix) {
        fixedCount++;
        fixedData.push({
          row: i + 2,
          oldCreatedDate: createdDate,
          newCreatedDate: newCreatedDate,
          oldUpdatedDate: updatedDate,
          newUpdatedDate: newUpdatedDate
        });
      }
      
      createdDates[i][0] = newCreatedDate;
      updatedDates[i][0] = newUpdatedDate;
    }
    
    if (fixedCount > 0) {
      // フォーマットをテキストに設定
      dateRange1.setNumberFormat('@');
      dateRange2.setNumberFormat('@');
      
      // 修正したデータを書き込み
      dateRange1.setValues(createdDates);
      dateRange2.setValues(updatedDates);
      
      addLog('日付フォーマット修正完了', {
        fixedCount: fixedCount,
        fixedData: fixedData.slice(0, 5) // 最初の5件のみログに出力
      });
    }
    
    endPerformanceTimer(startTime, '日付フォーマット修正');
    
    return {
      success: true,
      message: `${fixedCount}件の日付フォーマットを修正しました`,
      fixedCount: fixedCount,
      details: fixedData
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '日付フォーマット修正エラー');
    addLog('日付フォーマット修正エラー', { error: error.toString() });
    
    return {
      success: false,
      error: error.toString()
    };
  }
}