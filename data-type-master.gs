/**
 * データタイプマスタ管理機能
 * スプレッドシートビューアーの表示形式を管理
 */

// ========================================
// ヘルパー関数（他のファイルで定義されていない場合の代替）
// ========================================

/**
 * 現在の日付を取得（他で定義されていない場合の代替）
 */
function getCurrentDateForDataType() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
}

/**
 * 日付をフォーマット（他で定義されていない場合の代替）
 */
function formatDateForDataType(dateValue) {
  if (!dateValue) return '';
  
  if (typeof dateValue === 'string') {
    return dateValue;
  }
  
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, 'Asia/Tokyo', 'yyyy/MM/dd');
  }
  
  return dateValue.toString();
}

// ========================================
// データタイプマスタ操作
// ========================================

/**
 * データタイプマスタシートを取得または作成
 * @returns {Sheet} データタイプマスタシート
 */
function getDataTypeMasterSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(MASTER_SHEET_NAMES.dataType);
  
  if (!sheet) {
    // シートが存在しない場合は作成
    sheet = spreadsheet.insertSheet(MASTER_SHEET_NAMES.dataType);
    initializeDataTypeMasterSheet(sheet);
  }
  
  return sheet;
}

/**
 * データタイプマスタシートを初期化
 * @param {Sheet} sheet - 初期化するシート
 */
function initializeDataTypeMasterSheet(sheet) {
  try {
    // ヘッダー行を設定
    const headers = [
      'データタイプID',
      'データタイプ名',
      '説明',
      '表示順序',
      'フィルタ条件',
      'データソース設定',
      '表示列設定',
      'ステータス',
      '作成日',
      '更新日'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#f0f0f0');
    
    // 初期データを設定
    const initialData = getInitialDataTypes();
    if (initialData.length > 0) {
      sheet.getRange(2, 1, initialData.length, headers.length).setValues(initialData);
    }
    
    // 列幅を調整
    sheet.setColumnWidth(1, 150); // データタイプID
    sheet.setColumnWidth(2, 150); // データタイプ名
    sheet.setColumnWidth(3, 300); // 説明
    sheet.setColumnWidth(4, 100); // 表示順序
    sheet.setColumnWidth(5, 200); // フィルタ条件
    sheet.setColumnWidth(6, 400); // データソース設定
    sheet.setColumnWidth(7, 400); // 表示列設定
    sheet.setColumnWidth(8, 100); // ステータス
    sheet.setColumnWidth(9, 120); // 作成日
    sheet.setColumnWidth(10, 120); // 更新日
    
    // 日付列の書式を設定
    sheet.getRange(2, 9, sheet.getMaxRows() - 1, 2).setNumberFormat('@');
    
    console.log('データタイプマスタシートを初期化しました');
  } catch (error) {
    console.error('データタイプマスタシート初期化エラー:', error);
    throw error;
  }
}

/**
 * 初期データタイプ設定を取得
 * @returns {Array<Array>} 初期データ
 */
function getInitialDataTypes() {
  const now = getCurrentDateForDataType();
  
  return [
    [
      'NORMAL',
      '通常データ',
      '日常的な機器管理データ表示',
      1,
      '',
      JSON.stringify({
        sourceSheets: ['端末マスタ', 'プリンタマスタ', 'その他マスタ'],
        includeAllColumns: true
      }),
      JSON.stringify({
        columns: ['拠点管理番号', '機器種別', '機種名', '製造番号', 'ステータス', '更新日時']
      }),
      'active',
      now,
      now
    ],
    [
      'AUDIT',
      '監査データ',
      '監査用の詳細データ表示',
      2,
      '',
      JSON.stringify({
        sourceSheets: ['端末ステータス収集', 'プリンタステータス収集', 'その他ステータス収集'],
        includeHistory: true
      }),
      JSON.stringify({
        columns: ['タイムスタンプ', '拠点管理番号', '担当者', '機器種別', '機種名', 'ステータス', '変更理由']
      }),
      'active',
      now,
      now
    ],
    [
      'SUMMARY',
      'サマリーデータ',
      '集計・統計データ表示',
      3,
      '',
      JSON.stringify({
        sourceSheets: ['端末マスタ', 'プリンタマスタ', 'その他マスタ'],
        aggregation: true
      }),
      JSON.stringify({
        columns: ['拠点', '機器種別', '総数', '貸出中', '返却済み', '保管中', 'メンテナンス中']
      }),
      'active',
      now,
      now
    ]
  ];
}

/**
 * データタイプマスタを取得
 * @param {boolean} activeOnly - アクティブなデータのみ取得するか
 * @returns {Object} データタイプ情報
 */
function getDataTypeMaster(activeOnly = true) {
  try {
    const sheet = getDataTypeMasterSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      return {
        success: true,
        dataTypes: []
      };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const dataTypes = [];
    
    data.forEach((row, index) => {
      const [
        dataTypeId,
        dataTypeName,
        description,
        displayOrder,
        filterCondition,
        dataSourceConfig,
        displayColumnConfig,
        status,
        createdAt,
        updatedAt
      ] = row;
      
      // 空行はスキップ
      if (!dataTypeId) return;
      
      // アクティブのみの場合、非アクティブはスキップ
      if (activeOnly && status !== 'active') return;
      
      dataTypes.push({
        dataTypeId,
        dataTypeName,
        description,
        displayOrder: displayOrder || index + 1,
        filterCondition,
        dataSourceConfig: tryParseJSON(dataSourceConfig),
        displayColumnConfig: tryParseJSON(displayColumnConfig),
        status,
        createdAt: formatDateForDataType(createdAt),
        updatedAt: formatDateForDataType(updatedAt)
      });
    });
    
    // 表示順序でソート
    dataTypes.sort((a, b) => a.displayOrder - b.displayOrder);
    
    return {
      success: true,
      dataTypes
    };
  } catch (error) {
    console.error('getDataTypeMaster エラー:', error);
    // createErrorResponse関数が定義されていない可能性があるため、直接エラーレスポンスを返す
    return {
      success: false,
      error: error.message || 'データタイプマスタの取得に失敗しました',
      dataTypes: []
    };
  }
}

/**
 * データタイプを追加
 * @param {Object} dataTypeData - データタイプ情報
 * @returns {Object} 処理結果
 */
function addDataType(dataTypeData) {
  try {
    // 必須フィールドのチェック
    if (!dataTypeData.dataTypeId || !dataTypeData.dataTypeName) {
      throw new Error('データタイプIDとデータタイプ名は必須です');
    }
    
    const sheet = getDataTypeMasterSheet();
    const lastRow = sheet.getLastRow();
    
    // 重複チェック
    if (lastRow >= 2) {
      const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      if (existingIds.includes(dataTypeData.dataTypeId)) {
        throw new Error(
          'このデータタイプIDは既に存在します',
          ErrorTypes.VALIDATION,
          { dataTypeId: dataTypeData.dataTypeId }
        );
      }
    }
    
    // データタイプIDの形式チェック
    if (!/^[A-Z_]+$/.test(dataTypeData.dataTypeId)) {
      throw new AppError(
        'データタイプIDは大文字英字とアンダースコアのみ使用可能です',
        ErrorTypes.VALIDATION
      );
    }
    
    const now = getCurrentDateForDataType();
    const newRow = [
      dataTypeData.dataTypeId,
      dataTypeData.dataTypeName,
      dataTypeData.description || '',
      dataTypeData.displayOrder || lastRow,
      dataTypeData.filterCondition || '',
      JSON.stringify(dataTypeData.dataSourceConfig || {}),
      JSON.stringify(dataTypeData.displayColumnConfig || {}),
      'active',
      now,
      now
    ];
    
    sheet.appendRow(newRow);
    
    console.log(null, 'データタイプを追加しました', { dataTypeId: dataTypeData.dataTypeId });
    
    return {
      success: true,
      message: 'データタイプを追加しました',
      dataTypeId: dataTypeData.dataTypeId
    };
  } catch (error) {
    console.error('addDataType エラー:', error);
    return { success: false, error: error.message || 'データタイプの追加に失敗しました' };
  }
}

/**
 * データタイプを更新
 * @param {string} dataTypeId - データタイプID
 * @param {Object} dataTypeData - 更新するデータ
 * @returns {Object} 処理結果
 */
function updateDataType(dataTypeId, dataTypeData) {
  try {
    const sheet = getDataTypeMasterSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      throw new Error('データタイプが見つかりません', ErrorTypes.DATA_ACCESS);
    }
    
    const dataTypeIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const rowIndex = dataTypeIds.indexOf(dataTypeId);
    
    if (rowIndex === -1) {
      throw new Error(
        'データタイプが見つかりません',
        ErrorTypes.DATA_ACCESS,
        { dataTypeId }
      );
    }
    
    const actualRow = rowIndex + 2;
    const now = getCurrentDateForDataType();
    
    // 更新可能なフィールドのみ更新
    if (dataTypeData.dataTypeName !== undefined) {
      sheet.getRange(actualRow, 2).setValue(dataTypeData.dataTypeName);
    }
    if (dataTypeData.description !== undefined) {
      sheet.getRange(actualRow, 3).setValue(dataTypeData.description);
    }
    if (dataTypeData.displayOrder !== undefined) {
      sheet.getRange(actualRow, 4).setValue(dataTypeData.displayOrder);
    }
    if (dataTypeData.filterCondition !== undefined) {
      sheet.getRange(actualRow, 5).setValue(dataTypeData.filterCondition);
    }
    if (dataTypeData.dataSourceConfig !== undefined) {
      sheet.getRange(actualRow, 6).setValue(JSON.stringify(dataTypeData.dataSourceConfig));
    }
    if (dataTypeData.displayColumnConfig !== undefined) {
      sheet.getRange(actualRow, 7).setValue(JSON.stringify(dataTypeData.displayColumnConfig));
    }
    if (dataTypeData.status !== undefined) {
      sheet.getRange(actualRow, 8).setValue(dataTypeData.status);
    }
    
    // 更新日を更新
    sheet.getRange(actualRow, 10).setValue(now);
    
    console.log(null, 'データタイプを更新しました', { dataTypeId });
    
    return {
      success: true,
      message: 'データタイプを更新しました',
      dataTypeId
    };
  } catch (error) {
    console.log(error, 'updateDataType', { dataTypeId, dataTypeData });
    return { success: false, error: error.message || 'データタイプの更新に失敗しました' };
  }
}

/**
 * データタイプを削除（論理削除）
 * @param {string} dataTypeId - データタイプID
 * @returns {Object} 処理結果
 */
function deleteDataType(dataTypeId) {
  try {
    // 論理削除（ステータスを inactive に変更）
    return updateDataType(dataTypeId, { status: 'inactive' });
  } catch (error) {
    console.log(error, 'deleteDataType', { dataTypeId });
    return { success: false, error: error.message || 'データタイプの削除に失敗しました' };
  }
}

/**
 * JSON文字列を安全にパース
 * @param {string} jsonString - JSON文字列
 * @returns {Object} パースされたオブジェクトまたは空オブジェクト
 */
function tryParseJSON(jsonString) {
  if (!jsonString) return {};
  
  try {
    return JSON.parse(jsonString);
  } catch (error) {
    console.log('JSON パースエラー:', { jsonString, error: error.toString() });
    return {};
  }
}

/**
 * データタイプに基づいてデータを取得
 * @param {string} dataTypeId - データタイプID
 * @param {string} location - 拠点ID
 * @returns {Object} フィルタリングされたデータ
 */
function getDataByType(dataTypeId, location) {
  try {
    const dataTypeResult = getDataTypeMaster(true);
    if (!dataTypeResult.success) {
      return dataTypeResult;
    }
    
    const dataType = dataTypeResult.dataTypes.find(dt => dt.dataTypeId === dataTypeId);
    if (!dataType) {
      throw new Error(
        'データタイプが見つかりません',
        ErrorTypes.DATA_ACCESS,
        { dataTypeId }
      );
    }
    
    // データソース設定に基づいてデータを取得
    const sourceConfig = dataType.dataSourceConfig;
    const displayConfig = dataType.displayColumnConfig;
    
    // 実際のデータ取得処理
    // getSpreadsheetData関数を拡張して、データタイプに応じた処理を行う
    return getSpreadsheetDataWithType(location, dataType);
  } catch (error) {
    console.log(error, 'getDataByType', { dataTypeId, location });
    return { success: false, error: error.message || 'データの取得に失敗しました' };
  }
}

/**
 * データタイプ設定に基づいてスプレッドシートデータを取得
 * @param {string} location - 拠点ID
 * @param {Object} dataType - データタイプ設定
 * @returns {Object} 処理結果
 */
function getSpreadsheetDataWithType(location, dataType) {
  // この関数は getSpreadsheetData 関数を拡張する形で実装
  // データタイプの設定に基づいて、適切なデータソースから
  // 必要な列のみを取得し、フィルタリングを適用する
  
  // 一旦、既存の getSpreadsheetData を呼び出す
  const result = getSpreadsheetData(location);
  
  if (!result.success) {
    return result;
  }
  
  // データタイプの表示列設定に基づいてフィルタリング
  if (dataType.displayColumnConfig && dataType.displayColumnConfig.columns) {
    // ここで列のフィルタリング処理を実装
    // TODO: 実装を追加
  }
  
  return result;
}