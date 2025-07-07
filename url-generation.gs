// ========================================
// 共通フォームURL生成機能
// url-generation.gs
// ========================================

/**
 * 共通フォームURLに拠点管理番号を追加してURLを生成
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} deviceCategory - デバイスカテゴリ（SV/CL/プリンタ/その他）
 * @return {Object} 生成結果
 */
function generateCommonFormUrl(locationNumber, deviceCategory) {
  const startTime = startPerformanceTimer();
  addLog('URL生成開始', { locationNumber, deviceCategory });
  
  try {
    // 共通フォームURL設定を取得
    const settings = getCommonFormsSettings();
    
    // デバイスカテゴリに応じて適切な共通URLを選択
    let baseUrl;
    let formType;
    
    // 端末系のカテゴリ判定
    if (deviceCategory === 'SV' || deviceCategory === 'CL' || 
        deviceCategory === 'デスクトップPC' || deviceCategory === 'ノートPC' || 
        deviceCategory === '端末' || deviceCategory.includes('PC') || 
        deviceCategory.includes('サーバ') || deviceCategory.includes('クライアント')) {
      baseUrl = settings.terminalCommonFormUrl;
      formType = '端末';
      if (!baseUrl) {
        throw new Error('端末用共通フォームURLが設定されていません');
      }
    } else if (deviceCategory === 'プリンタ' || deviceCategory.includes('プリンタ')) {
      baseUrl = settings.printerCommonFormUrl;
      formType = 'プリンタ';
      if (!baseUrl) {
        throw new Error('プリンタ・その他用共通フォームURLが設定されていません');
      }
    } else {
      // その他のカテゴリはプリンタ・その他用として扱う
      baseUrl = settings.printerCommonFormUrl;
      formType = 'その他';
      if (!baseUrl) {
        throw new Error('プリンタ・その他用共通フォームURLが設定されていません');
      }
    }
    
    // URLパラメータとして拠点管理番号を追加
    const generatedUrl = addLocationNumberParameter(baseUrl, locationNumber);
    
    endPerformanceTimer(startTime, 'URL生成');
    addLog('URL生成完了', { locationNumber, deviceCategory, generatedUrl });
    
    return {
      success: true,
      url: generatedUrl,
      baseUrl: baseUrl,
      locationNumber: locationNumber,
      deviceCategory: deviceCategory
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'URL生成エラー');
    addLog('URL生成エラー', { locationNumber, deviceCategory, error: error.toString() });
    
    return {
      success: false,
      error: error.toString(),
      locationNumber: locationNumber,
      deviceCategory: deviceCategory
    };
  }
}

/**
 * 共通フォームURLに拠点管理番号をパラメータとして追加
 * @param {string} baseUrl - 基本URL
 * @param {string} locationNumber - 拠点管理番号
 * @return {string} パラメータ付きURL
 */
function addLocationNumberParameter(baseUrl, locationNumber) {
  try {
    // URLオブジェクトを作成
    const url = new URL(baseUrl);
    
    // 既存のパラメータがある場合は追加、ない場合は新規作成
    if (url.search) {
      // 既にパラメータがある場合
      return `${baseUrl}&entry.123456789=${encodeURIComponent(locationNumber)}`;
    } else {
      // パラメータがない場合
      const separator = baseUrl.includes('?') ? '&' : '?';
      return `${baseUrl}${separator}entry.123456789=${encodeURIComponent(locationNumber)}`;
    }
  } catch (error) {
    // URLパースエラーの場合は単純に追加
    const separator = baseUrl.includes('?') ? '&' : '?';
    return `${baseUrl}${separator}entry.123456789=${encodeURIComponent(locationNumber)}`;
  }
}

/**
 * 端末マスタに共通フォームURLを保存
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} generatedUrl - 生成されたURL
 * @return {Object} 保存結果
 */
function saveUrlToTerminalMaster(locationNumber, generatedUrl) {
  const startTime = startPerformanceTimer();
  addLog('端末マスタURL保存開始', { locationNumber, generatedUrl });
  
  try {
    // 端末マスタシートを取得
    const spreadsheet = SpreadsheetApp.openById(getMainSpreadsheetId());
    const sheet = spreadsheet.getSheetByName('端末マスタ');
    
    if (!sheet) {
      throw new Error('端末マスタシートが見つかりません');
    }
    
    // 拠点管理番号で該当行を検索
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    
    // 拠点管理番号の列インデックスを取得
    const locationNumberColIndex = headerRow.indexOf('拠点管理番号');
    if (locationNumberColIndex === -1) {
      throw new Error('端末マスタに拠点管理番号列が見つかりません');
    }
    
    // 共通フォームURL列を検索または作成
    let urlColIndex = headerRow.indexOf('共通フォームURL');
    if (urlColIndex === -1) {
      // 列が存在しない場合は追加
      urlColIndex = headerRow.length;
      sheet.getRange(1, urlColIndex + 1).setValue('共通フォームURL');
    }
    
    // 該当行を検索
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][locationNumberColIndex] === locationNumber) {
        targetRow = i + 1; // 1-based indexing
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error(`拠点管理番号 ${locationNumber} が端末マスタに見つかりません`);
    }
    
    // URLを保存
    sheet.getRange(targetRow, urlColIndex + 1).setValue(generatedUrl);
    
    endPerformanceTimer(startTime, '端末マスタURL保存');
    addLog('端末マスタURL保存完了', { locationNumber, targetRow, urlColIndex });
    
    return {
      success: true,
      savedRow: targetRow,
      savedColumn: urlColIndex + 1
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, '端末マスタURL保存エラー');
    addLog('端末マスタURL保存エラー', { locationNumber, error: error.toString() });
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * プリンタマスタに共通フォームURLを保存
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} generatedUrl - 生成されたURL
 * @return {Object} 保存結果
 */
function saveUrlToPrinterMaster(locationNumber, generatedUrl) {
  const startTime = startPerformanceTimer();
  addLog('プリンタマスタURL保存開始', { locationNumber, generatedUrl });
  
  try {
    // プリンタマスタシートを取得
    const spreadsheet = SpreadsheetApp.openById(getMainSpreadsheetId());
    const sheet = spreadsheet.getSheetByName('プリンタマスタ');
    
    if (!sheet) {
      throw new Error('プリンタマスタシートが見つかりません');
    }
    
    // 拠点管理番号で該当行を検索
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    
    // 拠点管理番号の列インデックスを取得
    const locationNumberColIndex = headerRow.indexOf('拠点管理番号');
    if (locationNumberColIndex === -1) {
      throw new Error('プリンタマスタに拠点管理番号列が見つかりません');
    }
    
    // 共通フォームURL列を検索または作成
    let urlColIndex = headerRow.indexOf('共通フォームURL');
    if (urlColIndex === -1) {
      // 列が存在しない場合は追加
      urlColIndex = headerRow.length;
      sheet.getRange(1, urlColIndex + 1).setValue('共通フォームURL');
    }
    
    // 該当行を検索
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][locationNumberColIndex] === locationNumber) {
        targetRow = i + 1; // 1-based indexing
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error(`拠点管理番号 ${locationNumber} がプリンタマスタに見つかりません`);
    }
    
    // URLを保存
    sheet.getRange(targetRow, urlColIndex + 1).setValue(generatedUrl);
    
    endPerformanceTimer(startTime, 'プリンタマスタURL保存');
    addLog('プリンタマスタURL保存完了', { locationNumber, targetRow, urlColIndex });
    
    return {
      success: true,
      savedRow: targetRow,
      savedColumn: urlColIndex + 1
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'プリンタマスタURL保存エラー');
    addLog('プリンタマスタURL保存エラー', { locationNumber, error: error.toString() });
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * URL生成と保存を一括実行
 * @param {Object} requestData - リクエストデータ
 * @return {Object} 実行結果
 */
function generateAndSaveCommonFormUrl(requestData) {
  const startTime = startPerformanceTimer();
  addLog('URL生成・保存一括処理開始', requestData);
  
  try {
    const { locationNumber, deviceCategory } = requestData;
    
    // バリデーション
    if (!locationNumber || !deviceCategory) {
      throw new Error('拠点管理番号とデバイスカテゴリは必須です');
    }
    
    // URL生成
    const urlResult = generateCommonFormUrl(locationNumber, deviceCategory);
    if (!urlResult.success) {
      throw new Error('URL生成に失敗しました: ' + urlResult.error);
    }
    
    // マスタデータに保存
    let saveResult;
    let masterType;
    
    // カテゴリに応じて保存先を決定
    if (deviceCategory === 'SV' || deviceCategory === 'CL' || 
        deviceCategory === 'デスクトップPC' || deviceCategory === 'ノートPC' || 
        deviceCategory === '端末' || deviceCategory.includes('PC') || 
        deviceCategory.includes('サーバ') || deviceCategory.includes('クライアント')) {
      saveResult = saveUrlToTerminalMaster(locationNumber, urlResult.url);
      masterType = '端末マスタ';
    } else if (deviceCategory === 'プリンタ' || deviceCategory.includes('プリンタ')) {
      saveResult = saveUrlToPrinterMaster(locationNumber, urlResult.url);
      masterType = 'プリンタマスタ';
    } else {
      // その他のカテゴリはプリンタマスタに保存
      saveResult = saveUrlToPrinterMaster(locationNumber, urlResult.url);
      masterType = 'プリンタマスタ';
    }
    
    if (!saveResult.success) {
      throw new Error('マスタデータへの保存に失敗しました: ' + saveResult.error);
    }
    
    endPerformanceTimer(startTime, 'URL生成・保存一括処理');
    addLog('URL生成・保存一括処理完了', {
      locationNumber,
      deviceCategory,
      generatedUrl: urlResult.url,
      savedRow: saveResult.savedRow
    });
    
    return {
      success: true,
      locationNumber: locationNumber,
      deviceCategory: deviceCategory,
      generatedUrl: urlResult.url,
      baseUrl: urlResult.baseUrl,
      savedTo: masterType,
      savedRow: saveResult.savedRow,
      savedColumn: saveResult.savedColumn
    };
    
  } catch (error) {
    endPerformanceTimer(startTime, 'URL生成・保存一括処理エラー');
    addLog('URL生成・保存一括処理エラー', { requestData, error: error.toString() });
    
    return {
      success: false,
      error: error.toString(),
      requestData: requestData
    };
  }
}

/**
 * メインスプレッドシートIDを取得
 * @return {string} スプレッドシートID
 */
function getMainSpreadsheetId() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID_DESTINATION');
  
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID_DESTINATIONが設定されていません');
  }
  
  return spreadsheetId;
}