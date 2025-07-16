// ========================================
// フォーム・ストレージ操作関連
// ========================================

/**
 * フォーム保存先設定を取得（拠点別対応）
 * @param {string} locationId - 拠点ID
 * @return {Object} 設定オブジェクト { locationFolder: string }
 */
function getFormStorageSettings(locationId) {
  const startTime = startPerformanceTimer();
  addLog('フォーム保存先設定取得を開始', { locationId });
  
  try {
    // 拠点IDの検証
    if (!locationId) {
      throw new Error('拠点IDが指定されていません');
    }
    
    const properties = PropertiesService.getScriptProperties();
    const settings = {
      locationFolder: properties.getProperty(`FORM_FOLDER_${locationId.toUpperCase()}`) || ''
    };
    
    endPerformanceTimer(startTime, 'フォーム保存先設定取得');
    addLog('フォーム保存先設定取得完了', { locationId, settings });
    
    return settings;
  } catch (error) {
    addLog('フォーム保存先設定取得エラー', { locationId, error: error.toString() });
    throw new Error('フォーム保存先設定の取得に失敗しました: ' + error.message);
  }
}

/**
 * フォーム保存先設定を保存（拠点別対応）
 * @param {string} locationId - 拠点ID
 * @param {Object} settings - 設定オブジェクト
 * @return {Object} 処理結果 { success: boolean, message: string }
 */
function saveFormStorageSettings(locationId, settings) {
  const startTime = startPerformanceTimer();
  addLog('フォーム保存先設定保存を開始', { locationId, settings });
  
  try {
    // バリデーション
    if (!locationId) {
      throw new Error('拠点IDが指定されていません');
    }
    
    if (!settings || typeof settings !== 'object') {
      throw new Error('設定データが不正です');
    }
    
    const properties = PropertiesService.getScriptProperties();
    const locationPrefix = `FORM_FOLDER_${locationId.toUpperCase()}`;
    
    // 拠点のフォルダ設定を保存
    if (settings.locationFolder) {
      properties.setProperty(locationPrefix, settings.locationFolder);
    } else {
      properties.deleteProperty(locationPrefix);
    }
    
    endPerformanceTimer(startTime, 'フォーム保存先設定保存');
    addLog('フォーム保存先設定保存完了', { locationId, settings });
    
    return { success: true, message: 'フォーム保存先設定が正常に保存されました' };
  } catch (error) {
    addLog('フォーム保存先設定保存エラー', { locationId, error: error.toString() });
    throw new Error('フォーム保存先設定の保存に失敗しました: ' + error.message);
  }
}

/**
 * Google Drive フォルダIDの検証
 * @param {string} folderId - フォルダID
 * @return {Object} 検証結果 { valid: boolean, error?: string, name?: string, id?: string }
 */
function validateDriveFolderId(folderId) {
  const startTime = startPerformanceTimer();
  addLog('Google Drive フォルダID検証を開始', { folderId });
  
  try {
    // フォルダIDの形式チェック
    const folderIdPattern = /^[a-zA-Z0-9_-]+$/;
    if (!folderIdPattern.test(folderId)) {
      return {
        valid: false,
        error: 'フォルダIDの形式が正しくありません'
      };
    }
    
    // 実際にフォルダにアクセスして確認
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();
    
    endPerformanceTimer(startTime, 'Google Drive フォルダID検証');
    addLog('Google Drive フォルダID検証完了', { folderId, folderName });
    
    return {
      valid: true,
      name: folderName,
      id: folderId
    };
  } catch (error) {
    endPerformanceTimer(startTime, 'Google Drive フォルダID検証エラー');
    addLog('Google Drive フォルダID検証エラー', { folderId, error: error.toString() });
    
    return {
      valid: false,
      error: 'フォルダにアクセスできません: ' + error.message
    };
  }
}

/**
 * フォーム作成時に適切な保存先フォルダIDを取得（拠点別対応・エラーハンドリング強化）
 * @param {string} locationId - 拠点ID
 * @param {string} deviceCategory - デバイスカテゴリ
 * @return {Object} フォルダID取得結果
 */
function getFormStorageFolderId(locationId, deviceCategory) {
  const startTime = startPerformanceTimer();
  addLog('フォーム保存先フォルダID取得を開始', { locationId, deviceCategory });
  
  try {
    // 拠点IDの検証
    if (!locationId) {
      throw new Error('拠点IDが指定されていません');
    }
    
    const settings = getFormStorageSettings(locationId);
    const folderId = settings.locationFolder;
    const categoryUsed = 'location';
    
    // 保存先未設定エラーチェック
    if (!folderId) {
      // 拠点マスタから拠点名を取得
      let locationName = locationId;
      try {
        const location = getLocationById(locationId);
        if (location) {
          locationName = location.locationName;
        }
      } catch (e) {
        // 拠点マスタが取得できない場合はIDをそのまま使用
      }
      
      throw new Error(`${locationName}拠点のフォーム保存先が設定されていません。設定画面で保存先を設定してください。`);
    }
    
    endPerformanceTimer(startTime, 'フォーム保存先フォルダID取得');
    addLog('フォーム保存先フォルダID取得完了', { locationId, deviceCategory, folderId, categoryUsed });
    
    return {
      success: true,
      folderId: folderId,
      locationId: locationId,
      category: deviceCategory,
      categoryUsed: categoryUsed,
      usedDefault: false
    };
  } catch (error) {
    endPerformanceTimer(startTime, 'フォーム保存先フォルダID取得エラー');
    addLog('フォーム保存先フォルダID取得エラー', { locationId, deviceCategory, error: error.toString() });
    
    return {
      success: false,
      error: error.toString(),
      locationId: locationId,
      category: deviceCategory
    };
  }
}

// ========================================
// 共通フォームURL設定関連
// ========================================

/**
 * 共通フォームURL設定を取得
 * @return {Object} 設定オブジェクト
 */
function getCommonFormsSettings() {
  const startTime = startPerformanceTimer();
  addLog('共通フォームURL設定取得開始');
  
  try {
    const properties = PropertiesService.getScriptProperties();
    
    const settings = {
      terminalCommonFormUrl: properties.getProperty('TERMINAL_COMMON_FORM_URL') || '',
      printerCommonFormUrl: properties.getProperty('PRINTER_COMMON_FORM_URL') || '',
      qrRedirectUrl: properties.getProperty('QR_REDIRECT_URL') || ''
    };
    
    endPerformanceTimer(startTime, '共通フォームURL設定取得');
    addLog('共通フォームURL設定取得完了', settings);
    
    return settings;
  } catch (error) {
    endPerformanceTimer(startTime, '共通フォームURL設定取得エラー');
    addLog('共通フォームURL設定取得エラー', { error: error.toString() });
    throw new Error('共通フォームURL設定の取得に失敗しました: ' + error.toString());
  }
}

/**
 * 共通フォームURL設定を保存
 * @param {Object} settings - 設定オブジェクト
 * @return {Object} 処理結果
 */
function saveCommonFormsSettings(settings) {
  const startTime = startPerformanceTimer();
  addLog('共通フォームURL設定保存開始', settings);
  
  try {
    const properties = PropertiesService.getScriptProperties();
    
    // 設定値を保存
    if (settings.terminalCommonFormUrl) {
      properties.setProperty('TERMINAL_COMMON_FORM_URL', settings.terminalCommonFormUrl);
    } else {
      properties.deleteProperty('TERMINAL_COMMON_FORM_URL');
    }
    
    if (settings.printerCommonFormUrl) {
      properties.setProperty('PRINTER_COMMON_FORM_URL', settings.printerCommonFormUrl);
    } else {
      properties.deleteProperty('PRINTER_COMMON_FORM_URL');
    }
    
    if (settings.qrRedirectUrl) {
      properties.setProperty('QR_REDIRECT_URL', settings.qrRedirectUrl);
    } else {
      properties.deleteProperty('QR_REDIRECT_URL');
    }
    
    endPerformanceTimer(startTime, '共通フォームURL設定保存');
    addLog('共通フォームURL設定保存完了');
    
    return {
      success: true,
      message: '共通フォームURL設定が正常に保存されました'
    };
  } catch (error) {
    endPerformanceTimer(startTime, '共通フォームURL設定保存エラー');
    addLog('共通フォームURL設定保存エラー', { error: error.toString() });
    throw new Error('共通フォームURL設定の保存に失敗しました: ' + error.toString());
  }
}

/**
 * Google FormsのURLを検証
 * @param {string} formUrl - フォームURL
 * @return {Object} 検証結果
 */
function validateGoogleFormUrl(formUrl) {
  const startTime = startPerformanceTimer();
  addLog('Google FormsのURL検証開始', { formUrl });
  
  try {
    // URLの形式チェック
    const urlPattern = /^https:\/\/docs\.google\.com\/forms\/d\/([a-zA-Z0-9_-]+)/;
    const match = formUrl.match(urlPattern);
    
    if (!match) {
      return {
        valid: false,
        error: 'Google FormsのURLの形式が正しくありません'
      };
    }
    
    const formId = match[1];
    
    try {
      // FormAppを使ってフォームにアクセスを試行
      const form = FormApp.openById(formId);
      const title = form.getTitle();
      
      endPerformanceTimer(startTime, 'Google FormsのURL検証');
      addLog('Google FormsのURL検証完了', { formId, title });
      
      return {
        valid: true,
        title: title,
        formId: formId
      };
    } catch (formError) {
      addLog('フォームアクセスエラー', { formId, error: formError.toString() });
      
      return {
        valid: false,
        error: 'フォームにアクセスできません。フォームが存在しないか、アクセス権限がありません'
      };
    }
  } catch (error) {
    endPerformanceTimer(startTime, 'Google FormsのURL検証エラー');
    addLog('Google FormsのURL検証エラー', { error: error.toString() });
    
    return {
      valid: false,
      error: 'URL検証中にエラーが発生しました: ' + error.toString()
    };
  }
}