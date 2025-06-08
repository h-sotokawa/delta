// ========================================
// Google Forms 動的作成機能
// form-creation.gs
// ========================================

// フォーム作成用の定数
const FORM_FOLDER_NAME = '代替機管理システム_自動生成フォーム';

// 拠点別フォルダーID管理
const LOCATION_FOLDER_KEYS = {
  'osaka-desktop': 'OSAKA_DESKTOP_FOLDER_ID',
  'osaka-server': 'OSAKA_SERVER_FOLDER_ID',
  'kobe-terminal': 'KOBE_TERMINAL_FOLDER_ID',
  'himeji-terminal': 'HIMEJI_TERMINAL_FOLDER_ID',
  'osaka-printer': 'OSAKA_PRINTER_FOLDER_ID',
  'hyogo-printer': 'HYOGO_PRINTER_FOLDER_ID'
};

// デバイスタイプ別URL管理
const FORM_BASE_DEVICE_URL_KEYS = {
  'form_base_terminal': 'FORM_BASE_TERMINAL_DESCRIPTION_URL',
  'form_base_printer': 'FORM_BASE_PRINTER_DESCRIPTION_URL'
};

// デバッグ用ログ関数（form-creation専用）
function addFormLog(message, data = null) {
  if (DEBUG) {
    const timestamp = new Date().toISOString();
    console.log(`[FORM-CREATION ${timestamp}] ${message}`, data);
  }
}

/**
 * デバイスタイプ別のURL付きフォーム説明を生成
 * @param {string} baseDescription - 基本説明文
 * @param {string} deviceType - デバイスタイプ ('terminal' または 'printer')
 * @param {string} locationNumber - 拠点管理番号
 */
function generateFormDescription(baseDescription, deviceType, locationNumber) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    // deviceTypeを対応するキーに変換
    const deviceKey = `form_base_${deviceType}`;
    const urlKey = FORM_BASE_DEVICE_URL_KEYS[deviceKey];
    
    if (!urlKey) {
      addFormLog('未知のデバイスタイプ', { deviceType, deviceKey });
      return baseDescription || '';
    }
    
    const deviceUrl = properties.getProperty(urlKey);
    
    if (deviceUrl) {
      // URLの後ろに拠点管理番号を追加
      const fullUrl = `${deviceUrl}${locationNumber || ''}`;
      const urlSection = `\n\n詳細情報: ${fullUrl}`;
      const fullDescription = (baseDescription || '') + urlSection;
      addFormLog('URL付きフォーム説明を生成', { 
        deviceType, 
        baseUrl: deviceUrl, 
        locationNumber, 
        fullUrl 
      });
      return fullDescription;
    } else {
      addFormLog('デバイス用URLが設定されていません', { deviceType, urlKey });
      return baseDescription || '';
    }
    
  } catch (error) {
    addFormLog('フォーム説明生成エラー', {
      error: error.toString(),
      deviceType,
      baseDescription,
      locationNumber
    });
    return baseDescription || '';
  }
}

/**
 * 新しいGoogle Formを作成する関数
 * @param {Object} formConfig - フォーム設定
 * @param {string} formConfig.title - フォームタイトル
 * @param {string} formConfig.description - フォーム説明
 * @param {string} formConfig.locationNumber - 拠点管理番号
 * @param {Array} formConfig.questions - 質問設定の配列
 */
function createGoogleForm(formConfig) {
  addFormLog('createGoogleForm関数が呼び出されました', formConfig);
  
  try {
    // バリデーション
    if (!formConfig || typeof formConfig !== 'object') {
      throw new Error('フォーム設定が無効です');
    }
    
    if (!formConfig.title || !formConfig.locationNumber) {
      throw new Error('フォームタイトルと拠点管理番号は必須です');
    }
    
    // 拠点別フォルダを取得または作成
    const folder = getLocationFolder(formConfig.location);
    
    // フォームを作成
    const form = FormApp.create(formConfig.title);
    
    // デバイスタイプ別のURL付き説明文を生成（拠点管理番号付き）
    const formDescription = generateFormDescription(formConfig.description, formConfig.deviceType, formConfig.locationNumber);
    
    // フォームの基本設定
    form.setDescription(formDescription);
    form.setCollectEmail(false);  // メールアドレスを収集しない
    form.setAllowResponseEdits(false);
    form.setShowLinkToRespondAgain(false);
    form.setAcceptingResponses(false);  // 回答を受付しない
    form.setConfirmationMessage('このフォームは現在回答を受け付けていません。');
    
    // フォームファイルを適切なフォルダに移動
    const formFile = DriveApp.getFileById(form.getId());
    folder.addFile(formFile);
    DriveApp.getRootFolder().removeFile(formFile);
    
    // 質問項目は作成しない（要求に応じて削除）
    // フォームは空の状態で作成され、後で手動編集される
    
    // 回答先スプレッドシートは作成しない（要求に応じて削除）
    
    addFormLog('フォーム作成成功', {
      formId: form.getId(),
      formUrl: form.getEditUrl(),
      publicUrl: form.getPublishedUrl()
    });
    
    return {
      success: true,
      message: 'フォームを正常に作成しました（回答受付無効）',
      data: {
        formId: form.getId(),
        title: form.getTitle(),
        editUrl: form.getEditUrl(),
        publicUrl: form.getPublishedUrl(),
        folderId: folder.getId(),
        locationNumber: formConfig.locationNumber,
        acceptingResponses: false,
        collectEmail: false
      }
    };
    
  } catch (error) {
    addFormLog('createGoogleFormでエラーが発生', {
      error: error.toString(),
      formConfig
    });
    
    return {
      success: false,
      error: error.message || error.toString()
    };
  }
}





/**
 * 拠点別フォルダーを取得する関数
 * @param {string} location - 拠点識別子
 */
function getLocationFolder(location) {
  try {
    // スクリプトプロパティから拠点別フォルダーIDを取得
    const properties = PropertiesService.getScriptProperties();
    const folderKey = LOCATION_FOLDER_KEYS[location];
    
    if (folderKey) {
      const folderId = properties.getProperty(folderKey);
      
      if (folderId) {
        try {
          const folder = DriveApp.getFolderById(folderId);
          addFormLog('拠点別フォルダを使用', { location, folderId });
          return folder;
        } catch (folderError) {
          addFormLog('指定フォルダが見つからない', { location, folderId, error: folderError.toString() });
        }
      }
    }
    
    // フォルダーIDが設定されていない場合はデフォルトフォルダを使用
    addFormLog('デフォルトフォルダを使用', { location });
    return getOrCreateFormFolder();
    
  } catch (error) {
    addFormLog('拠点別フォルダ取得エラー', {
      location,
      error: error.toString()
    });
    return getOrCreateFormFolder();
  }
}

/**
 * フォーム保存用デフォルトフォルダを取得または作成
 */
function getOrCreateFormFolder() {
  try {
    // 既存フォルダを検索
    const folders = DriveApp.getFoldersByName(FORM_FOLDER_NAME);
    
    if (folders.hasNext()) {
      addFormLog('既存フォルダを使用', FORM_FOLDER_NAME);
      return folders.next();
    }
    
    // 新しいフォルダを作成
    const folder = DriveApp.createFolder(FORM_FOLDER_NAME);
    addFormLog('新しいフォルダを作成', FORM_FOLDER_NAME);
    return folder;
    
  } catch (error) {
    addFormLog('フォルダ作成エラー', {
      error: error.toString(),
      folderName: FORM_FOLDER_NAME
    });
    // エラーの場合はルートフォルダを返す
    return DriveApp.getRootFolder();
  }
}



/**
 * 作成済みフォーム一覧を取得
 */
function getCreatedFormsList() {
  addFormLog('作成済みフォーム一覧取得開始');
  
  try {
    const folder = getOrCreateFormFolder();
    const forms = [];
    
    // フォルダ内のファイルを取得
    const files = folder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      
      // Google Formsファイルのみを対象
      if (file.getMimeType() === 'application/vnd.google-apps.form') {
        try {
          const form = FormApp.openById(file.getId());
          
          forms.push({
            id: form.getId(),
            title: form.getTitle(),
            description: form.getDescription(),
            editUrl: form.getEditUrl(),
            publicUrl: form.getPublishedUrl(),
            createdDate: file.getDateCreated(),
            lastModified: file.getLastUpdated(),
            responseCount: form.getResponses().length
          });
          
        } catch (formError) {
          addFormLog('フォーム情報取得エラー', {
            fileId: file.getId(),
            fileName: file.getName(),
            error: formError.toString()
          });
        }
      }
    }
    
    // 作成日時で降順ソート
    forms.sort((a, b) => new Date(b.createdDate) - new Date(a.createdDate));
    
    addFormLog('フォーム一覧取得成功', { formCount: forms.length });
    
    return {
      success: true,
      forms: forms,
      totalCount: forms.length
    };
    
  } catch (error) {
    addFormLog('フォーム一覧取得エラー', {
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.message || error.toString(),
      forms: []
    };
  }
}

/**
 * フォームを削除する関数
 * @param {string} formId - 削除するフォームのID
 */
function deleteGoogleForm(formId) {
  addFormLog('フォーム削除開始', { formId });
  
  try {
    // フォームの存在確認
    const form = FormApp.openById(formId);
    const formTitle = form.getTitle();
    
    // 関連するスプレッドシートがある場合のみ削除（通常は作成されていない）
    const destinationId = form.getDestinationId();
    if (destinationId) {
      try {
        DriveApp.getFileById(destinationId).setTrashed(true);
        addFormLog('関連スプレッドシート削除', { destinationId });
      } catch (spreadsheetError) {
        addFormLog('スプレッドシート削除スキップ（存在しない）', {
          destinationId,
          error: spreadsheetError.toString()
        });
      }
    }
    
    // フォームファイルを削除
    DriveApp.getFileById(formId).setTrashed(true);
    
    addFormLog('フォーム削除成功', { formId, formTitle });
    
    return {
      success: true,
      message: `フォーム「${formTitle}」を削除しました`,
      deletedFormId: formId
    };
    
  } catch (error) {
    addFormLog('フォーム削除エラー', {
      formId,
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.message || error.toString()
    };
  }
}

/**
 * 拠点別フォルダーIDを設定する関数（管理者用）
 * @param {Object} folderIds - 拠点別フォルダーIDのオブジェクト
 */
function setLocationFolderIds(folderIds) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    Object.keys(folderIds).forEach(location => {
      const folderKey = LOCATION_FOLDER_KEYS[location];
      if (folderKey && folderIds[location]) {
        properties.setProperty(folderKey, folderIds[location]);
        addFormLog('拠点フォルダーID設定', { location, folderId: folderIds[location] });
      }
    });
    
    return {
      success: true,
      message: '拠点別フォルダーIDを設定しました',
      setFolders: Object.keys(folderIds).length
    };
    
  } catch (error) {
    addFormLog('拠点フォルダーID設定エラー', {
      error: error.toString(),
      folderIds
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 現在の拠点別フォルダー設定を取得する関数
 */
function getLocationFolderSettings() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const settings = {};
    
    Object.keys(LOCATION_FOLDER_KEYS).forEach(location => {
      const folderKey = LOCATION_FOLDER_KEYS[location];
      const folderId = properties.getProperty(folderKey);
      settings[location] = folderId || null;
    });
    
    return {
      success: true,
      settings: settings
    };
    
  } catch (error) {
    addFormLog('拠点フォルダー設定取得エラー', {
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * デバイスタイプ別のURL設定を管理する関数
 * @param {Object} urls - デバイスタイプ別URLのオブジェクト
 */
function setDeviceUrls(urls) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    Object.keys(urls).forEach(deviceType => {
      // deviceTypeを対応するキーに変換
      const deviceKey = `form_base_${deviceType}`;
      const urlKey = FORM_BASE_DEVICE_URL_KEYS[deviceKey];
      if (urlKey && urls[deviceType]) {
        properties.setProperty(urlKey, urls[deviceType]);
        addFormLog('デバイス用URL設定', { deviceType, url: urls[deviceType] });
      }
    });
    
    return {
      success: true,
      message: 'デバイス別URLを設定しました',
      setUrls: Object.keys(urls).length
    };
    
  } catch (error) {
    addFormLog('デバイス用URL設定エラー', {
      error: error.toString(),
      urls
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 現在のデバイス別URL設定を取得する関数
 */
function getDeviceUrlSettings() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const settings = {};
    
    Object.keys(FORM_BASE_DEVICE_URL_KEYS).forEach(deviceKey => {
      const urlKey = FORM_BASE_DEVICE_URL_KEYS[deviceKey];
      const url = properties.getProperty(urlKey);
      // form_base_terminal -> terminal に変換
      const deviceType = deviceKey.replace('form_base_', '');
      settings[deviceType] = url || null;
    });
    
    return {
      success: true,
      settings: settings
    };
    
  } catch (error) {
    addFormLog('デバイス用URL設定取得エラー', {
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * フォームの詳細情報を取得
 * @param {string} formId - フォームID
 */
function getFormDetails(formId) {
  addFormLog('フォーム詳細取得開始', { formId });
  
  try {
    const form = FormApp.openById(formId);
    const items = form.getItems();
    
    const questions = items.map((item, index) => {
      return {
        index: index,
        title: item.getTitle(),
        type: item.getType().toString(),
        required: item.getHelpText ? item.getHelpText() : '',
        id: item.getId()
      };
    });
    
    const responses = form.getResponses();
    const recentResponses = responses.slice(-5).map(response => {
      return {
        timestamp: response.getTimestamp(),
        respondentEmail: response.getRespondentEmail(),
        responseCount: response.getItemResponses().length
      };
    });
    
    const details = {
      id: form.getId(),
      title: form.getTitle(),
      description: form.getDescription(),
      editUrl: form.getEditUrl(),
      publicUrl: form.getPublishedUrl(),
      collectEmail: form.collectsEmail(),
      allowResponseEdits: form.canEditResponse(),
      confirmationMessage: form.getConfirmationMessage(),
      destinationId: form.getDestinationId(),
      isAcceptingResponses: form.isAcceptingResponses(),
      questions: questions,
      totalResponses: responses.length,
      recentResponses: recentResponses
    };
    
    addFormLog('フォーム詳細取得成功', { formId, questionCount: questions.length });
    
    return {
      success: true,
      form: details
    };
    
  } catch (error) {
    addFormLog('フォーム詳細取得エラー', {
      formId,
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.message || error.toString()
    };
  }
}

/**
 * 拠点管理番号のバリデーション
 * @param {string} locationNumber - 拠点管理番号
 */
function validateLocationNumber(locationNumber) {
  if (!locationNumber || typeof locationNumber !== 'string') {
    return { valid: false, error: '拠点管理番号が無効です' };
  }
  
  // 拠点管理番号の形式チェック（例：Osaka_001、Hime_123）
  const pattern = /^[A-Za-z]+_[A-Za-z0-9]+$/;
  if (!pattern.test(locationNumber)) {
    return { 
      valid: false, 
      error: '拠点管理番号は「プレフィックス_サフィックス」の形式で入力してください（例：Osaka_001）' 
    };
  }
  
  return { valid: true };
} 