// ========================================
// Google Forms 動的作成機能
// form-creation.gs
// ========================================

// フォーム作成用の定数
const FORM_FOLDER_NAME = '代替機管理システム_自動生成フォーム';

// QRコード保存先フォルダID（スクリプトプロパティで管理）
const QR_CODE_FOLDER_KEY = 'QR_CODE_FOLDER_ID';

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
    
    // fullURLを生成（確認メッセージ用）
    let fullUrl = '';
    try {
      const properties = PropertiesService.getScriptProperties();
      const deviceKey = `form_base_${formConfig.deviceType}`;
      const urlKey = FORM_BASE_DEVICE_URL_KEYS[deviceKey];
      
      if (urlKey) {
        const deviceUrl = properties.getProperty(urlKey);
        if (deviceUrl) {
          fullUrl = `${deviceUrl}${formConfig.locationNumber || ''}`;
        }
      }
    } catch (urlError) {
      addFormLog('fullURL生成エラー', { error: urlError.toString() });
    }
    
    // カスタム確認メッセージを生成（回答送信後用）
    const confirmationMessage = 'ご回答いただき、ありがとうございました。';
    
    // 回答受付無効時のカスタムメッセージを生成
    const closedFormMessage = fullUrl 
      ? `このフォームは現在回答を受け付けていません。\n\n${fullUrl}\n\nこちらから回答をお願いします。`
      : 'このフォームは現在回答を受け付けていません。\n\nこちらから回答をお願いします。';
    
    // フォームの基本設定
    form.setDescription(formDescription);
    form.setCollectEmail(false);  // メールアドレスを収集しない
    form.setAllowResponseEdits(false);
    form.setShowLinkToRespondAgain(false);
    form.setAcceptingResponses(false);  // 回答を受付しない
    form.setConfirmationMessage(confirmationMessage);
    
    // 回答受付無効時のメッセージ設定（エラーハンドリング付き）
    let customMessageSuccess = false;
    try {
      form.setCustomClosedFormMessage(closedFormMessage);
      addFormLog('カスタム回答受付無効メッセージ設定成功', { messageLength: closedFormMessage.length });
      customMessageSuccess = true;
    } catch (customMessageError) {
      // setCustomClosedFormMessageが失敗した場合はログに記録してスキップ
      addFormLog('カスタム回答受付無効メッセージ設定失敗', { 
        error: customMessageError.toString(),
        fallbackMessage: '代替策を実行' 
      });
      customMessageSuccess = false;
    }
    
    // setCustomClosedFormMessageが失敗した場合の代替策
    if (!customMessageSuccess && fullUrl) {
      // 方法1: フォーム説明にfullURLを追加
      const enhancedDescription = formDescription + `\n\n【重要】回答はこちらから: ${fullUrl}`;
      form.setDescription(enhancedDescription);
      addFormLog('代替策1: フォーム説明にURL追加', { enhancedDescription });
      
      // 方法2: フォームタイトルの先頭にURL情報を追加
      const originalTitle = form.getTitle();
      const enhancedTitle = `${originalTitle} - 回答URL(下記より回答をお願いします。): ${fullUrl}`;
      form.setTitle(enhancedTitle);
      addFormLog('代替策2: フォームタイトルにURL追加', { enhancedTitle });
      
      // 方法3: フォームに情報表示用のセクションヘッダーを追加
      const infoSection = form.addSectionHeaderItem();
      infoSection.setTitle('【回答URL】');
      infoSection.setHelpText(`回答はこちらのURLから行ってください:\n${fullUrl}\n\nこちらから回答をお願いします。`);
      addFormLog('代替策3: 情報セクション追加完了', { fullUrl });
    }
    
    addFormLog('フォーム確認メッセージ設定', {
      fullUrl: fullUrl,
      confirmationMessageLength: confirmationMessage.length,
      closedFormMessageLength: closedFormMessage.length
    });
    
    // フォームファイルを適切なフォルダに移動
    const formFile = DriveApp.getFileById(form.getId());
    folder.addFile(formFile);
    DriveApp.getRootFolder().removeFile(formFile);
    
    // QRコード生成（作成されたフォームのURL使用）
    let qrCodeResult = null;
    try {
      // 作成されたフォームの公開URLを取得
      const formPublicUrl = form.getPublishedUrl();
      
      addFormLog('フォーム公開URL取得', {
        formId: form.getId(),
        publicUrl: formPublicUrl,
        locationNumber: formConfig.locationNumber
      });
      
      // 高品質QRコードを生成（フォームの公開URLを使用）
      const qrCodeBlob = generateHighQualityQRCode(formPublicUrl, {
        size: '600x600',     // 大きなサイズ
        ecc: 'H',           // 最高エラー訂正レベル
        margin: '30'        // 大きなマージン
      });
      
      // QRコードを保存
      qrCodeResult = saveQRCode(qrCodeBlob, formConfig.locationNumber);
      
      if (qrCodeResult.success) {
        addFormLog('QRコード生成・保存完了', {
          formUrl: formPublicUrl,
          fileName: qrCodeResult.fileName,
          fileId: qrCodeResult.fileId,
          locationNumber: formConfig.locationNumber
        });
      }
      
    } catch (qrError) {
      addFormLog('QRコード生成処理でエラー', {
        error: qrError.toString(),
        locationNumber: formConfig.locationNumber,
        formId: form.getId()
      });
      // QRコード生成エラーはフォーム作成自体は継続
    }
    
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
        collectEmail: false,
        fullUrl: fullUrl, // 確認メッセージに表示されるfullURL
        confirmationMessage: confirmationMessage, // 設定された確認メッセージ
        customClosedFormMessage: closedFormMessage, // 回答受付無効時のメッセージ
        customClosedFormMessageSuccess: customMessageSuccess, // カスタムメッセージ設定の成功可否
        fallbackMethodsUsed: !customMessageSuccess && fullUrl, // 代替策が使用されたかどうか
        qrCode: qrCodeResult ? {
          success: qrCodeResult.success,
          fileId: qrCodeResult.fileId,
          fileName: qrCodeResult.fileName,
          fileUrl: qrCodeResult.fileUrl,
          folderId: qrCodeResult.folderId,
          qrCodeUrl: form.getPublishedUrl() // QRコードが指すフォームURL
        } : { success: false, error: 'QRコード生成に失敗しました' }
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
 * 拠点管理番号のバリデーション（プレフィックスのみ検証）
 * @param {string} locationNumber - 拠点管理番号
 */
function validateLocationNumber(locationNumber) {
  if (!locationNumber || typeof locationNumber !== 'string') {
    return { valid: false, error: '拠点管理番号が無効です' };
  }
  
  // アンダースコアで分割
  const parts = locationNumber.split('_');
  if (parts.length < 2) {
    return { 
      valid: false, 
      error: '拠点管理番号は「プレフィックス_サフィックス」の形式で入力してください（例：Osaka_001）' 
    };
  }
  
  // プレフィックス（アンダースコア前の部分）のみを検証
  const prefix = parts[0];
  
  // プレフィックスは英字のみ許可
  const prefixPattern = /^[A-Za-z]+$/;
  if (!prefixPattern.test(prefix)) {
    return { 
      valid: false, 
      error: 'プレフィックスは英字のみで入力してください（例：Osaka、Hime）' 
    };
  }
  
  // プレフィックスの長さチェック（2文字以上）
  if (prefix.length < 2) {
    return { 
      valid: false, 
      error: 'プレフィックスは2文字以上で入力してください' 
    };
  }
  
  // サフィックスは検証しない（どんな文字列でも許可）
  return { valid: true };
}

/**
 * QRコードを生成する関数（QR Server API使用）
 * @param {string} text - エンコードするテキスト
 * @returns {Blob} QRコードの画像Blob
 */
function generateQRCode(text) {
  try {
    addFormLog('QRコード生成開始', { text });
    
    // QR Server APIを使用してQRコードを生成
    // https://api.qrserver.com/v1/create-qr-code/
    const qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/';
    const params = {
      'size': '400x400',        // サイズ（400x400ピクセル）
      'data': text,             // エンコードするデータ
      'format': 'png',          // 出力形式
      'ecc': 'M',              // エラー訂正レベル（L, M, Q, H）
      'margin': '10',          // マージン（ピクセル）
      'qzone': '2',            // Quiet Zone
      'bgcolor': 'FFFFFF',     // 背景色（白）
      'color': '000000'        // 前景色（黒）
    };
    
    // URLパラメータを構築
    const queryParams = Object.keys(params)
      .map(key => `${key}=${encodeURIComponent(params[key])}`)
      .join('&');
    
    const fullUrl = `${qrApiUrl}?${queryParams}`;
    
    addFormLog('QRコードAPI呼び出し', { 
      url: fullUrl.substring(0, 100) + '...', 
      textLength: text.length 
    });
    
    // APIを呼び出してQRコードを取得
    const response = UrlFetchApp.fetch(fullUrl, {
      method: 'GET',
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Google Apps Script QR Generator'
      }
    });
    
    // レスポンスの確認
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      throw new Error(`QR Server API エラー: HTTP ${responseCode} - ${response.getContentText()}`);
    }
    
    const blob = response.getBlob();
    blob.setName('qrcode.png');
    
    addFormLog('QRコード生成成功', { 
      textLength: text.length,
      blobSize: blob.getBytes().length,
      contentType: blob.getContentType(),
      responseCode: responseCode
    });
    
    return blob;
    
  } catch (error) {
    addFormLog('QRコード生成エラー', { error: error.toString(), text });
    throw new Error('QRコード生成に失敗しました: ' + error.message);
  }
}

/**
 * 高品質QRコード生成（オプション付き）
 * @param {string} text - エンコードするテキスト
 * @param {Object} options - オプション設定
 * @returns {Blob} QRコードの画像Blob
 */
function generateHighQualityQRCode(text, options = {}) {
  try {
    addFormLog('高品質QRコード生成開始', { text, options });
    
    // デフォルトオプション
    const defaultOptions = {
      size: '500x500',        // より大きなサイズ
      format: 'png',          // PNG形式
      ecc: 'H',              // 高エラー訂正レベル
      margin: '20',          // 大きなマージン
      qzone: '4',            // 大きなQuiet Zone
      bgcolor: 'FFFFFF',     // 白背景
      color: '000000'        // 黒前景
    };
    
    // オプションをマージ
    const finalOptions = { ...defaultOptions, ...options, data: text };
    
    const qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/';
    const queryParams = Object.keys(finalOptions)
      .map(key => `${key}=${encodeURIComponent(finalOptions[key])}`)
      .join('&');
    
    const fullUrl = `${qrApiUrl}?${queryParams}`;
    
    const response = UrlFetchApp.fetch(fullUrl, {
      method: 'GET',
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Google Apps Script QR Generator v2.0'
      }
    });
    
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      throw new Error(`QR Server API エラー: HTTP ${responseCode}`);
    }
    
    const blob = response.getBlob();
    blob.setName('qrcode_hq.png');
    
    addFormLog('高品質QRコード生成成功', { 
      textLength: text.length,
      blobSize: blob.getBytes().length,
      options: finalOptions
    });
    
    return blob;
    
  } catch (error) {
    addFormLog('高品質QRコード生成エラー', { error: error.toString(), text, options });
    throw new Error('高品質QRコード生成に失敗しました: ' + error.message);
  }
}

/**
 * QRコード保存先フォルダを取得または作成
 */
function getQRCodeFolder() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const folderId = properties.getProperty(QR_CODE_FOLDER_KEY);
    
    if (folderId) {
      try {
        const folder = DriveApp.getFolderById(folderId);
        addFormLog('QRコード保存フォルダを使用', { folderId });
        return folder;
      } catch (folderError) {
        addFormLog('指定QRコードフォルダが見つからない', { folderId, error: folderError.toString() });
      }
    }
    
    // フォルダーIDが設定されていない場合はデフォルトフォルダを使用
    addFormLog('デフォルトQRコードフォルダを使用');
    return DriveApp.getRootFolder();
    
  } catch (error) {
    addFormLog('QRコードフォルダ取得エラー', { error: error.toString() });
    return DriveApp.getRootFolder();
  }
}

/**
 * QRコードを保存する関数
 * @param {Blob} qrCodeBlob - QRコードのBlob
 * @param {string} locationNumber - 拠点管理番号
 * @returns {Object} 保存結果
 */
function saveQRCode(qrCodeBlob, locationNumber) {
  try {
    // 現在の日時を取得
    const now = new Date();
    const dateTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    
    // QRコードファイル名を生成: qr_拠点管理番号_日時
    const fileName = `qr_${locationNumber}_${dateTime}.png`;
    
    // QRコード保存先フォルダを取得
    const folder = getQRCodeFolder();
    
    // QRコードファイルを保存
    const file = folder.createFile(qrCodeBlob.setName(fileName));
    
    addFormLog('QRコード保存成功', {
      fileName,
      fileId: file.getId(),
      locationNumber,
      folderId: folder.getId()
    });
    
    return {
      success: true,
      fileId: file.getId(),
      fileName: fileName,
      fileUrl: file.getUrl(),
      folderId: folder.getId()
    };
    
  } catch (error) {
    addFormLog('QRコード保存エラー', {
      error: error.toString(),
      locationNumber
    });
    
    return {
      success: false,
      error: error.message || error.toString()
    };
  }
}

/**
 * QRコード保存先フォルダIDを設定する関数（管理者用）
 * @param {string} folderId - QRコード保存先フォルダID
 */
function setQRCodeFolderId(folderId) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    // フォルダの存在確認
    if (folderId) {
      try {
        DriveApp.getFolderById(folderId);
        properties.setProperty(QR_CODE_FOLDER_KEY, folderId);
        addFormLog('QRコード保存先フォルダーID設定', { folderId });
      } catch (folderError) {
        throw new Error('指定されたフォルダIDが無効です: ' + folderId);
      }
    } else {
      // フォルダIDをクリア
      properties.deleteProperty(QR_CODE_FOLDER_KEY);
      addFormLog('QRコード保存先フォルダーID削除');
    }
    
    return {
      success: true,
      message: folderId ? 'QRコード保存先フォルダーIDを設定しました' : 'QRコード保存先フォルダーIDを削除しました',
      folderId: folderId
    };
    
  } catch (error) {
    addFormLog('QRコード保存先フォルダーID設定エラー', {
      error: error.toString(),
      folderId
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 現在のQRコード保存先フォルダー設定を取得する関数
 */
function getQRCodeFolderSettings() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const folderId = properties.getProperty(QR_CODE_FOLDER_KEY);
    
    let folderInfo = null;
    if (folderId) {
      try {
        const folder = DriveApp.getFolderById(folderId);
        folderInfo = {
          id: folderId,
          name: folder.getName(),
          url: folder.getUrl()
        };
      } catch (folderError) {
        addFormLog('QRコードフォルダ情報取得エラー', { folderId, error: folderError.toString() });
      }
    }
    
    return {
      success: true,
      settings: {
        folderId: folderId,
        folderInfo: folderInfo
      }
    };
    
  } catch (error) {
    addFormLog('QRコードフォルダー設定取得エラー', {
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * QRコード機能のデバッグ用テスト関数
 * @param {string} testLocationNumber - テスト用拠点管理番号
 * @param {string} testDeviceType - テスト用デバイスタイプ（'terminal' または 'printer'）
 */
function debugQRCodeGeneration(testLocationNumber = 'Test_001', testDeviceType = 'terminal') {
  console.log('=== QRコード生成デバッグ開始 ===');
  const debugResults = [];
  
  try {
    // 1. スクリプトプロパティの確認
    console.log('1. スクリプトプロパティの確認...');
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    
    debugResults.push({
      step: 'スクリプトプロパティ確認',
      success: true,
      data: allProperties
    });
    console.log('スクリプトプロパティ:', allProperties);
    
    // 2. QRコード保存先フォルダの確認
    console.log('2. QRコード保存先フォルダの確認...');
    const qrFolderId = properties.getProperty(QR_CODE_FOLDER_KEY);
    console.log('QRコード保存先フォルダID:', qrFolderId);
    
    let qrFolder;
    if (qrFolderId) {
      try {
        qrFolder = DriveApp.getFolderById(qrFolderId);
        debugResults.push({
          step: 'QRコード保存先フォルダ取得',
          success: true,
          data: {
            folderId: qrFolderId,
            folderName: qrFolder.getName(),
            folderUrl: qrFolder.getUrl()
          }
        });
        console.log('QRコード保存先フォルダ情報:', {
          id: qrFolderId,
          name: qrFolder.getName(),
          url: qrFolder.getUrl()
        });
      } catch (folderError) {
        debugResults.push({
          step: 'QRコード保存先フォルダ取得',
          success: false,
          error: folderError.toString()
        });
        console.error('QRコード保存先フォルダエラー:', folderError.toString());
        qrFolder = DriveApp.getRootFolder();
        console.log('ルートフォルダを使用します');
      }
    } else {
      debugResults.push({
        step: 'QRコード保存先フォルダ設定',
        success: false,
        error: 'QRコード保存先フォルダIDが設定されていません'
      });
      console.warn('QRコード保存先フォルダIDが設定されていません。ルートフォルダを使用します。');
      qrFolder = DriveApp.getRootFolder();
    }
    
    // 3. デバイスURL取得テスト
    console.log('3. デバイスURL取得テスト...');
    const deviceKey = `form_base_${testDeviceType}`;
    const urlKey = FORM_BASE_DEVICE_URL_KEYS[deviceKey];
    console.log('デバイスキー:', deviceKey, 'URLキー:', urlKey);
    
    if (urlKey) {
      const deviceUrl = properties.getProperty(urlKey);
      console.log('デバイスURL:', deviceUrl);
      
      if (deviceUrl) {
        const fullUrl = `${deviceUrl}${testLocationNumber}`;
        console.log('完全URL:', fullUrl);
        
        debugResults.push({
          step: 'URL生成',
          success: true,
          data: {
            deviceKey,
            urlKey,
            deviceUrl,
            fullUrl
          }
        });
        
        // 4. QRコード生成テスト
        console.log('4. QRコード生成テスト...');
        try {
          const qrCodeBlob = generateQRCode(fullUrl);
          console.log('QRコード生成成功:', {
            blobSize: qrCodeBlob.getBytes().length,
            contentType: qrCodeBlob.getContentType(),
            name: qrCodeBlob.getName()
          });
          
          debugResults.push({
            step: 'QRコード生成',
            success: true,
            data: {
              blobSize: qrCodeBlob.getBytes().length,
              contentType: qrCodeBlob.getContentType()
            }
          });
          
          // 5. QRコード保存テスト
          console.log('5. QRコード保存テスト...');
          const saveResult = saveQRCode(qrCodeBlob, testLocationNumber);
          console.log('QRコード保存結果:', saveResult);
          
          debugResults.push({
            step: 'QRコード保存',
            success: saveResult.success,
            data: saveResult.success ? {
              fileId: saveResult.fileId,
              fileName: saveResult.fileName,
              fileUrl: saveResult.fileUrl,
              folderId: saveResult.folderId
            } : null,
            error: saveResult.success ? null : saveResult.error
          });
          
        } catch (qrError) {
          debugResults.push({
            step: 'QRコード生成',
            success: false,
            error: qrError.toString()
          });
          console.error('QRコード生成エラー:', qrError.toString());
        }
        
      } else {
        debugResults.push({
          step: 'デバイスURL取得',
          success: false,
          error: `デバイスURL（${urlKey}）が設定されていません`
        });
        console.error(`デバイスURL（${urlKey}）が設定されていません`);
      }
    } else {
      debugResults.push({
        step: 'デバイスキー確認',
        success: false,
        error: `未知のデバイスタイプ: ${testDeviceType}`
      });
      console.error(`未知のデバイスタイプ: ${testDeviceType}`);
    }
    
  } catch (mainError) {
    debugResults.push({
      step: 'メイン処理',
      success: false,
      error: mainError.toString()
    });
    console.error('メイン処理エラー:', mainError.toString());
  }
  
  console.log('=== QRコード生成デバッグ完了 ===');
  console.log('デバッグ結果サマリー:');
  debugResults.forEach((result, index) => {
    console.log(`${index + 1}. ${result.step}: ${result.success ? '成功' : '失敗'}`);
    if (!result.success && result.error) {
      console.log(`   エラー: ${result.error}`);
    }
  });
  
  return {
    success: true,
    message: 'QRコード生成デバッグが完了しました',
    results: debugResults,
    testParameters: {
      locationNumber: testLocationNumber,
      deviceType: testDeviceType
    }
  };
}

/**
 * 現在の設定状況を確認する関数
 */
function checkQRCodeSettings() {
  console.log('=== QRコード設定確認 ===');
  
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProperties = properties.getProperties();
    
    console.log('全スクリプトプロパティ:');
    Object.keys(allProperties).forEach(key => {
      console.log(`  ${key}: ${allProperties[key]}`);
    });
    
    // QRコード関連設定の確認
    const qrFolderId = properties.getProperty(QR_CODE_FOLDER_KEY);
    console.log(`\nQRコード保存先フォルダID (${QR_CODE_FOLDER_KEY}):`, qrFolderId);
    
    // デバイスURL設定の確認
    console.log('\nデバイスURL設定:');
    Object.keys(FORM_BASE_DEVICE_URL_KEYS).forEach(deviceKey => {
      const urlKey = FORM_BASE_DEVICE_URL_KEYS[deviceKey];
      const url = properties.getProperty(urlKey);
      console.log(`  ${deviceKey} (${urlKey}): ${url}`);
    });
    
    return {
      success: true,
      settings: {
        qrFolderId: qrFolderId,
        deviceUrls: Object.keys(FORM_BASE_DEVICE_URL_KEYS).reduce((acc, deviceKey) => {
          const urlKey = FORM_BASE_DEVICE_URL_KEYS[deviceKey];
          acc[deviceKey] = properties.getProperty(urlKey);
          return acc;
        }, {}),
        allProperties: allProperties
      }
    };
    
  } catch (error) {
    console.error('設定確認エラー:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * QRコード生成の単純テスト関数
 * @param {string} testText - テスト用テキスト
 */
function testQRCodeGeneration(testText = 'https://example.com/Test_001') {
  console.log('=== QRコード生成テスト開始 ===');
  
  try {
    console.log('テスト対象URL:', testText);
    
    // 1. QRコード生成テスト（標準品質）
    const qrBlob = generateQRCode(testText);
    
    // 1.2. 高品質QRコード生成テスト
    const hqBlob = generateHighQualityQRCode(testText);
    console.log('QRコード生成成功:', {
      size: qrBlob.getBytes().length,
      contentType: qrBlob.getContentType(),
      name: qrBlob.getName()
    });
    
    // 2. 保存テスト
    const saveResult = saveQRCode(qrBlob, 'TEST_001');
    console.log('保存結果:', saveResult);
    
    return {
      success: true,
      message: 'QRコード生成テスト完了',
      qrGenerated: true,
      saved: saveResult.success,
      fileInfo: saveResult.success ? {
        fileId: saveResult.fileId,
        fileName: saveResult.fileName,
        fileUrl: saveResult.fileUrl
      } : null,
      error: saveResult.success ? null : saveResult.error
    };
    
  } catch (error) {
    console.error('QRコード生成テストエラー:', error.toString());
    return {
      success: false,
      message: 'QRコード生成テスト失敗',
      error: error.toString()
    };
  }
}

/**
 * フォーム作成とQRコード生成の統合テスト
 * @param {string} testLocationNumber - テスト用拠点管理番号
 */
function testFormWithQRCode(testLocationNumber = 'QRTest_001') {
  console.log('=== フォーム＋QRコード統合テスト開始 ===');
  
  try {
    // QRコード保存先フォルダの確認
    const folderSettings = getQRCodeFolderSettings();
    console.log('QRコード保存先設定:', folderSettings);
    
    if (!folderSettings.success || !folderSettings.settings.folderId) {
      console.warn('QRコード保存先フォルダが設定されていません。ルートフォルダを使用します。');
    }
    
    // テスト用フォーム作成
    const formConfig = {
      title: `QRテストフォーム_${testLocationNumber}`,
      description: 'QRコード生成テスト用のフォームです',
      locationNumber: testLocationNumber,
      deviceType: 'terminal',
      location: 'osaka-desktop'
    };
    
    console.log('フォーム作成開始:', formConfig);
    
    const result = createGoogleForm(formConfig);
    
    if (result.success) {
      console.log('フォーム作成成功:', {
        formId: result.data.formId,
        title: result.data.title,
        publicUrl: result.data.publicUrl,
        qrCodeGenerated: result.data.qrCode?.success || false
      });
      
      if (result.data.qrCode?.success) {
        console.log('QRコード生成成功:', {
          fileName: result.data.qrCode.fileName,
          fileId: result.data.qrCode.fileId,
          qrCodeUrl: result.data.qrCode.qrCodeUrl,
          fileUrl: result.data.qrCode.fileUrl
        });
        
        console.log('✅ QRコードをスキャンすると以下のURLにアクセスします:');
        console.log(`📱 ${result.data.qrCode.qrCodeUrl}`);
        
        return {
          success: true,
          message: 'フォームとQRコードの作成が正常に完了しました',
          formData: {
            formId: result.data.formId,
            title: result.data.title,
            publicUrl: result.data.publicUrl
          },
          qrCodeData: {
            fileName: result.data.qrCode.fileName,
            fileId: result.data.qrCode.fileId,
            targetUrl: result.data.qrCode.qrCodeUrl,
            fileUrl: result.data.qrCode.fileUrl
          }
        };
      } else {
        console.error('QRコード生成に失敗:', result.data.qrCode?.error);
        return {
          success: false,
          message: 'フォームは作成されましたが、QRコード生成に失敗しました',
          error: result.data.qrCode?.error,
          formData: {
            formId: result.data.formId,
            publicUrl: result.data.publicUrl
          }
        };
      }
    } else {
      console.error('フォーム作成に失敗:', result.error);
      return {
        success: false,
        message: 'フォーム作成に失敗しました',
        error: result.error
      };
    }
    
  } catch (error) {
    console.error('統合テストエラー:', error.toString());
    return {
      success: false,
      message: '統合テストでエラーが発生しました',
      error: error.toString()
    };
  }
}

/**
 * フォームの確認メッセージをテストする関数
 * @param {string} testLocationNumber - テスト用拠点管理番号
 * @param {string} testDeviceType - テスト用デバイスタイプ
 */
function testFormConfirmationMessage(testLocationNumber = 'MsgTest_001', testDeviceType = 'terminal') {
  console.log('=== フォーム確認メッセージテスト開始 ===');
  
  try {
    // デバイスURL設定の確認
    const properties = PropertiesService.getScriptProperties();
    const deviceKey = `form_base_${testDeviceType}`;
    const urlKey = FORM_BASE_DEVICE_URL_KEYS[deviceKey];
    const deviceUrl = properties.getProperty(urlKey);
    
    console.log('デバイス設定確認:', {
      deviceType: testDeviceType,
      deviceKey: deviceKey,
      urlKey: urlKey,
      deviceUrl: deviceUrl
    });
    
    if (!deviceUrl) {
      console.warn(`⚠️ デバイスURL（${urlKey}）が設定されていません。`);
      console.log('設定例:');
      console.log(`setDeviceUrls({ '${testDeviceType}': 'https://your-url.com/' })`);
    }
    
    // テスト用フォーム作成
    const formConfig = {
      title: `確認メッセージテスト_${testLocationNumber}`,
      description: '確認メッセージのテスト用フォーム',
      locationNumber: testLocationNumber,
      deviceType: testDeviceType,
      location: 'osaka-desktop'
    };
    
    console.log('フォーム作成開始:', formConfig);
    
    const result = createGoogleForm(formConfig);
    
    if (result.success) {
      console.log('✅ フォーム作成成功');
      console.log('📋 フォーム情報:', {
        formId: result.data.formId,
        title: result.data.title,
        editUrl: result.data.editUrl,
        publicUrl: result.data.publicUrl
      });
      
      console.log('📄 確認メッセージ（回答送信後）:');
      console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      console.log(result.data.confirmationMessage);
      console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      
      console.log('📄 回答受付無効時メッセージ設定状況:');
      console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      console.log('カスタムメッセージ設定成功:', result.data.customClosedFormMessageSuccess);
      console.log('代替策使用:', result.data.fallbackMethodsUsed);
      console.log('設定を試行したメッセージ:', result.data.customClosedFormMessage);
      console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      
      if (result.data.fullUrl) {
        console.log('🔗 fullURL:', result.data.fullUrl);
      } else {
        console.log('⚠️ fullURLが設定されていません');
      }
      
      // フォームの設定確認
      try {
        const form = FormApp.openById(result.data.formId);
        const actualConfirmationMessage = form.getConfirmationMessage();
        const isAcceptingResponses = form.isAcceptingResponses();
        const actualDescription = form.getDescription();
        
        let actualClosedFormMessage = '';
        let customMessageAvailable = false;
        
        try {
          actualClosedFormMessage = form.getCustomClosedFormMessage();
          customMessageAvailable = true;
        } catch (getMessageError) {
          console.log('📝 カスタム回答受付無効メッセージの取得に失敗（既知の問題）');
          customMessageAvailable = false;
        }
        
        console.log('📊 フォーム設定確認:', {
          acceptingResponses: isAcceptingResponses,
          confirmationMessageMatch: actualConfirmationMessage === result.data.confirmationMessage,
          customMessageAvailable: customMessageAvailable,
          closedFormMessageMatch: customMessageAvailable ? actualClosedFormMessage === result.data.customClosedFormMessage : false,
          descriptionContainsUrl: result.data.fullUrl ? actualDescription.includes(result.data.fullUrl) : false
        });
        
        if (actualConfirmationMessage !== result.data.confirmationMessage) {
          console.log('⚠️ 確認メッセージが一致しません');
          console.log('期待値:', result.data.confirmationMessage);
          console.log('実際の値:', actualConfirmationMessage);
        }
        
        if (customMessageAvailable && actualClosedFormMessage !== result.data.customClosedFormMessage) {
          console.log('⚠️ 回答受付無効時メッセージが一致しません');
          console.log('期待値:', result.data.customClosedFormMessage);
          console.log('実際の値:', actualClosedFormMessage);
        }
        
        if (!customMessageAvailable && result.data.fullUrl) {
          console.log('📝 代替策の効果確認:');
          console.log('フォーム説明:', actualDescription);
          console.log('フォームタイトル:', form.getTitle());
          
          // セクションヘッダーの確認
          const items = form.getItems();
          const sectionHeaders = items.filter(item => item.getType() === FormApp.ItemType.SECTION_HEADER);
          if (sectionHeaders.length > 0) {
            console.log('追加されたセクションヘッダー:');
            sectionHeaders.forEach((header, index) => {
              console.log(`  ${index + 1}. ${header.getTitle()}: ${header.getHelpText()}`);
            });
          }
        }
        
      } catch (formError) {
        console.error('フォーム確認エラー:', formError.toString());
      }
      
      return {
        success: true,
        message: 'フォーム確認メッセージテスト完了',
        formData: {
          formId: result.data.formId,
          publicUrl: result.data.publicUrl,
          editUrl: result.data.editUrl
        },
        messageData: {
          fullUrl: result.data.fullUrl,
          confirmationMessage: result.data.confirmationMessage,
          customClosedFormMessage: result.data.customClosedFormMessage,
          hasDeviceUrl: !!deviceUrl
        }
      };
      
    } else {
      console.error('❌ フォーム作成失敗:', result.error);
      return {
        success: false,
        message: 'フォーム作成に失敗しました',
        error: result.error
      };
    }
    
  } catch (error) {
    console.error('確認メッセージテストエラー:', error.toString());
    return {
      success: false,
      message: 'テストでエラーが発生しました',
      error: error.toString()
    };
  }
}

/**
 * setCustomClosedFormMessageメソッドのテスト関数
 */
function testCustomClosedFormMessage() {
  console.log('=== setCustomClosedFormMessage テスト開始 ===');
  
  try {
    // テスト用フォームを作成
    const testForm = FormApp.create('CustomClosedMessageTest_' + new Date().getTime());
    console.log('テスト用フォーム作成:', testForm.getId());
    
    // 基本設定
    testForm.setAcceptingResponses(false);
    
    // カスタムメッセージを設定試行
    const testMessage = 'これはテスト用のカスタムメッセージです。\n\nhttps://example.com/test\n\nテスト用URLです。';
    
    try {
      testForm.setCustomClosedFormMessage(testMessage);
      console.log('✅ setCustomClosedFormMessage 成功');
      
      // 設定されたメッセージを取得試行
      try {
        const retrievedMessage = testForm.getCustomClosedFormMessage();
        console.log('✅ getCustomClosedFormMessage 成功');
        console.log('設定したメッセージ:', testMessage);
        console.log('取得したメッセージ:', retrievedMessage);
        console.log('メッセージ一致:', testMessage === retrievedMessage);
      } catch (getError) {
        console.log('❌ getCustomClosedFormMessage 失敗:', getError.toString());
      }
      
    } catch (setError) {
      console.log('❌ setCustomClosedFormMessage 失敗:', setError.toString());
      
      // 代替策：フォーム説明にメッセージを設定
      console.log('代替策実行：フォーム説明にメッセージを設定');
      const fallbackDescription = '回答受付停止中\n\n' + testMessage;
      testForm.setDescription(fallbackDescription);
      console.log('代替メッセージ設定完了');
    }
    
    console.log('テスト用フォームURL:', testForm.getEditUrl());
    console.log('テスト用フォーム公開URL:', testForm.getPublishedUrl());
    
    // 30秒後にテストフォームを削除
    setTimeout(() => {
      try {
        DriveApp.getFileById(testForm.getId()).setTrashed(true);
        console.log('テスト用フォームを削除しました');
      } catch (deleteError) {
        console.log('テスト用フォーム削除エラー:', deleteError.toString());
      }
    }, 30000);
    
    return {
      success: true,
      message: 'setCustomClosedFormMessage テスト完了',
      testFormId: testForm.getId(),
      testFormUrl: testForm.getEditUrl()
    };
    
  } catch (error) {
    console.error('❌ テスト全体でエラー:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 拠点管理番号検証機能のテスト関数
 */
function testLocationNumberValidation() {
  console.log('=== 拠点管理番号検証テスト開始 ===');
  
  const testCases = [
    // 有効なケース
    { input: 'Osaka_001', expected: true, description: '標準形式' },
    { input: 'Tokyo_123', expected: true, description: '標準形式2' },
    { input: 'Hime_abc', expected: true, description: 'サフィックス英字（許可）' },
    { input: 'Test_999999', expected: true, description: '長いサフィックス（許可）' },
    { input: 'AB_x', expected: true, description: '短いプレフィックス（2文字）' },
    { input: 'LongPrefix_suffix', expected: true, description: '長いプレフィックス' },
    { input: 'Mixed_123abc!@#', expected: true, description: 'サフィックス特殊文字（許可）' },
    
    // 無効なケース
    { input: 'A_123', expected: false, description: 'プレフィックス1文字（無効）' },
    { input: '123_suffix', expected: false, description: 'プレフィックス数字開始（無効）' },
    { input: 'Test123_suffix', expected: false, description: 'プレフィックス数字含む（無効）' },
    { input: 'Test-Name_suffix', expected: false, description: 'プレフィックス特殊文字（無効）' },
    { input: 'NoUnderscore', expected: false, description: 'アンダースコアなし（無効）' },
    { input: '_suffix', expected: false, description: 'プレフィックス空（無効）' },
    { input: 'prefix_', expected: true, description: 'サフィックス空（許可）' },
    { input: '', expected: false, description: '空文字列（無効）' },
    { input: null, expected: false, description: 'null（無効）' },
    { input: undefined, expected: false, description: 'undefined（無効）' },
  ];
  
  let passed = 0;
  let failed = 0;
  
  testCases.forEach((testCase, index) => {
    try {
      const result = validateLocationNumber(testCase.input);
      const actualValid = result.valid;
      
      if (actualValid === testCase.expected) {
        console.log(`✅ テスト ${index + 1}: ${testCase.description}`);
        console.log(`   入力: "${testCase.input}" -> 結果: ${actualValid}`);
        passed++;
      } else {
        console.log(`❌ テスト ${index + 1}: ${testCase.description}`);
        console.log(`   入力: "${testCase.input}"`);
        console.log(`   期待: ${testCase.expected}, 実際: ${actualValid}`);
        if (!result.valid) {
          console.log(`   エラー: ${result.error}`);
        }
        failed++;
      }
    } catch (error) {
      console.log(`💥 テスト ${index + 1}: ${testCase.description} - 例外発生`);
      console.log(`   エラー: ${error.toString()}`);
      failed++;
    }
  });
  
  console.log('=== テスト結果 ===');
  console.log(`✅ 成功: ${passed}件`);
  console.log(`❌ 失敗: ${failed}件`);
  console.log(`📊 成功率: ${((passed / (passed + failed)) * 100).toFixed(1)}%`);
  
  return {
    success: failed === 0,
    passed: passed,
    failed: failed,
    total: testCases.length,
    message: `拠点管理番号検証テスト完了: ${passed}/${testCases.length}件成功`
  };
} 