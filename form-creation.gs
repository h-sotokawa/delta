// ========================================
// Google Forms 動的作成機能
// form-creation.gs
// ========================================

// フォーム作成用の定数
const FORM_FOLDER_NAME = '代替機管理システム_自動生成フォーム';
const FORM_TEMPLATE_CONFIG = {
  backgroundColor: '#1f4e79',
  confirmationMessage: 'ご回答ありがとうございました。内容を確認次第、担当者よりご連絡いたします。'
};

// 拠点別フォルダーID管理
const LOCATION_FOLDER_KEYS = {
  'osaka-desktop': 'OSAKA_DESKTOP_FOLDER_ID',
  'osaka-server': 'OSAKA_SERVER_FOLDER_ID',
  'kobe-terminal': 'KOBE_TERMINAL_FOLDER_ID',
  'himeji-terminal': 'HIMEJI_TERMINAL_FOLDER_ID',
  'osaka-printer': 'OSAKA_PRINTER_FOLDER_ID',
  'hyogo-printer': 'HYOGO_PRINTER_FOLDER_ID'
};

// デバッグ用ログ関数（form-creation専用）
function addFormLog(message, data = null) {
  if (DEBUG) {
    const timestamp = new Date().toISOString();
    console.log(`[FORM-CREATION ${timestamp}] ${message}`, data);
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
    
    // フォームの基本設定
    form.setDescription(formConfig.description || '');
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
 * フォームに質問を追加する関数
 * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
 * @param {Object} questionConfig - 質問設定
 * @param {number} index - 質問のインデックス
 */
function addQuestionToForm(form, questionConfig, index) {
  addFormLog('質問を追加', { questionConfig, index });
  
  try {
    const { type, title, required = false, options = [], validation = null, defaultValue = '' } = questionConfig;
    
    let item;
    
    switch (type) {
      case 'TEXT':
        item = form.addTextItem()
          .setTitle(title)
          .setRequired(required);
        if (defaultValue) {
          item.setHelpText('デフォルト値: ' + defaultValue);
        }
        if (validation) {
          applyTextValidation(item, validation);
        }
        break;
        
      case 'PARAGRAPH':
        item = form.addParagraphTextItem()
          .setTitle(title)
          .setRequired(required);
        break;
        
      case 'MULTIPLE_CHOICE':
        item = form.addMultipleChoiceItem()
          .setTitle(title)
          .setRequired(required);
        if (options.length > 0) {
          item.setChoices(options.map(opt => item.createChoice(opt)));
        }
        break;
        
      case 'CHECKBOX':
        item = form.addCheckboxItem()
          .setTitle(title)
          .setRequired(required);
        if (options.length > 0) {
          item.setChoices(options.map(opt => item.createChoice(opt)));
        }
        break;
        
      case 'DROPDOWN':
        item = form.addListItem()
          .setTitle(title)
          .setRequired(required);
        if (options.length > 0) {
          const choices = options.map(opt => item.createChoice(opt));
          item.setChoices(choices);
          
          // デフォルト値が設定されている場合は先頭に配置
          if (defaultValue && options.includes(defaultValue)) {
            const defaultChoices = [
              item.createChoice(defaultValue),
              ...choices.filter(choice => choice.getValue() !== defaultValue)
            ];
            item.setChoices(defaultChoices);
          }
        }
        break;
        
      case 'SCALE':
        item = form.addScaleItem()
          .setTitle(title)
          .setRequired(required)
          .setBounds(1, 5);
        break;
        
      case 'DATE':
        item = form.addDateItem()
          .setTitle(title)
          .setRequired(required);
        break;
        
      case 'TIME':
        item = form.addTimeItem()
          .setTitle(title)
          .setRequired(required);
        break;
        
      default:
        addFormLog('未知の質問タイプ', { type, title });
        return;
    }
    
    addFormLog('質問追加成功', { type, title, index });
    
  } catch (error) {
    addFormLog('質問追加エラー', {
      error: error.toString(),
      questionConfig,
      index
    });
  }
}

/**
 * テキスト項目にバリデーションを適用
 * @param {GoogleAppsScript.Forms.TextItem} item - テキスト項目
 * @param {Object} validation - バリデーション設定
 */
function applyTextValidation(item, validation) {
  try {
    // FormApp.createTextValidation()を使用（正しいAPI）
    const builder = FormApp.createTextValidation();
    
    switch (validation.type) {
      case 'EMAIL':
        builder.requireTextIsEmail();
        break;
      case 'URL':
        builder.requireTextIsUrl();
        break;
      case 'PATTERN':
        if (validation.pattern) {
          builder.requireTextMatchesPattern(validation.pattern);
        }
        break;
      case 'LENGTH':
        if (validation.minLength !== undefined && validation.maxLength !== undefined) {
          builder.requireTextLengthRange(validation.minLength, validation.maxLength);
        }
        break;
      case 'NUMBER':
        builder.requireTextIsNumber();
        break;
    }
    
    if (validation.errorMessage) {
      builder.setHelpText(validation.errorMessage);
    }
    
    item.setValidation(builder.build());
    
  } catch (error) {
    addFormLog('バリデーション適用エラー', {
      error: error.toString(),
      validation
    });
    // バリデーション適用に失敗した場合は、ヘルプテキストのみ設定
    if (validation.errorMessage) {
      try {
        item.setHelpText(validation.errorMessage);
      } catch (helpError) {
        addFormLog('ヘルプテキスト設定エラー', helpError);
      }
    }
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
 * フォーム回答用スプレッドシートを作成
 * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
 * @param {Object} formConfig - フォーム設定
 */
function createResponseSpreadsheet(form, formConfig) {
  try {
    const spreadsheetName = `${form.getTitle()}_回答データ`;
    const spreadsheet = SpreadsheetApp.create(spreadsheetName);
    
    // フォームと回答シートを関連付け
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    
    // スプレッドシートファイルを適切なフォルダに移動
    const folder = getLocationFolder(formConfig.location);
    const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
    folder.addFile(spreadsheetFile);
    DriveApp.getRootFolder().removeFile(spreadsheetFile);
    
    // 管理用シートを追加
    const managementSheet = spreadsheet.insertSheet('管理情報');
    managementSheet.getRange('A1:B11').setValues([
      ['項目', '値'],
      ['フォーム名', form.getTitle()],
      ['フォームID', form.getId()],
      ['作成日時', new Date()],
      ['作成者', Session.getActiveUser().getEmail()],
      ['拠点管理番号', formConfig.locationNumber],
      ['拠点', formConfig.location],
      ['フォーム編集URL', form.getEditUrl()],
      ['フォーム回答URL', form.getPublishedUrl()],
      ['', ''],
      ['', '']
    ]);
    
    addFormLog('回答スプレッドシート作成成功', {
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl()
    });
    
    return spreadsheet;
    
  } catch (error) {
    addFormLog('回答スプレッドシート作成エラー', {
      error: error.toString(),
      formTitle: form.getTitle()
    });
    throw error;
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
  
  // 拠点管理番号の形式チェック（例：LOC-001）
  const pattern = /^[A-Z]{3}-\d{3}$/;
  if (!pattern.test(locationNumber)) {
    return { 
      valid: false, 
      error: '拠点管理番号は「AAA-000」の形式で入力してください（例：OSA-001）' 
    };
  }
  
  return { valid: true };
} 