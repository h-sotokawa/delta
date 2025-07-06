// ========================================
// Google Forms 動的作成機能
// form-creation.gs
// ========================================

// フォーム作成用の定数
const FORM_FOLDER_NAME = '代替機管理システム_自動生成フォーム';

// QRコード保存先フォルダID（スクリプトプロパティで管理）
const QR_CODE_FOLDER_KEY = 'QR_CODE_FOLDER_ID';

// スプレッドシート設定
const SPREADSHEET_ID_DESTINATION_KEY = 'SPREADSHEET_ID_DESTINATION';

// 拠点別フォルダーID管理
const LOCATION_FOLDER_KEYS = {
  'osaka-desktop': 'OSAKA_DESKTOP_FOLDER_ID',
  'osaka-server': 'OSAKA_SERVER_FOLDER_ID',
  'kobe-terminal': 'KOBE_TERMINAL_FOLDER_ID',
  'himeji-terminal': 'HIMEJI_TERMINAL_FOLDER_ID',
  'osaka-printer': 'OSAKA_PRINTER_FOLDER_ID',
  'hyogo-printer': 'HYOGO_PRINTER_FOLDER_ID'
};

// 拠点別シート名マッピング
const LOCATION_SHEET_NAMES = {
  'osaka-desktop': '大阪(端末)',
  'osaka-server': '大阪(端末)',
  'kobe-terminal': '神戸(端末)',
  'himeji-terminal': '姫路(端末)',
  'osaka-printer': '大阪(プリンタ)',
  'hyogo-printer': '兵庫(プリンタ)'
};

// スプレッドシート列マッピング（フィールド名 → 列名の候補リスト）
const SPREADSHEET_COLUMN_MAPPING = {
  'assetNumber': ['資産番号', 'アセット番号', 'Asset'],
  'modelNumber': ['型番', 'モデル', 'Model'],
  'serial': ['シリアル', 'シリアル番号', 'Serial', '製造番号'],
  'software': ['ソフト', 'ソフトウェア', 'Software'],
  'os': ['OS', 'オペレーティングシステム', 'Operating System'],
  'deviceType': ['代替機種別', '機種別', 'デバイス種別', 'Device Type'],
  'locationNumber': ['拠点管理番号', '拠点コード', '拠点番号', 'Location Code'],
  // 数式追加用の列
  'assignee': ['担当者'],
  'lendingStatus': ['貸出ステータス'],
  'lendingDestination': ['貸出先'],
  'lendingDate': ['貸出日'],
  'userDeviceDeposit': ['ユーザー預り機有'],
  'depositReceiptNo': ['お預かり証No.'],
  'remarks': ['備考']
};

// ステータス収集シート名の定義
const STATUS_COLLECTION_SHEETS = {
  terminal: '端末ステータス収集',  // SV, CL用
  printer: 'プリンタステータス収集'  // プリンタ, その他用
};

// ステータス収集シートの列マッピング（動的検知用）
const STATUS_SHEET_COLUMN_MAPPING = {
  'locationNumber': ['0-0.拠点管理番号'],
  'assignee': ['0-1.担当者'],
  'status': ['0-4.ステータス'],
  'destination': ['1-1.顧客名または貸出先'],
  'timestamp': ['タイムスタンプ'],
  'userDeviceDeposit': ['1-4.ユーザー機の預り有無'],
  'depositReceiptNo': ['1-8.お預かり証No.'],
  'remarks': ['1-6.備考']
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
 * ヘッダー行から列インデックスを検出する関数
 * @param {Array} headerRow - ヘッダー行データ
 * @return {Object} 列名とインデックスのマッピング
 */
function detectColumnIndexes(headerRow) {
  const columnIndexes = {};
  
  if (!headerRow || !Array.isArray(headerRow)) {
    addFormLog('ヘッダー行が無効', { headerRow });
    return columnIndexes;
  }
  
  // 各フィールドの列インデックスを検索
  Object.keys(SPREADSHEET_COLUMN_MAPPING).forEach(fieldName => {
    const candidateNames = SPREADSHEET_COLUMN_MAPPING[fieldName];
    let foundIndex = -1;
    
    for (let i = 0; i < headerRow.length; i++) {
      const headerValue = headerRow[i] ? headerRow[i].toString().trim() : '';
      
      // 候補名との一致を確認
      for (const candidateName of candidateNames) {
        if (headerValue === candidateName) {
          foundIndex = i;
          break;
        }
      }
      
      if (foundIndex !== -1) break;
    }
    
    if (foundIndex !== -1) {
      columnIndexes[fieldName] = foundIndex;
    }
  });
  
  addFormLog('列インデックス検出結果', {
    detectedColumns: columnIndexes,
    headerRowLength: headerRow.length,
    totalMappings: Object.keys(SPREADSHEET_COLUMN_MAPPING).length
  });
  
  return columnIndexes;
}

/**
 * ヘッダー行を検出する関数
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 * @return {Object} ヘッダー情報
 */
function detectHeaderRow(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  
  if (lastRow === 0 || lastColumn === 0) {
    return {
      found: false,
      error: 'シートにデータがありません'
    };
  }
  
  // 1行目と3行目をチェック
  const candidateRows = [1, 3].filter(rowNum => rowNum <= lastRow);
  let bestHeaderResult = null;
  let bestScore = -1;
  
  for (const rowNum of candidateRows) {
    const rowData = sheet.getRange(rowNum, 1, 1, lastColumn).getValues()[0];
    const columnIndexes = detectColumnIndexes(rowData);
    const score = Object.keys(columnIndexes).length;
    
    if (score > bestScore) {
      bestScore = score;
      bestHeaderResult = {
        rowIndex: rowNum,
        rowData: rowData,
        columnIndexes: columnIndexes,
        score: score
      };
    }
  }
  
  if (bestHeaderResult && bestScore > 0) {
    addFormLog('ヘッダー行検出成功', {
      rowIndex: bestHeaderResult.rowIndex,
      score: bestScore,
      detectedColumns: Object.keys(bestHeaderResult.columnIndexes)
    });
    
    return {
      found: true,
      ...bestHeaderResult
    };
  } else {
    addFormLog('ヘッダー行検出失敗', {
      candidateRows,
      lastRow,
      lastColumn
    });
    
    return {
      found: false,
      error: '有効なヘッダー行が見つかりませんでした'
    };
  }
}

/**
 * デバイスタイプに応じて必要なフィールドを取得する関数
 * @param {string} deviceType - デバイスタイプ（SV, CL, プリンタ, その他）
 * @return {Array} 必要なフィールド名の配列
 */
function getRequiredFieldsForDeviceType(deviceType) {
  const allFields = Object.keys(SPREADSHEET_COLUMN_MAPPING);
  
  // プリンタ・その他の場合は特定の列を除外
  if (deviceType === 'プリンタ' || deviceType === 'その他') {
    const excludedFields = ['assetNumber', 'software', 'os', 'depositReceiptNo'];
    const requiredFields = allFields.filter(field => !excludedFields.includes(field));
    
    addFormLog('プリンタ・その他用フィールド設定', {
      deviceType,
      allFields: allFields.length,
      excludedFields: excludedFields,
      requiredFields: requiredFields.length,
      excludedCount: excludedFields.length
    });
    
    return requiredFields;
  }
  
  // 端末（SV, CL）の場合はすべてのフィールドが必要
  addFormLog('端末用フィールド設定', {
    deviceType,
    requiredFields: allFields.length,
    allFieldsIncluded: true
  });
  
  return allFields;
}

/**
 * 不足しているヘッダーを検出する関数
 * @param {Object} currentColumnIndexes - 現在検出されている列インデックス
 * @param {string} deviceType - デバイスタイプ（SV, CL, プリンタ, その他）
 * @return {Array} 不足しているフィールド名の配列
 */
function detectMissingHeaders(currentColumnIndexes, deviceType = null) {
  const requiredFields = deviceType ? 
    getRequiredFieldsForDeviceType(deviceType) : 
    Object.keys(SPREADSHEET_COLUMN_MAPPING);
    
  const missingFields = requiredFields.filter(fieldName => 
    currentColumnIndexes[fieldName] === undefined
  );
  
  addFormLog('不足ヘッダー検出', {
    deviceType: deviceType || '未指定',
    requiredFields: requiredFields.length,
    currentFields: Object.keys(currentColumnIndexes).length,
    missingFields: missingFields,
    excludedForDevice: deviceType && (deviceType === 'プリンタ' || deviceType === 'その他')
  });
  
  return missingFields;
}

/**
 * シートの3行目に不足しているヘッダーを追加する関数
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 * @param {Array} missingFields - 不足しているフィールド名の配列
 * @param {string} deviceType - デバイスタイプ（ログ用）
 * @return {Object} 追加結果
 */
function addMissingHeadersToSheet(sheet, missingFields, deviceType = null) {
  try {
    if (!missingFields || missingFields.length === 0) {
      return {
        success: true,
        added: 0,
        message: '追加するヘッダーはありません'
      };
    }
    
    const HEADER_ROW = 3; // ヘッダーは3行目
    let currentLastColumn = sheet.getLastColumn();
    
    // 3行目が存在しない場合は行を追加
    const currentLastRow = sheet.getLastRow();
    if (currentLastRow < HEADER_ROW) {
      // 3行目まで行を追加
      for (let i = currentLastRow + 1; i <= HEADER_ROW; i++) {
        sheet.insertRowAfter(sheet.getLastRow());
      }
      addFormLog('行を追加してヘッダー行を確保', {
        previousLastRow: currentLastRow,
        newLastRow: sheet.getLastRow(),
        headerRow: HEADER_ROW
      });
    }
    
    const addedHeaders = [];
    let columnPosition = currentLastColumn + 1;
    
    // 各不足フィールドのヘッダーを追加
    missingFields.forEach(fieldName => {
      const candidateNames = SPREADSHEET_COLUMN_MAPPING[fieldName];
      if (candidateNames && candidateNames.length > 0) {
        // 最初の候補名を使用
        const headerName = candidateNames[0];
        
        // 3行目の指定列にヘッダーを設定
        sheet.getRange(HEADER_ROW, columnPosition).setValue(headerName);
        
        addedHeaders.push({
          fieldName: fieldName,
          headerName: headerName,
          column: columnPosition,
          columnLetter: getColumnLetter(columnPosition)
        });
        
        addFormLog('ヘッダー追加', {
          fieldName,
          headerName,
          column: columnPosition,
          columnLetter: getColumnLetter(columnPosition),
          position: `${getColumnLetter(columnPosition)}${HEADER_ROW}`
        });
        
        columnPosition++;
      }
    });
    
    // 追加後に列数を更新
    const newLastColumn = sheet.getLastColumn();
    
    addFormLog('ヘッダー追加完了', {
      sheetName: sheet.getName(),
      headerRow: HEADER_ROW,
      previousLastColumn: currentLastColumn,
      newLastColumn: newLastColumn,
      addedCount: addedHeaders.length,
      addedHeaders: addedHeaders
    });
    
    return {
      success: true,
      added: addedHeaders.length,
      addedHeaders: addedHeaders,
      headerRow: HEADER_ROW,
      newLastColumn: newLastColumn,
      message: `${addedHeaders.length}個のヘッダーを3行目に追加しました`
    };
    
  } catch (error) {
    addFormLog('ヘッダー追加エラー', {
      error: error.toString(),
      sheetName: sheet.getName(),
      missingFields
    });
    
    return {
      success: false,
      error: error.message,
      added: 0
    };
  }
}

/**
 * ヘッダー行を検出し、不足している場合は自動追加する関数
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - シートオブジェクト
 * @param {string} deviceType - デバイスタイプ（SV, CL, プリンタ, その他）
 * @return {Object} ヘッダー情報（不足ヘッダー追加後）
 */
function detectAndEnsureHeaders(sheet, deviceType = null) {
  try {
    // 最初にヘッダー行を検出
    let headerInfo = detectHeaderRow(sheet);
    
    if (!headerInfo.found) {
      // ヘッダー行が見つからない場合、3行目にすべて作成
      addFormLog('ヘッダー行未検出、新規作成を開始', {
        sheetName: sheet.getName(),
        deviceType: deviceType
      });
      
      const allRequiredFields = getRequiredFieldsForDeviceType(deviceType);
      const headerAddResult = addMissingHeadersToSheet(sheet, allRequiredFields, deviceType);
      
      if (headerAddResult.success) {
        // ヘッダー追加後に再検出
        headerInfo = detectHeaderRow(sheet);
        if (headerInfo.found) {
          addFormLog('新規ヘッダー作成後の検出成功', {
            rowIndex: headerInfo.rowIndex,
            detectedColumns: Object.keys(headerInfo.columnIndexes)
          });
          
          return {
            ...headerInfo,
            headersAdded: true,
            addedHeadersInfo: headerAddResult
          };
        }
      }
      
      return {
        found: false,
        error: 'ヘッダー作成後も検出に失敗しました',
        headerAddResult: headerAddResult
      };
    }
    
    // ヘッダー行は見つかったが、不足している列がある場合
    const missingFields = detectMissingHeaders(headerInfo.columnIndexes, deviceType);
    
    if (missingFields.length > 0) {
      addFormLog('不足ヘッダーを自動追加開始', {
        sheetName: sheet.getName(),
        deviceType: deviceType,
        missingFields: missingFields,
        currentColumns: Object.keys(headerInfo.columnIndexes)
      });
      
      const headerAddResult = addMissingHeadersToSheet(sheet, missingFields, deviceType);
      
      if (headerAddResult.success) {
        // ヘッダー追加後に再検出
        const updatedHeaderInfo = detectHeaderRow(sheet);
        if (updatedHeaderInfo.found) {
          addFormLog('不足ヘッダー追加後の検出成功', {
            rowIndex: updatedHeaderInfo.rowIndex,
            detectedColumns: Object.keys(updatedHeaderInfo.columnIndexes),
            previousColumns: Object.keys(headerInfo.columnIndexes).length,
            newColumns: Object.keys(updatedHeaderInfo.columnIndexes).length
          });
          
          return {
            ...updatedHeaderInfo,
            headersAdded: true,
            addedHeadersInfo: headerAddResult
          };
        }
      }
      
      // ヘッダー追加に失敗した場合は元の情報を返す
      addFormLog('不足ヘッダー追加失敗、元の情報を使用', {
        error: headerAddResult.error,
        originalColumns: Object.keys(headerInfo.columnIndexes)
      });
    }
    
    return {
      ...headerInfo,
      headersAdded: missingFields.length > 0,
      addedHeadersInfo: missingFields.length > 0 ? { success: false, added: 0 } : null
    };
    
  } catch (error) {
    addFormLog('ヘッダー検出・追加処理エラー', {
      error: error.toString(),
      sheetName: sheet.getName()
    });
    
    return {
      found: false,
      error: error.message
    };
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

    // 重複チェックを最初に実行（フォーム作成前）
    let duplicateCheckResult = null;
    try {
      const checkData = {
        locationNumber: formConfig.locationNumber,
        assetNumber: formConfig.attributes?.assetNumber || '',
        serial: formConfig.attributes?.serial || ''
      };
      
      duplicateCheckResult = checkDuplicateValues(formConfig.deviceCategory, checkData);
      
      if (!duplicateCheckResult.success) {
        addFormLog('重複チェック失敗', {
          error: duplicateCheckResult.error,
          deviceCategory: formConfig.deviceCategory
        });
        
        return {
          success: false,
          error: `重複チェックに失敗しました: ${duplicateCheckResult.error}`,
          errorType: 'DUPLICATE_CHECK_FAILED'
        };
      } else if (duplicateCheckResult.hasDuplicates) {
        addFormLog('重複データ検出（フォーム作成前）', {
          duplicates: duplicateCheckResult.duplicates,
          deviceCategory: formConfig.deviceCategory
        });
        
        // 重複詳細メッセージを生成（シンプル版）
        const duplicateMessages = duplicateCheckResult.duplicates.map((duplicate, index) => {
          return `${duplicate.fieldDisplayName}「${duplicate.value}」は既に使用されています`;
        });
        
        const detailedErrorMessage = `入力エラー: 以下の値が重複しています\n\n${duplicateMessages.join('\n')}\n\n異なる値を入力してください。`;
        
        // 重複がある場合はフォーム作成を開始せずにエラーを返す
        return {
          success: false,
          error: detailedErrorMessage,
          errorType: 'DUPLICATE_DATA',
          duplicateCheck: duplicateCheckResult,
          duplicateDetails: {
            count: duplicateCheckResult.duplicates.length,
            fields: duplicateCheckResult.duplicates.map(d => d.field),
            messages: duplicateMessages,
            sheetName: duplicateCheckResult.sheetName,
            items: duplicateCheckResult.duplicates.map(duplicate => ({
              field: duplicate.field,
              fieldDisplayName: duplicate.fieldDisplayName,
              value: duplicate.value,
              location: {
                sheetName: duplicate.sheetName,
                rowNumber: duplicate.rowNumber,
                columnLetter: duplicate.columnLetter
              },
              additionalInfo: duplicate.additionalInfo || {}
            }))
          }
        };
      } else {
        addFormLog('重複チェック完了（重複なし）- フォーム作成を開始', {
          checkedFields: Object.keys(checkData),
          deviceCategory: formConfig.deviceCategory
        });
      }
      
    } catch (duplicateCheckError) {
      addFormLog('重複チェック処理でエラー（フォーム作成前）', {
        error: duplicateCheckError.toString(),
        deviceCategory: formConfig.deviceCategory
      });
      
      return {
        success: false,
        error: `重複チェック処理でエラーが発生しました: ${duplicateCheckError.message}`,
        errorType: 'DUPLICATE_CHECK_ERROR'
      };
    }
    
    // 拠点別フォルダを取得
    const folderResult = getFormStorageFolderId(formConfig.location, formConfig.deviceCategory);
    if (!folderResult.success) {
      return {
        success: false,
        error: folderResult.error,
        errorType: 'FOLDER_NOT_CONFIGURED'
      };
    }
    
    let folder;
    try {
      folder = DriveApp.getFolderById(folderResult.folderId);
      addFormLog('拠点別フォルダを使用', { 
        location: formConfig.location, 
        folderId: folderResult.folderId 
      });
    } catch (folderError) {
      return {
        success: false,
        error: `フォルダへのアクセスに失敗しました: ${folderError.message}`,
        errorType: 'FOLDER_ACCESS_ERROR'
      };
    }
    
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
    


    // スプレッドシートに行を追加
    let spreadsheetResult = null;
    try {
      const additionalData = {
        formId: form.getId(),
        createdDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss'),
        formUrl: form.getEditUrl(),
        publicUrl: form.getPublishedUrl()
      };
      
      // フロントエンドからのformConfigを、スプレッドシート用のformDataに変換
      const spreadsheetFormData = {
        locationNumber: formConfig.locationNumber,
        assetNumber: formConfig.attributes?.assetNumber || '',
        modelNumber: formConfig.attributes?.model || '',
        serial: formConfig.attributes?.serial || '',
        software: formConfig.attributes?.software || '',
        os: formConfig.attributes?.os || '',
        deviceType: formConfig.deviceCategory || ''  // deviceCategoryを使用（SV/CL等）
      };
      
      addFormLog('スプレッドシート用データ変換', {
        originalFormConfig: {
          locationNumber: formConfig.locationNumber,
          deviceType: formConfig.deviceType,
          deviceCategory: formConfig.deviceCategory,
          attributes: formConfig.attributes
        },
        convertedSpreadsheetData: spreadsheetFormData
      });
      
      spreadsheetResult = addRowToSpreadsheet(formConfig.deviceCategory, spreadsheetFormData, additionalData);
      
      if (spreadsheetResult.success) {
        addFormLog('スプレッドシート連携成功', {
          sheetName: spreadsheetResult.sheetName,
          spreadsheetId: spreadsheetResult.spreadsheetId
        });
      } else {
        addFormLog('スプレッドシート連携失敗', {
          error: spreadsheetResult.error,
          deviceCategory: formConfig.deviceCategory
        });
      }
      
    } catch (spreadsheetError) {
      addFormLog('スプレッドシート連携処理でエラー', {
        error: spreadsheetError.toString(),
        deviceCategory: formConfig.deviceCategory,
        formId: form.getId()
      });
      // スプレッドシート連携エラーはフォーム作成自体は継続
      spreadsheetResult = {
        success: false,
        error: spreadsheetError.toString()
      };
    }

    // ステータス収集シートに行を追加
    let statusSheetResult = null;
    try {
      statusSheetResult = addRowToStatusCollectionSheet(formConfig.locationNumber, formConfig.deviceCategory);
      
      if (statusSheetResult.success) {
        addFormLog('ステータス収集シート追記成功', {
          sheetName: statusSheetResult.sheetName,
          locationNumber: formConfig.locationNumber,
          status: '999.フォーム作成完了'
        });
      } else {
        addFormLog('ステータス収集シート追記失敗', {
          error: statusSheetResult.error,
          locationNumber: formConfig.locationNumber
        });
      }
      
    } catch (statusSheetError) {
      addFormLog('ステータス収集シート追記処理でエラー', {
        error: statusSheetError.toString(),
        locationNumber: formConfig.locationNumber,
        deviceCategory: formConfig.deviceCategory
      });
      // ステータス収集シート追記エラーはフォーム作成自体は継続
      statusSheetResult = {
        success: false,
        error: statusSheetError.toString()
      };
    }
    
    addFormLog('フォーム作成成功', {
      formId: form.getId(),
      formUrl: form.getEditUrl(),
      publicUrl: form.getPublishedUrl(),
      spreadsheetLinked: spreadsheetResult?.success || false
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
          publicImageUrl: qrCodeResult.publicImageUrl, // 画像表示用の公開URL
          thumbnailUrl: qrCodeResult.thumbnailUrl, // サムネイル用URL
          base64ImageData: qrCodeResult.base64ImageData, // Base64エンコード画像データ
          folderId: qrCodeResult.folderId,
          qrCodeUrl: form.getPublishedUrl() // QRコードが指すフォームURL
        } : { success: false, error: 'QRコード生成に失敗しました' },
        spreadsheet: spreadsheetResult ? {
          success: spreadsheetResult.success,
          sheetName: spreadsheetResult.sheetName,
          spreadsheetId: spreadsheetResult.spreadsheetId,
          rowData: spreadsheetResult.rowData,
          error: spreadsheetResult.error
        } : { success: false, error: 'スプレッドシート連携が実行されませんでした' },
        statusSheet: statusSheetResult ? {
          success: statusSheetResult.success,
          sheetName: statusSheetResult.sheetName,
          spreadsheetId: statusSheetResult.spreadsheetId,
          locationNumber: formConfig.locationNumber,
          status: '999.フォーム作成完了',
          error: statusSheetResult.error
        } : { success: false, error: 'ステータス収集シート追記が実行されませんでした' },
        duplicateCheck: duplicateCheckResult ? {
          success: duplicateCheckResult.success,
          hasDuplicates: duplicateCheckResult.hasDuplicates,
          duplicates: duplicateCheckResult.duplicates,
          checkedFields: duplicateCheckResult.checkedFields,
          error: duplicateCheckResult.error
        } : { success: false, error: '重複チェックが実行されませんでした' }
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
 * 拠点管理番号のバリデーション（製造番号を含む形式に対応）
 * @param {string} locationNumber - 拠点管理番号
 */
function validateLocationNumber(locationNumber) {
  if (!locationNumber || typeof locationNumber !== 'string') {
    return { valid: false, error: '拠点管理番号が無効です' };
  }
  
  // アンダースコアで分割
  const parts = locationNumber.split('_');
  if (parts.length < 4) {
    return { 
      valid: false, 
      error: '拠点管理番号は「拠点_カテゴリ_モデル_製造番号_連番」の形式で入力してください（例：OSA_SV_Server_ABC12345_001）' 
    };
  }
  
  // 各部分を検証
  const [location, category, model, serial, number] = parts;
  
  // 拠点コード（1番目）は英字のみ許可
  const locationPattern = /^[A-Za-z]+$/;
  if (!locationPattern.test(location)) {
    return { 
      valid: false, 
      error: '拠点コードは英字のみで入力してください（例：OSA、KOB）' 
    };
  }
  
  // 拠点コードの長さチェック（2文字以上）
  if (location.length < 2) {
    return { 
      valid: false, 
      error: '拠点コードは2文字以上で入力してください' 
    };
  }
  
  // カテゴリ（2番目）は英数字のみ許可
  const categoryPattern = /^[A-Za-z0-9]+$/;
  if (!categoryPattern.test(category)) {
    return { 
      valid: false, 
      error: 'カテゴリは英数字のみで入力してください（例：SV、CL、Printer）' 
    };
  }
  
  // モデル（3番目）は英数字とアンダースコア許可
  const modelPattern = /^[A-Za-z0-9_]+$/;
  if (!modelPattern.test(model)) {
    return { 
      valid: false, 
      error: 'モデルは英数字とアンダースコアのみで入力してください（例：Server、Desktop）' 
    };
  }
  
  // 製造番号（4番目）は英数字のみ許可
  const serialPattern = /^[A-Za-z0-9]+$/;
  if (!serialPattern.test(serial)) {
    return { 
      valid: false, 
      error: '製造番号は英数字のみで入力してください（例：ABC12345）' 
    };
  }
  
  // 連番（5番目）は数字のみ許可（オプション）
  if (number && !/^[0-9]+$/.test(number)) {
    return { 
      valid: false, 
      error: '連番は数字のみで入力してください（例：001）' 
    };
  }
  
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
    
    // ファイルを誰でも閲覧可能に設定（画像表示用）
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingError) {
      addFormLog('QRコードファイル共有設定エラー', {
        error: sharingError.toString(),
        fileId: file.getId()
      });
    }
    
    // 複数の公開画像URL形式を用意
    const fileId = file.getId();
    const publicImageUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
    const thumbnailUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w400`;
    
    // Base64エンコードした画像データも生成（CORS回避用）
    let base64ImageData = null;
    try {
      const imageBytes = qrCodeBlob.getBytes();
      const base64String = Utilities.base64Encode(imageBytes);
      base64ImageData = `data:image/png;base64,${base64String}`;
    } catch (base64Error) {
      addFormLog('Base64変換エラー', {
        error: base64Error.toString(),
        fileId: fileId
      });
    }
    
    addFormLog('QRコード保存成功', {
      fileName,
      fileId: fileId,
      locationNumber,
      folderId: folder.getId(),
      publicImageUrl,
      thumbnailUrl,
      hasBase64Data: !!base64ImageData
    });
    
    return {
      success: true,
      fileId: fileId,
      fileName: fileName,
      fileUrl: file.getUrl(),
      publicImageUrl: publicImageUrl, // 画像表示用の公開URL
      thumbnailUrl: thumbnailUrl, // サムネイル用URL
      base64ImageData: base64ImageData, // Base64エンコード画像データ
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
      deviceCategory: 'CL',  // SV/CL形式で設定
      location: 'osaka-desktop',
      // テスト用attributes（スプレッドシート連携テスト）
      attributes: {
        assetNumber: 'QR-ASSET-001',
        model: 'QR-MODEL-123',
        serial: 'QR-SER789012',
        software: 'QRテストソフト',
        os: 'Windows 11'
      }
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
      deviceCategory: testDeviceType === 'printer' ? 'プリンタ' : 'SV',  // プリンタの場合は「プリンタ」
      location: testDeviceType === 'printer' ? 'osaka-printer' : 'osaka-desktop',
      // テスト用attributes（スプレッドシート連携テスト）
      attributes: {
        assetNumber: testDeviceType === 'terminal' ? 'MSG-ASSET-001' : '',  // プリンタには資産番号なし
        model: 'MSG-MODEL-123',
        serial: 'MSG-SER345678',
        software: testDeviceType === 'terminal' ? 'メッセージテストソフト' : '',  // プリンタにはソフトなし
        os: testDeviceType === 'terminal' ? 'Windows 11' : ''  // プリンタにはOSなし
      }
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

/**
 * スプレッドシートに行を追加する関数
 * @param {string} location - 拠点識別子
 * @param {Object} formData - フォームデータ
 * @param {Object} additionalData - 追加データ
 */
function addRowToSpreadsheet(deviceCategory, formData, additionalData = {}) {
  addFormLog('スプレッドシート行追加開始', { deviceCategory, formData, additionalData });
  
  try {
    // スプレッドシートIDを取得
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty(SPREADSHEET_ID_DESTINATION_KEY);
    
    if (!spreadsheetId) {
      throw new Error('SPREADSHEET_ID_DESTINATIONが設定されていません');
    }
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // デバイスカテゴリ別シート名を取得
    const sheetName = getTargetSheetNameByCategory(deviceCategory);
    if (!sheetName) {
      throw new Error(`未知のデバイスカテゴリ: ${deviceCategory}`);
    }
    
    // シートを取得（存在しない場合はエラー）
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません。事前にシートを作成してください。`);
    }
    
    // ヘッダー行を動的に検出し、不足している場合は自動追加
    const headerInfo = detectAndEnsureHeaders(sheet, formData.deviceType);
    
    if (!headerInfo.found) {
      throw new Error(`シート「${sheetName}」でヘッダー行を検出できませんでした: ${headerInfo.error}`);
    }
    
    addFormLog('ヘッダー行検出完了', {
      sheetName,
      headerRowIndex: headerInfo.rowIndex,
      detectedColumns: Object.keys(headerInfo.columnIndexes),
      totalColumns: headerInfo.rowData.length,
      headersAdded: headerInfo.headersAdded || false,
      addedHeadersCount: headerInfo.addedHeadersInfo?.added || 0
    });
    
    // 次の行番号を取得（数式で使用）
    const nextRowNumber = sheet.getLastRow() + 1;
    
    addFormLog('行データ作成開始', {
      sheetName,
      nextRowNumber,
      formDataDeviceType: formData.deviceType,
      headerColumnsCount: Object.keys(headerInfo.columnIndexes).length
    });
    
    // 行データを作成（数式付き、動的列インデックス使用）
    const rowData = createRowDataWithFormulas(formData, additionalData, headerInfo.columnIndexes, headerInfo.rowData.length, nextRowNumber);
    
    addFormLog('行データ作成完了', {
      rowDataLength: rowData.length,
      nonEmptyValues: rowData.filter(val => val !== '').length,
      formulaValues: rowData.filter(val => typeof val === 'string' && val.startsWith('=')).length,
      sampleData: rowData.slice(0, 5)  // 最初の5列のサンプル
    });
    
    // 数式を正しく保存するため、setValuesを使用
    const targetRange = sheet.getRange(nextRowNumber, 1, 1, rowData.length);
    targetRange.setValues([rowData]);
    
    addFormLog('数式付きデータをsetValuesで追加', {
      targetRow: nextRowNumber,
      targetColumns: rowData.length,
      range: targetRange.getA1Notation(),
      containsFormulas: rowData.some(value => typeof value === 'string' && value.startsWith('='))
    });
    
         addFormLog('スプレッドシート行追加完了', {
       sheetName,
       rowDataLength: rowData.length,
       lastRowAfter: sheet.getLastRow()
     });
     
     return {
       success: true,
       sheetName: sheetName,
       spreadsheetId: spreadsheetId,
       rowData: rowData,
       headerRowIndex: headerInfo.rowIndex,
       detectedColumns: Object.keys(headerInfo.columnIndexes)
     };
    
  } catch (error) {
    addFormLog('スプレッドシート行追加エラー', {
      deviceCategory,
      error: error.toString(),
      stack: error.stack
    });
    
    // エラーが発生してもフォーム作成は継続
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * デバイスカテゴリからステータス収集シート名を取得
 * @param {string} deviceCategory - デバイスカテゴリ（SV, CL, プリンタ, その他）
 * @return {string} ステータス収集シート名
 */
function getStatusCollectionSheetName(deviceCategory) {
  if (!deviceCategory) {
    return null;
  }
  
  // SV, CLは端末ステータス収集シート
  if (deviceCategory === 'SV' || deviceCategory === 'CL') {
    return STATUS_COLLECTION_SHEETS.terminal;
  }
  
  // プリンタ, その他はプリンタステータス収集シート
  if (deviceCategory === 'プリンタ' || deviceCategory === 'その他') {
    return STATUS_COLLECTION_SHEETS.printer;
  }
  
  return null;
}

/**
 * ステータス収集シートの列インデックスを動的に検出する関数
 * @param {string} statusSheetName - ステータス収集シート名
 * @return {Object} 列名とインデックスのマッピング
 */
function detectStatusSheetColumns(statusSheetName) {
  try {
    const spreadsheetSettings = getSpreadsheetSettings();
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      return { success: false, error: 'スプレッドシート設定が見つかりません' };
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetSettings.settings.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(statusSheetName);
    
    if (!sheet) {
      return { success: false, error: `ステータス収集シート '${statusSheetName}' が見つかりません` };
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow === 0 || lastColumn === 0) {
      return { success: false, error: 'ステータス収集シートが空です（行数：' + lastRow + '、列数：' + lastColumn + '）' };
    }
    
    // 1行目と3行目をチェック（メインシートと同様）
    const candidateRows = [1, 3].filter(rowNum => rowNum <= lastRow);
    
    addFormLog('ステータス収集シート構造確認', {
      sheetName: statusSheetName,
      lastRow,
      lastColumn,
      candidateRows
    });
    let bestResult = null;
    let bestScore = -1;
    
    for (const rowNum of candidateRows) {
      const headerRow = sheet.getRange(rowNum, 1, 1, lastColumn).getValues()[0];
      const columnIndexes = {};
      let score = 0;
      
      addFormLog('ステータス収集シートヘッダー行確認', {
        sheetName: statusSheetName,
        rowNum,
        headerRow: headerRow.slice(0, 10), // 最初の10列だけ表示
        headerRowLength: headerRow.length
      });
      
      // 各フィールドの列インデックスを検索
      Object.keys(STATUS_SHEET_COLUMN_MAPPING).forEach(fieldName => {
        const candidateNames = STATUS_SHEET_COLUMN_MAPPING[fieldName];
        
        for (let i = 0; i < headerRow.length; i++) {
          const headerValue = headerRow[i] ? headerRow[i].toString().trim() : '';
          
          for (const candidateName of candidateNames) {
            if (headerValue === candidateName) {
              columnIndexes[fieldName] = getColumnLetter(i + 1);
              score++;
              addFormLog('ステータス収集シート列マッチ', {
                fieldName,
                candidateName,
                headerValue,
                columnLetter: getColumnLetter(i + 1),
                columnIndex: i
              });
              break;
            }
          }
          
          if (columnIndexes[fieldName]) break;
        }
      });
      
      addFormLog('ステータス収集シート行検出結果', {
        rowNum,
        score,
        detectedColumns: Object.keys(columnIndexes),
        columnIndexes
      });
      
      if (score > bestScore) {
        bestScore = score;
        bestResult = {
          headerRowIndex: rowNum,
          columnIndexes: columnIndexes,
          score: score
        };
      }
    }
    
    if (bestResult && bestScore > 0) {
      addFormLog('ステータス収集シート列検出成功', {
        sheetName: statusSheetName,
        headerRowIndex: bestResult.headerRowIndex,
        detectedColumns: Object.keys(bestResult.columnIndexes),
        score: bestScore
      });
      
      return {
        success: true,
        columnIndexes: bestResult.columnIndexes,
        headerRowIndex: bestResult.headerRowIndex
      };
    } else {
      return {
        success: false,
        error: `ステータス収集シート '${statusSheetName}' で有効な列が見つかりませんでした`
      };
    }
    
  } catch (error) {
    addFormLog('ステータス収集シート列検出エラー', {
      sheetName: statusSheetName,
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 動的列検出を使用したVLOOKUP数式を生成する関数
 * @param {string} statusSheetName - ステータス収集シート名
 * @param {string} lookupValue - 検索値（拠点管理番号のセル参照）
 * @param {string} fieldName - 取得するフィールド名
 * @param {Object} statusColumns - ステータス収集シートの列マッピング
 * @param {string} condition - 条件（オプション）
 * @return {string} 数式
 */
function generateDynamicVlookupFormula(statusSheetName, lookupValue, fieldName, statusColumns, condition = null) {
  if (!statusSheetName || !lookupValue || !fieldName || !statusColumns) {
    return '';
  }
  
  const returnColumn = statusColumns[fieldName];
  if (!returnColumn) {
    addFormLog('数式生成スキップ', {
      reason: `フィールド '${fieldName}' の列が見つかりません`,
      availableColumns: Object.keys(statusColumns)
    });
    return '';
  }
  
  // 拠点管理番号列を動的に取得（デフォルトはB列）
  const lookupColumn = statusColumns['locationNumber'] || 'B';
  
  // 一致する拠点管理番号の最後の行を取得する数式（配列数式使用）
  // FILTER関数で最後の一致を取得（Google スプレッドシート用、無限範囲対応）
  let formula = `=IFERROR(INDEX(FILTER('${statusSheetName}'!${returnColumn}:${returnColumn}, '${statusSheetName}'!${lookupColumn}:${lookupColumn}=${lookupValue}), ROWS(FILTER('${statusSheetName}'!${returnColumn}:${returnColumn}, '${statusSheetName}'!${lookupColumn}:${lookupColumn}=${lookupValue}))), "")`;
  
  // 条件がある場合の処理
  if (condition) {
    // 貸出ステータスが"1.貸出中"の場合の条件式
    if (condition === 'lendingOnly') {
      const statusColumn = statusColumns['status'];
      if (statusColumn) {
        // 最後の行のステータスをチェックしてから値を取得
        const statusCheckFormula = `INDEX(FILTER('${statusSheetName}'!${statusColumn}:${statusColumn}, '${statusSheetName}'!${lookupColumn}:${lookupColumn}=${lookupValue}), ROWS(FILTER('${statusSheetName}'!${statusColumn}:${statusColumn}, '${statusSheetName}'!${lookupColumn}:${lookupColumn}=${lookupValue})))`;
        const valueFormula = `INDEX(FILTER('${statusSheetName}'!${returnColumn}:${returnColumn}, '${statusSheetName}'!${lookupColumn}:${lookupColumn}=${lookupValue}), ROWS(FILTER('${statusSheetName}'!${returnColumn}:${returnColumn}, '${statusSheetName}'!${lookupColumn}:${lookupColumn}=${lookupValue})))`;
        formula = `=IF(${statusCheckFormula}="1.貸出中", IFERROR(${valueFormula}, ""), "")`;
      } else {
        addFormLog('条件付き数式生成警告', {
          reason: 'ステータス列が見つからないため、条件なしの数式を使用します',
          fieldName: fieldName,
          availableColumns: Object.keys(statusColumns)
        });
        // ステータス列が見つからない場合は条件なしの数式を使用
      }
    }
  }
  
  addFormLog('動的数式生成完了（最後の行参照）', {
    statusSheetName,
    fieldName,
    returnColumn,
    lookupColumn,
    condition,
    formula,
    note: '一致する拠点管理番号の最後の行を参照'
  });
  
  return formula;
}



/**
 * 重複値をチェックする関数
 * @param {string} deviceCategory - デバイスカテゴリ（SV, CL, プリンタ, その他）
 * @param {Object} checkData - チェック対象データ
 * @param {string} checkData.locationNumber - 拠点管理番号
 * @param {string} checkData.assetNumber - 資産管理番号
 * @param {string} checkData.serial - シリアル番号
 * @return {Object} チェック結果
 */
function checkDuplicateValues(deviceCategory, checkData) {
  try {
    // スプレッドシート設定を取得
    const spreadsheetSettings = getSpreadsheetSettings();
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      return {
        success: false,
        error: 'スプレッドシート設定が見つかりません'
      };
    }

    // デバイスカテゴリに対応するシート名を取得
    const sheetName = getTargetSheetNameByCategory(deviceCategory);
    if (!sheetName) {
      return {
        success: false,
        error: `デバイスカテゴリ '${deviceCategory}' に対応するシートが見つかりません`
      };
    }

    // スプレッドシートとシートを取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetSettings.settings.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      return {
        success: false,
        error: `シート '${sheetName}' が見つかりません`
      };
    }

    // ヘッダー行を検出
    const headerInfo = detectHeaderRow(sheet);
    if (!headerInfo.found) {
      return {
        success: false,
        error: `シート '${sheetName}' でヘッダー行が見つかりません: ${headerInfo.error}`
      };
    }

    // 重複チェック対象の列を特定
    const columnIndexes = headerInfo.columnIndexes;
    const duplicates = [];
    const checkedFields = [];

    // データ範囲を取得（ヘッダー行の次の行から最終行まで）
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow <= headerInfo.rowIndex) {
      // データがない場合は重複なし
      return {
        success: true,
        hasDuplicates: false,
        duplicates: [],
        checkedFields: Object.keys(checkData).filter(key => checkData[key])
      };
    }

    const dataRange = sheet.getRange(headerInfo.rowIndex + 1, 1, lastRow - headerInfo.rowIndex, lastColumn);
    const dataValues = dataRange.getValues();

    // デバイスタイプに応じて除外するフィールドを決定
    const excludedFields = (deviceType === 'プリンタ' || deviceType === 'その他') 
      ? ['assetNumber', 'software', 'os', 'depositReceiptNo'] 
      : [];

    // 各フィールドの重複をチェック
    Object.keys(checkData).forEach(fieldName => {
      const value = checkData[fieldName];
      if (!value) return; // 空の値はスキップ

      // プリンタ・その他の場合、除外対象フィールドはスキップ
      if (excludedFields.includes(fieldName)) {
        addFormLog('重複チェック: デバイスタイプにより除外', {
          fieldName,
          deviceType,
          value,
          excludedFields
        });
        return;
      }

      const columnIndex = columnIndexes[fieldName];
      if (columnIndex === undefined) {
        addFormLog('重複チェック: 列が見つかりません', {
          fieldName,
          sheetName,
          availableColumns: Object.keys(columnIndexes)
        });
        return;
      }

      checkedFields.push(fieldName);

             // データ行で重複をチェック
       for (let i = 0; i < dataValues.length; i++) {
         const cellValue = dataValues[i][columnIndex];
         if (cellValue && cellValue.toString().trim() === value.toString().trim()) {
           // 重複行の他の情報も取得（参考情報として）
           const rowData = dataValues[i];
           const additionalInfo = {};
           
           // 拠点管理番号、資産番号、型番などの参考情報を取得
           ['locationNumber', 'assetNumber', 'modelNumber'].forEach(refField => {
             const refColumnIndex = columnIndexes[refField];
             if (refColumnIndex !== undefined && rowData[refColumnIndex]) {
               additionalInfo[refField] = rowData[refColumnIndex].toString().trim();
             }
           });
           
           duplicates.push({
             field: fieldName,
             fieldDisplayName: getFieldDisplayName(fieldName),
             value: value,
             duplicateValue: cellValue,
             rowNumber: headerInfo.rowIndex + 1 + i + 1, // 実際の行番号
             columnLetter: getColumnLetter(columnIndex + 1),
             sheetName: sheetName,
             additionalInfo: additionalInfo
           });
         }
       }
    });

    addFormLog('重複チェック完了', {
      sheetName,
      deviceType: deviceType,
      excludedFields: excludedFields,
      checkedFields,
      duplicatesFound: duplicates.length,
      duplicates: duplicates.map(d => ({
        field: d.field,
        value: d.value,
        rowNumber: d.rowNumber
      }))
    });

    return {
      success: true,
      hasDuplicates: duplicates.length > 0,
      duplicates: duplicates,
      checkedFields: checkedFields,
      sheetName: sheetName,
      totalDataRows: dataValues.length
    };

  } catch (error) {
    addFormLog('重複チェックエラー', {
      location,
      checkData,
      error: error.toString(),
      stack: error.stack
    });

    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ステータス収集シートに行を追加する関数
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} deviceCategory - デバイスカテゴリ（SV, CL, プリンタ, その他）
 * @return {Object} 追加結果
 */
function addRowToStatusCollectionSheet(locationNumber, deviceCategory) {
  try {
    // スプレッドシート設定を取得
    const spreadsheetSettings = getSpreadsheetSettings();
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      return {
        success: false,
        error: 'スプレッドシート設定が見つかりません'
      };
    }

    // ステータス収集シート名を取得
    const statusSheetName = getStatusCollectionSheetName(deviceCategory);
    if (!statusSheetName) {
      return {
        success: false,
        error: `デバイスカテゴリ '${deviceCategory}' に対応するステータス収集シートが見つかりません`
      };
    }

    // スプレッドシートとシートを取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetSettings.settings.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(statusSheetName);
    
    if (!sheet) {
      return {
        success: false,
        error: `ステータス収集シート '${statusSheetName}' が見つかりません`
      };
    }

    // ステータス収集シートの列を動的に検出
    const statusSheetColumns = detectStatusSheetColumns(statusSheetName);
    
    if (!statusSheetColumns.success) {
      return {
        success: false,
        error: `ステータス収集シート '${statusSheetName}' の列検出に失敗しました: ${statusSheetColumns.error}`
      };
    }

    // 必要な列が存在するかチェック
    const locationColumn = statusSheetColumns.columnIndexes['locationNumber'];
    const statusColumn = statusSheetColumns.columnIndexes['status'];
    
    if (!locationColumn || !statusColumn) {
      return {
        success: false,
        error: `必要な列が見つかりません (拠点管理番号: ${locationColumn}, ステータス: ${statusColumn})`
      };
    }

    // 新しい行を追加
    const lastRow = sheet.getLastRow();
    const newRowNumber = lastRow + 1;
    
    // タイムスタンプを生成
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    
    // 行データを作成（検出された列に対応）
    const totalColumns = sheet.getLastColumn() || Object.keys(statusSheetColumns.columnIndexes).length;
    const rowData = new Array(totalColumns).fill('');
    
    // 拠点管理番号列のインデックスを取得（列文字から数値に変換）
    const locationColumnIndex = getColumnIndex(locationColumn);
    const statusColumnIndex = getColumnIndex(statusColumn);
    
    // データを設定
    rowData[locationColumnIndex - 1] = locationNumber;  // 0ベースに調整
    rowData[statusColumnIndex - 1] = '999.フォーム作成完了';  // 0ベースに調整
    
    // タイムスタンプ列があれば設定
    const timestampColumn = statusSheetColumns.columnIndexes['timestamp'];
    if (timestampColumn) {
      const timestampColumnIndex = getColumnIndex(timestampColumn);
      rowData[timestampColumnIndex - 1] = timestamp;  // 0ベースに調整
    }

    // 行を追加
    const targetRange = sheet.getRange(newRowNumber, 1, 1, rowData.length);
    targetRange.setValues([rowData]);

    addFormLog('ステータス収集シート行追加完了', {
      sheetName: statusSheetName,
      newRowNumber,
      locationNumber,
      status: '999.フォーム作成完了',
      timestamp,
      detectedColumns: Object.keys(statusSheetColumns.columnIndexes)
    });

    return {
      success: true,
      sheetName: statusSheetName,
      spreadsheetId: spreadsheetSettings.settings.spreadsheetId,
      rowNumber: newRowNumber,
      locationNumber: locationNumber,
      status: '999.フォーム作成完了',
      timestamp: timestamp
    };

  } catch (error) {
    addFormLog('ステータス収集シート行追加エラー', {
      locationNumber,
      deviceCategory,
      error: error.toString(),
      stack: error.stack
    });

    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * フィールド名を日本語表示名に変換する関数
 * @param {string} fieldName - フィールド名
 * @return {string} 日本語表示名
 */
function getFieldDisplayName(fieldName) {
  const displayNames = {
    'locationNumber': '拠点管理番号',
    'assetNumber': '資産管理番号',
    'serial': 'シリアル番号(製造番号)',
    'modelNumber': '型番',
    'software': 'ソフトウェア',
    'os': 'OS',
    'deviceType': 'デバイス種別'
  };
  
  return displayNames[fieldName] || fieldName;
}

/**
 * 列文字（A, B, C等）を列番号（1, 2, 3等）に変換する関数
 * @param {string} columnLetter - 列文字
 * @return {number} 列番号（1ベース）
 */
function getColumnIndex(columnLetter) {
  let result = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    result = result * 26 + (columnLetter.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * 列番号をアルファベットに変換する関数
 * @param {number} columnNumber - 列番号（1ベース）
 * @return {string} 列のアルファベット表記
 */
function getColumnLetter(columnNumber) {
  let result = '';
  while (columnNumber > 0) {
    columnNumber--;
    result = String.fromCharCode(65 + (columnNumber % 26)) + result;
    columnNumber = Math.floor(columnNumber / 26);
  }
  return result;
}

/**
 * 数式付きの行データを作成する関数（動的列検出対応）
 * @param {Object} formData - フォームデータ
 * @param {Object} additionalData - 追加データ
 * @param {Object} columnIndexes - 列インデックスマッピング
 * @param {number} totalColumns - 総列数
 * @param {number} currentRowNumber - 現在の行番号
 * @return {Array} 行データ（数式含む）
 */
function createRowDataWithFormulas(formData, additionalData, columnIndexes, totalColumns, currentRowNumber) {
  try {
    // 基本の行データを作成
    const rowData = createRowData(formData, additionalData, columnIndexes, totalColumns);
    
    addFormLog('基本行データ作成成功', {
      rowDataLength: rowData.length,
      totalColumns,
      deviceCategory: formData.deviceType
    });
    
    // 数式が必要な列の処理
    const deviceCategory = formData.deviceType;  // deviceTypeフィールドから取得
    const statusSheetName = getStatusCollectionSheetName(deviceCategory);
  
  if (statusSheetName && columnIndexes.locationNumber !== undefined) {
    const locationNumberCell = getColumnLetter(columnIndexes.locationNumber + 1) + currentRowNumber;
    
    // ステータス収集シートの列を動的に検出
    const statusSheetColumns = detectStatusSheetColumns(statusSheetName);
    
    if (statusSheetColumns.success) {
      addFormLog('ステータス収集シート列検出成功', {
        statusSheetName,
        detectedColumns: Object.keys(statusSheetColumns.columnIndexes),
        columnCount: Object.keys(statusSheetColumns.columnIndexes).length
      });
      
      // 各列の数式を設定（動的列検出使用）
      const formulaMapping = {
        assignee: generateDynamicVlookupFormula(statusSheetName, locationNumberCell, 'assignee', statusSheetColumns.columnIndexes),
        lendingStatus: generateDynamicVlookupFormula(statusSheetName, locationNumberCell, 'status', statusSheetColumns.columnIndexes),
        lendingDestination: generateDynamicVlookupFormula(statusSheetName, locationNumberCell, 'destination', statusSheetColumns.columnIndexes),
        lendingDate: generateDynamicVlookupFormula(statusSheetName, locationNumberCell, 'timestamp', statusSheetColumns.columnIndexes, 'lendingOnly'),
        userDeviceDeposit: generateDynamicVlookupFormula(statusSheetName, locationNumberCell, 'userDeviceDeposit', statusSheetColumns.columnIndexes, 'lendingOnly'),
        remarks: generateDynamicVlookupFormula(statusSheetName, locationNumberCell, 'remarks', statusSheetColumns.columnIndexes, 'lendingOnly')
      };
      
      // お預かり証No.は端末（SV/CL）かつ貸出中の場合のみ
      if (deviceCategory === 'SV' || deviceCategory === 'CL') {
        formulaMapping.depositReceiptNo = generateDynamicVlookupFormula(statusSheetName, locationNumberCell, 'depositReceiptNo', statusSheetColumns.columnIndexes, 'lendingOnly');
      }
      
      // 数式を行データに設定
      let formulasAdded = 0;
      Object.keys(formulaMapping).forEach(fieldName => {
        if (columnIndexes[fieldName] !== undefined && formulaMapping[fieldName]) {
          rowData[columnIndexes[fieldName]] = formulaMapping[fieldName];
          formulasAdded++;
        }
      });
      
      addFormLog('動的数式追加完了', {
        deviceCategory,
        statusSheetName,
        currentRowNumber,
        locationNumberCell,
        formulasAdded,
        totalFormulaFields: Object.keys(formulaMapping).length,
        detectedStatusColumns: Object.keys(statusSheetColumns.columnIndexes),
        targetColumns: Object.keys(formulaMapping).filter(key => 
          columnIndexes[key] !== undefined
        ),
        generatedFormulas: Object.keys(formulaMapping).reduce((acc, key) => {
          if (columnIndexes[key] !== undefined && formulaMapping[key]) {
            acc[key] = formulaMapping[key];
          }
          return acc;
        }, {})
      });
      
    } else {
      // ステータス収集シートの列検出に失敗した場合は数式なしで続行
      addFormLog('ステータス収集シート列検出失敗、数式なしで続行', {
        statusSheetName,
        error: statusSheetColumns.error,
        note: '基本データのみ追記します'
             });
     }
   } else {
     // ステータス収集シートが不要またはlocationNumber列が見つからない場合
     addFormLog('数式スキップ', {
       reason: !statusSheetName ? 'ステータス収集シート名が取得できません' : '拠点管理番号列が見つかりません',
       deviceCategory: formData.deviceType,
       statusSheetName,
       hasLocationNumberColumn: columnIndexes.locationNumber !== undefined
     });
   }
    
    return rowData;
    
  } catch (error) {
    addFormLog('数式付き行データ作成エラー', {
      error: error.toString(),
      deviceCategory: formData.deviceType,
      note: '基本データのみで行データを作成します'
    });
    
    // エラーが発生した場合は数式なしの基本行データを返す
    return createRowData(formData, additionalData, columnIndexes, totalColumns);
  }
}

/**
 * 行データを作成する関数（動的列対応）
 * @param {Object} formData - フォームデータ
 * @param {Object} additionalData - 追加データ
 * @param {Object} columnIndexes - 列インデックスマッピング
 * @param {number} totalColumns - 総列数
 */
function createRowData(formData, additionalData, columnIndexes, totalColumns) {
  try {
    // 検出された列数に対応するデータ配列を作成（すべて空文字で初期化）
    const rowData = new Array(totalColumns).fill('');
    
    // 設定されたデータの記録用
    const setFields = [];
    
    // 各フィールドのデータを対応する列に設定
    Object.keys(SPREADSHEET_COLUMN_MAPPING).forEach(fieldName => {
      const columnIndex = columnIndexes[fieldName];
      const fieldValue = formData[fieldName];
      
      if (columnIndex !== undefined && fieldValue) {
        rowData[columnIndex] = fieldValue;
        setFields.push({
          field: fieldName,
          value: fieldValue,
          columnIndex: columnIndex
        });
      }
    });
    
    addFormLog('動的行データ作成完了', {
      totalColumns: totalColumns,
      rowDataLength: rowData.length,
      detectedColumns: Object.keys(columnIndexes).length,
      setFields: setFields,
      columnMapping: columnIndexes
    });
    
    return rowData;
    
  } catch (error) {
    addFormLog('行データ作成エラー', {
      error: error.toString(),
      formData,
      additionalData,
      columnIndexes,
      totalColumns
    });
    throw error;
  }
}

/**
 * スプレッドシートIDを設定する関数（管理者用）
 * @param {string} spreadsheetId - 設定するスプレッドシートID
 */
function setSpreadsheetDestination(spreadsheetId) {
  try {
    const properties = PropertiesService.getScriptProperties();
    
    if (spreadsheetId) {
      // スプレッドシートの存在確認
      try {
        SpreadsheetApp.openById(spreadsheetId);
        properties.setProperty(SPREADSHEET_ID_DESTINATION_KEY, spreadsheetId);
        addFormLog('スプレッドシートID設定', { spreadsheetId });
      } catch (spreadsheetError) {
        throw new Error('指定されたスプレッドシートIDが無効です: ' + spreadsheetId);
      }
    } else {
      // スプレッドシートIDをクリア
      properties.deleteProperty(SPREADSHEET_ID_DESTINATION_KEY);
      addFormLog('スプレッドシートID削除');
    }
    
    return {
      success: true,
      message: spreadsheetId ? 'スプレッドシートIDを設定しました' : 'スプレッドシートIDを削除しました',
      spreadsheetId: spreadsheetId
    };
    
  } catch (error) {
    addFormLog('スプレッドシートID設定エラー', {
      error: error.toString(),
      spreadsheetId
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 現在のスプレッドシート設定を取得する関数
 */
/**
 * プリンタ・その他のスプレッドシート連携テスト関数
 * @param {string} testLocationNumber - テスト用拠点管理番号
 * @param {string} testLocation - テスト用拠点
 * @param {string} testDeviceCategory - テスト用詳細カテゴリ（プリンタ または その他）
 */
function testPrinterSpreadsheetIntegration(testLocationNumber = 'PrinterTest_001', testLocation = 'osaka-printer', testDeviceCategory = 'プリンタ') {
  console.log('=== プリンタ・その他 スプレッドシート連携テスト開始 ===');
  
  try {
    // スプレッドシート設定の確認
    const spreadsheetSettings = getSpreadsheetSettings();
    console.log('スプレッドシート設定確認:', spreadsheetSettings);
    
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      console.warn('⚠️ スプレッドシートIDが設定されていません。');
      return {
        success: false,
        message: 'スプレッドシートIDが設定されていません',
        error: 'SPREADSHEET_ID_DESTINATION が設定されていません'
      };
    }
    
    // テスト用プリンタフォームデータ
    const testFormData = {
      title: `プリンタスプレッドシート連携テスト_${testLocationNumber}`,
      description: 'プリンタ・その他のスプレッドシート連携機能テスト用フォーム',
      locationNumber: testLocationNumber,
      deviceType: 'printer',
      deviceCategory: testDeviceCategory,  // 「プリンタ」または「その他」
      location: testLocation,
      // プリンタ用attributes（必要な項目のみ）
      attributes: {
        assetNumber: '',  // プリンタには資産番号なし
        model: 'PRINTER-MODEL-456',
        serial: 'PRINTER-SER456789',
        software: '',  // プリンタにはソフトなし
        os: ''  // プリンタにはOSなし
      }
    };
    
    console.log('プリンタフォーム作成開始:', testFormData);
    
    // フォーム作成（スプレッドシート連携含む）
    const result = createGoogleForm(testFormData);
    
    if (result.success) {
      console.log('✅ プリンタフォーム作成成功');
      console.log('📋 フォーム情報:', {
        formId: result.data.formId,
        title: result.data.title,
        publicUrl: result.data.publicUrl
      });
      
      // スプレッドシート連携結果の確認
      if (result.data.spreadsheet?.success) {
        console.log('✅ プリンタスプレッドシート連携成功');
        console.log('📊 スプレッドシート情報:', {
          sheetName: result.data.spreadsheet.sheetName,
          spreadsheetId: result.data.spreadsheet.spreadsheetId,
          detectedColumns: result.data.spreadsheet.detectedColumns,
          deviceCategory: testDeviceCategory
        });
        
        return {
          success: true,
          message: 'プリンタ・その他のスプレッドシート連携テスト完了',
          formData: {
            formId: result.data.formId,
            publicUrl: result.data.publicUrl
          },
          spreadsheetData: {
            sheetName: result.data.spreadsheet.sheetName,
            spreadsheetId: result.data.spreadsheet.spreadsheetId,
            rowAdded: true,
            deviceCategory: testDeviceCategory
          }
        };
        
      } else {
        console.error('❌ プリンタスプレッドシート連携失敗:', result.data.spreadsheet?.error);
        return {
          success: false,
          message: 'プリンタフォームは作成されましたが、スプレッドシート連携に失敗しました',
          error: result.data.spreadsheet?.error
        };
      }
      
    } else {
      console.error('❌ プリンタフォーム作成失敗:', result.error);
      return {
        success: false,
        message: 'プリンタフォーム作成に失敗しました',
        error: result.error
      };
    }
    
  } catch (error) {
    console.error('プリンタスプレッドシート連携テストエラー:', error.toString());
    return {
      success: false,
      message: 'プリンタテストでエラーが発生しました',
      error: error.toString()
    };
  }
}

/**
 * ステータス収集シートの詳細状況確認関数
 * @param {string} statusSheetName - 確認対象のシート名
 */
function debugStatusSheetStructure(statusSheetName = '端末ステータス収集シート') {
  console.log('=== ステータス収集シート詳細確認開始 ===');
  console.log('対象シート:', statusSheetName);
  
  try {
    const spreadsheetSettings = getSpreadsheetSettings();
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      console.error('❌ スプレッドシート設定が見つかりません');
      return { success: false, error: 'スプレッドシート設定なし' };
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetSettings.settings.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(statusSheetName);
    
    if (!sheet) {
      console.error('❌ シートが見つかりません:', statusSheetName);
      return { success: false, error: 'シートが見つかりません' };
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    console.log('📊 シート基本情報:');
    console.log('  - 最終行:', lastRow);
    console.log('  - 最終列:', lastColumn);
    console.log('  - シートURL:', spreadsheet.getUrl() + '#gid=' + sheet.getSheetId());
    
    if (lastRow === 0 || lastColumn === 0) {
      console.warn('⚠️ シートが空です');
      return { success: false, error: 'シートが空' };
    }
    
    // 1行目と3行目のヘッダーを確認
    const candidateRows = [1, 3].filter(rowNum => rowNum <= lastRow);
    console.log('📋 ヘッダー候補行:', candidateRows);
    
    for (const rowNum of candidateRows) {
      console.log(`\n--- ${rowNum}行目の内容 ---`);
      const headerRow = sheet.getRange(rowNum, 1, 1, lastColumn).getValues()[0];
      
      console.log('列数:', headerRow.length);
      console.log('内容（最初の15列）:', headerRow.slice(0, 15));
      
      // STATUS_SHEET_COLUMN_MAPPINGとのマッチング確認
      console.log('\n🔍 列名マッチング確認:');
      Object.keys(STATUS_SHEET_COLUMN_MAPPING).forEach(fieldName => {
        const candidateNames = STATUS_SHEET_COLUMN_MAPPING[fieldName];
        let found = false;
        
        for (let i = 0; i < headerRow.length; i++) {
          const headerValue = headerRow[i] ? headerRow[i].toString().trim() : '';
          
          for (const candidateName of candidateNames) {
            if (headerValue === candidateName) {
              console.log(`  ✅ ${fieldName}: "${candidateName}" → ${getColumnLetter(i + 1)}列`);
              found = true;
              break;
            }
          }
          
          if (found) break;
        }
        
        if (!found) {
          console.log(`  ❌ ${fieldName}: 見つかりません (候補: ${candidateNames.join(', ')})`);
        }
      });
    }
    
    // 列検出のテスト実行
    console.log('\n🧪 実際の列検出テスト:');
    const result = detectStatusSheetColumns(statusSheetName);
    console.log('検出結果:', result);
    
    return {
      success: true,
      sheetInfo: {
        lastRow,
        lastColumn,
        candidateRows
      },
      detectionResult: result
    };
    
  } catch (error) {
    console.error('❌ シート確認エラー:', error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * 数式生成の詳細テスト関数
 * @param {string} testLocationNumber - テスト用拠点管理番号
 * @param {string} testDeviceCategory - テスト用詳細カテゴリ
 */
function testFormulaGeneration(testLocationNumber = 'Formula_Test_001', testDeviceCategory = 'SV') {
  console.log('=== 数式生成詳細テスト開始 ===');
  
  try {
    const statusSheetName = getStatusCollectionSheetName(testDeviceCategory);
    console.log('対象ステータス収集シート:', statusSheetName);
    
    // ステータス収集シートの列検出
    const statusSheetColumns = detectStatusSheetColumns(statusSheetName);
    
    if (!statusSheetColumns.success) {
      console.error('❌ ステータス収集シート列検出失敗:', statusSheetColumns.error);
      return { success: false, error: statusSheetColumns.error };
    }
    
    console.log('✅ ステータス収集シート列検出成功:', statusSheetColumns.columnIndexes);
    
    // テスト用のセル参照
    const testLocationCell = 'Q10';  // 拠点管理番号があるセル
    
    // 各数式を生成してテスト
    const formulaTests = {
      assignee: generateDynamicVlookupFormula(statusSheetName, testLocationCell, 'assignee', statusSheetColumns.columnIndexes),
      lendingStatus: generateDynamicVlookupFormula(statusSheetName, testLocationCell, 'status', statusSheetColumns.columnIndexes),
      lendingDestination: generateDynamicVlookupFormula(statusSheetName, testLocationCell, 'destination', statusSheetColumns.columnIndexes),
      lendingDate: generateDynamicVlookupFormula(statusSheetName, testLocationCell, 'timestamp', statusSheetColumns.columnIndexes, 'lendingOnly'),
      userDeviceDeposit: generateDynamicVlookupFormula(statusSheetName, testLocationCell, 'userDeviceDeposit', statusSheetColumns.columnIndexes, 'lendingOnly'),
      remarks: generateDynamicVlookupFormula(statusSheetName, testLocationCell, 'remarks', statusSheetColumns.columnIndexes, 'lendingOnly')
    };
    
    // 端末の場合はお預かり証No.も追加
    if (testDeviceCategory === 'SV' || testDeviceCategory === 'CL') {
      formulaTests.depositReceiptNo = generateDynamicVlookupFormula(statusSheetName, testLocationCell, 'depositReceiptNo', statusSheetColumns.columnIndexes, 'lendingOnly');
    }
    
    console.log('📋 生成された数式一覧:');
    Object.keys(formulaTests).forEach(fieldName => {
      const formula = formulaTests[fieldName];
      console.log(`  - ${fieldName}: ${formula || '（生成失敗）'}`);
    });
    
    // 数式の構文チェック
    const validFormulas = Object.keys(formulaTests).filter(key => 
      formulaTests[key] && formulaTests[key].startsWith('=')
    );
    
    console.log(`✅ 有効な数式: ${validFormulas.length}個`);
    console.log(`❌ 無効な数式: ${Object.keys(formulaTests).length - validFormulas.length}個`);
    
    return {
      success: true,
      statusSheetName,
      detectedColumns: statusSheetColumns.columnIndexes,
      generatedFormulas: formulaTests,
      validFormulaCount: validFormulas.length,
      totalFormulaCount: Object.keys(formulaTests).length
    };
    
  } catch (error) {
    console.error('数式生成テストエラー:', error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * ステータス収集シートの動的列検出テスト関数
 * @param {string} statusSheetName - テスト対象のステータス収集シート名
 */
function testStatusSheetColumnDetection(statusSheetName = '端末ステータス収集シート') {
  console.log('=== ステータス収集シート動的列検出テスト開始 ===');
  console.log('対象シート:', statusSheetName);
  
  try {
    // ステータス収集シートの列検出を実行
    const result = detectStatusSheetColumns(statusSheetName);
    
    if (result.success) {
      console.log('✅ 動的列検出成功');
      console.log('📊 検出結果:', {
        headerRowIndex: result.headerRowIndex,
        detectedColumns: result.columnIndexes,
        columnCount: Object.keys(result.columnIndexes).length
      });
      
      // 各列の詳細表示
      console.log('📋 検出された列マッピング:');
      Object.keys(result.columnIndexes).forEach(fieldName => {
        const column = result.columnIndexes[fieldName];
        const mapping = STATUS_SHEET_COLUMN_MAPPING[fieldName];
        console.log(`  - ${fieldName}: ${column}列 (候補: ${mapping.join(', ')})`);
      });
      
      // 数式生成のテスト
      console.log('🧪 数式生成テスト:');
      const testLocationCell = 'Q10';  // テスト用セル参照
      
      Object.keys(result.columnIndexes).forEach(fieldName => {
        const formula = generateDynamicVlookupFormula(statusSheetName, testLocationCell, fieldName, result.columnIndexes);
        console.log(`  - ${fieldName}: ${formula || '（数式生成なし）'}`);
      });
      
      // 条件付き数式のテスト
      console.log('🧪 条件付き数式生成テスト:');
      const conditionalFields = ['timestamp', 'userDeviceDeposit', 'depositReceiptNo', 'remarks'];
      conditionalFields.forEach(fieldName => {
        if (result.columnIndexes[fieldName]) {
          const formula = generateDynamicVlookupFormula(statusSheetName, testLocationCell, fieldName, result.columnIndexes, 'lendingOnly');
          console.log(`  - ${fieldName} (貸出中のみ): ${formula || '（数式生成なし）'}`);
        }
      });
      
      return {
        success: true,
        message: 'ステータス収集シート動的列検出テスト完了',
        detectedColumns: result.columnIndexes,
        headerRowIndex: result.headerRowIndex,
        columnCount: Object.keys(result.columnIndexes).length
      };
      
    } else {
      console.error('❌ 動的列検出失敗:', result.error);
      return {
        success: false,
        message: 'ステータス収集シート動的列検出に失敗しました',
        error: result.error
      };
    }
    
  } catch (error) {
    console.error('ステータス収集シート動的列検出テストエラー:', error.toString());
    return {
      success: false,
      message: 'ステータス収集シート動的列検出テストでエラーが発生しました',
      error: error.toString()
    };
  }
}

/**
 * 重複チェック機能のテスト関数
 * @param {string} testDeviceCategory - テスト対象デバイスカテゴリ（SV, CL, プリンタ, その他）
 * @param {Object} testData - テストデータ
 */
function testDuplicateCheck(testDeviceCategory = 'SV', testData = null) {
  console.log('=== 重複チェック機能テスト開始 ===');
  console.log('対象デバイスカテゴリ:', testDeviceCategory);
  
  // デフォルトテストデータ
  const defaultTestData = {
    locationNumber: 'TEST_001',
    assetNumber: 'ASSET_001', 
    serial: 'SERIAL_001'
  };
  
  const checkData = testData || defaultTestData;
  console.log('チェック対象データ:', checkData);
  
  try {
    const result = checkDuplicateValues(testDeviceCategory, checkData);
    
    if (result.success) {
      console.log('✅ 重複チェック実行成功');
      console.log('📊 チェック結果:', {
        sheetName: result.sheetName,
        checkedFields: result.checkedFields,
        hasDuplicates: result.hasDuplicates,
        duplicatesCount: result.duplicates?.length || 0,
        totalDataRows: result.totalDataRows
      });
      
      if (result.hasDuplicates) {
        console.log('⚠️ 重複データが見つかりました:');
        result.duplicates.forEach((duplicate, index) => {
          console.log(`  ${index + 1}. ${duplicate.field}: "${duplicate.value}" (行${duplicate.rowNumber}, ${duplicate.columnLetter}列)`);
        });
      } else {
        console.log('✅ 重複データはありませんでした');
      }
      
      return {
        success: true,
        message: '重複チェック機能テスト完了',
        result: result
      };
      
    } else {
      console.error('❌ 重複チェック実行失敗:', result.error);
      return {
        success: false,
        message: '重複チェック機能テストに失敗しました',
        error: result.error
      };
    }
    
  } catch (error) {
    console.error('重複チェック機能テストエラー:', error.toString());
    return {
      success: false,
      message: '重複チェック機能テストでエラーが発生しました',
      error: error.toString()
    };
  }
}

/**
 * 動的数式機能のテスト関数
 * @param {string} testLocationNumber - テスト用拠点管理番号
 * @param {string} testDeviceCategory - テスト用詳細カテゴリ（SV, CL, プリンタ, その他）
 */
function testDynamicFormulaIntegration(testLocationNumber = 'DynamicTest_001', testDeviceCategory = 'SV') {
  console.log('=== 動的数式機能テスト開始 ===');
  
  try {
    // スプレッドシート設定の確認
    const spreadsheetSettings = getSpreadsheetSettings();
    console.log('スプレッドシート設定確認:', spreadsheetSettings);
    
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      console.warn('⚠️ スプレッドシートIDが設定されていません。');
      return {
        success: false,
        message: 'スプレッドシートIDが設定されていません',
        error: 'SPREADSHEET_ID_DESTINATION が設定されていません'
      };
    }
    
    // デバイスタイプを決定
    const deviceType = (testDeviceCategory === 'SV' || testDeviceCategory === 'CL') ? 'terminal' : 'printer';
    const location = deviceType === 'terminal' ? 'osaka-desktop' : 'osaka-printer';
    const statusSheetName = getStatusCollectionSheetName(testDeviceCategory);
    
    // まず動的列検出をテスト
    console.log('📊 ステータス収集シート動的列検出実行中...');
    const columnDetectionResult = testStatusSheetColumnDetection(statusSheetName);
    
         if (!columnDetectionResult.success) {
       console.error('❌ 動的列検出に失敗しました。フォーム作成時にエラーが発生する可能性があります');
       console.warn('⚠️ ステータス収集シートの列構成を確認してください');
     }
    
    // テスト用フォームデータ
    const testFormData = {
      title: `動的数式機能テスト_${testLocationNumber}_${testDeviceCategory}`,
      description: '動的数式機能のテスト用フォーム（ステータス収集シートから動的に列を検出）',
      locationNumber: testLocationNumber,
      deviceType: deviceType,
      deviceCategory: testDeviceCategory,
      location: location,
      // テスト用attributes
      attributes: {
        assetNumber: deviceType === 'terminal' ? 'DYNAMIC-ASSET-001' : '',
        model: 'DYNAMIC-MODEL-123',
        serial: 'DYNAMIC-SER456789',
        software: deviceType === 'terminal' ? '動的数式テストソフト' : '',
        os: deviceType === 'terminal' ? 'Windows 11' : ''
      }
    };
    
    console.log('動的数式テスト用フォーム作成開始:', testFormData);
    
    // フォーム作成（動的数式機能含む）
    const result = createGoogleForm(testFormData);
    
    if (result.success) {
      console.log('✅ 動的数式テスト用フォーム作成成功');
      console.log('📋 フォーム情報:', {
        formId: result.data.formId,
        title: result.data.title,
        publicUrl: result.data.publicUrl
      });
      
      // スプレッドシート連携結果の確認
      if (result.data.spreadsheet?.success) {
        console.log('✅ 動的数式機能付きスプレッドシート連携成功');
        console.log('📊 スプレッドシート情報:', {
          sheetName: result.data.spreadsheet.sheetName,
          spreadsheetId: result.data.spreadsheet.spreadsheetId,
          detectedColumns: result.data.spreadsheet.detectedColumns,
          deviceCategory: testDeviceCategory,
          statusSheetName: statusSheetName
        });
        
                 console.log('📋 動的数式機能の特徴:');
         console.log('  ✨ ステータス収集シートのヘッダーから自動で列を検出');
         console.log('  ✨ 列の順序が変わっても対応可能');
         console.log('  ✨ 列名の変更に柔軟に対応（候補名リスト使用）');
         console.log('  ✨ 列検出失敗時は明確なエラーメッセージを表示');
         console.log('  ✨ デバイス種別に応じたシート自動選択');
        
        return {
          success: true,
          message: '動的数式機能テスト完了',
          formData: {
            formId: result.data.formId,
            publicUrl: result.data.publicUrl
          },
          spreadsheetData: {
            sheetName: result.data.spreadsheet.sheetName,
            spreadsheetId: result.data.spreadsheet.spreadsheetId,
            rowAdded: true,
            deviceCategory: testDeviceCategory,
            statusSheetName: statusSheetName,
            dynamicFormulaEnabled: true,
            columnDetectionResult: columnDetectionResult
          }
        };
        
      } else {
        console.error('❌ 動的数式機能付きスプレッドシート連携失敗:', result.data.spreadsheet?.error);
        return {
          success: false,
          message: 'フォームは作成されましたが、動的数式機能付きスプレッドシート連携に失敗しました',
          error: result.data.spreadsheet?.error
        };
      }
      
    } else {
      console.error('❌ 動的数式テスト用フォーム作成失敗:', result.error);
      return {
        success: false,
        message: '動的数式テスト用フォーム作成に失敗しました',
        error: result.error
      };
    }
    
  } catch (error) {
    console.error('動的数式機能テストエラー:', error.toString());
    return {
      success: false,
      message: '動的数式機能テストでエラーが発生しました',
      error: error.toString()
    };
  }
}

/**
 * 数式機能のテスト関数
 * @param {string} testLocationNumber - テスト用拠点管理番号
 * @param {string} testDeviceCategory - テスト用詳細カテゴリ（SV, CL, プリンタ, その他）
 */
function testFormulaIntegration(testLocationNumber = 'FormulaTest_001', testDeviceCategory = 'SV') {
  console.log('=== 数式機能テスト開始 ===');
  
  try {
    // スプレッドシート設定の確認
    const spreadsheetSettings = getSpreadsheetSettings();
    console.log('スプレッドシート設定確認:', spreadsheetSettings);
    
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      console.warn('⚠️ スプレッドシートIDが設定されていません。');
      return {
        success: false,
        message: 'スプレッドシートIDが設定されていません',
        error: 'SPREADSHEET_ID_DESTINATION が設定されていません'
      };
    }
    
    // デバイスタイプを決定
    const deviceType = (testDeviceCategory === 'SV' || testDeviceCategory === 'CL') ? 'terminal' : 'printer';
    const location = deviceType === 'terminal' ? 'osaka-desktop' : 'osaka-printer';
    
    // テスト用フォームデータ
    const testFormData = {
      title: `数式機能テスト_${testLocationNumber}_${testDeviceCategory}`,
      description: '数式機能のテスト用フォーム',
      locationNumber: testLocationNumber,
      deviceType: deviceType,
      deviceCategory: testDeviceCategory,
      location: location,
      // テスト用attributes
      attributes: {
        assetNumber: deviceType === 'terminal' ? 'FORMULA-ASSET-001' : '',
        model: 'FORMULA-MODEL-123',
        serial: 'FORMULA-SER456789',
        software: deviceType === 'terminal' ? '数式テストソフト' : '',
        os: deviceType === 'terminal' ? 'Windows 11' : ''
      }
    };
    
    console.log('数式テスト用フォーム作成開始:', testFormData);
    
    // フォーム作成（数式機能含む）
    const result = createGoogleForm(testFormData);
    
    if (result.success) {
      console.log('✅ 数式テスト用フォーム作成成功');
      console.log('📋 フォーム情報:', {
        formId: result.data.formId,
        title: result.data.title,
        publicUrl: result.data.publicUrl
      });
      
      // スプレッドシート連携結果の確認
      if (result.data.spreadsheet?.success) {
        console.log('✅ 数式機能付きスプレッドシート連携成功');
        console.log('📊 スプレッドシート情報:', {
          sheetName: result.data.spreadsheet.sheetName,
          spreadsheetId: result.data.spreadsheet.spreadsheetId,
          detectedColumns: result.data.spreadsheet.detectedColumns,
          deviceCategory: testDeviceCategory,
          statusSheetName: getStatusCollectionSheetName(testDeviceCategory)
        });
        
        console.log('📋 追加された数式情報:');
        console.log('  - 担当者: ステータス収集シートB列から取得');
        console.log('  - 貸出ステータス: ステータス収集シートE列から取得');
        console.log('  - 貸出先: ステータス収集シートI列から取得');
        console.log('  - 貸出日: 貸出中の場合のみC列から取得');
        console.log('  - ユーザー預り機有: 貸出中の場合のみL列から取得');
        if (testDeviceCategory === 'SV' || testDeviceCategory === 'CL') {
          console.log('  - お預かり証No.: 端末かつ貸出中の場合のみP列から取得');
        }
        console.log('  - 備考: 貸出中の場合のみN列から取得');
        
        return {
          success: true,
          message: '数式機能テスト完了',
          formData: {
            formId: result.data.formId,
            publicUrl: result.data.publicUrl
          },
          spreadsheetData: {
            sheetName: result.data.spreadsheet.sheetName,
            spreadsheetId: result.data.spreadsheet.spreadsheetId,
            rowAdded: true,
            deviceCategory: testDeviceCategory,
            statusSheetName: getStatusCollectionSheetName(testDeviceCategory),
            formulasEnabled: true
          }
        };
        
      } else {
        console.error('❌ 数式機能付きスプレッドシート連携失敗:', result.data.spreadsheet?.error);
        return {
          success: false,
          message: 'フォームは作成されましたが、数式機能付きスプレッドシート連携に失敗しました',
          error: result.data.spreadsheet?.error
        };
      }
      
    } else {
      console.error('❌ 数式テスト用フォーム作成失敗:', result.error);
      return {
        success: false,
        message: '数式テスト用フォーム作成に失敗しました',
        error: result.error
      };
    }
    
  } catch (error) {
    console.error('数式機能テストエラー:', error.toString());
    return {
      success: false,
      message: '数式機能テストでエラーが発生しました',
      error: error.toString()
    };
  }
}

function getSpreadsheetSettings() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty(SPREADSHEET_ID_DESTINATION_KEY);
    
    let spreadsheetInfo = null;
    if (spreadsheetId) {
      try {
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        spreadsheetInfo = {
          id: spreadsheetId,
          name: spreadsheet.getName(),
          url: spreadsheet.getUrl(),
          sheets: spreadsheet.getSheets().map(sheet => ({
            name: sheet.getName(),
            rows: sheet.getLastRow(),
            columns: sheet.getLastColumn()
          }))
        };
      } catch (spreadsheetError) {
        addFormLog('スプレッドシート情報取得エラー', { spreadsheetId, error: spreadsheetError.toString() });
      }
    }
    
    return {
      success: true,
      settings: {
        spreadsheetId: spreadsheetId,
        spreadsheetInfo: spreadsheetInfo,
        locationSheetNames: LOCATION_SHEET_NAMES,
        columnMapping: SPREADSHEET_COLUMN_MAPPING
      }
    };
    
  } catch (error) {
    addFormLog('スプレッドシート設定取得エラー', {
      error: error.toString()
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * スプレッドシート連携機能のテスト関数
 * @param {string} testLocationNumber - テスト用拠点管理番号
 * @param {string} testLocation - テスト用拠点識別子
 */
function testSpreadsheetIntegration(testLocationNumber = 'SpreadTest_001', testLocation = 'osaka-desktop') {
  console.log('=== スプレッドシート連携テスト開始 ===');
  
  try {
    // スプレッドシート設定の確認
    const spreadsheetSettings = getSpreadsheetSettings();
    console.log('スプレッドシート設定確認:', spreadsheetSettings);
    
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      console.warn('⚠️ スプレッドシートIDが設定されていません。');
      console.log('設定例:');
      console.log('setSpreadsheetDestination("YOUR_SPREADSHEET_ID")');
      return {
        success: false,
        message: 'スプレッドシートIDが設定されていません',
        error: 'SPREADSHEET_ID_DESTINATION が設定されていません'
      };
    }
    
             // テスト用フォームデータ
    const testFormData = {
      title: `スプレッドシート連携テスト_${testLocationNumber}`,
      description: 'スプレッドシート連携機能のテスト用フォーム',
      locationNumber: testLocationNumber,
      deviceType: 'terminal',
      deviceCategory: 'SV',  // SV/CL形式で設定
      location: testLocation,
      // attributesオブジェクトに設定（フロントエンドと同じ構造）
      attributes: {
        assetNumber: 'ASSET-TEST-001',
        model: 'MODEL-TEST-123',
        serial: 'SER123456789',
        software: 'テストソフトウェア',
        os: 'Windows 11'
      }
    };
    
    console.log('フォーム作成開始:', testFormData);
    
    // フォーム作成（スプレッドシート連携含む）
    const result = createGoogleForm(testFormData);
    
    if (result.success) {
      console.log('✅ フォーム作成成功');
      console.log('📋 フォーム情報:', {
        formId: result.data.formId,
        title: result.data.title,
        publicUrl: result.data.publicUrl
      });
      
      // スプレッドシート連携結果の確認
      if (result.data.spreadsheet?.success) {
        console.log('✅ スプレッドシート連携成功');
        console.log('📊 スプレッドシート情報:', {
          sheetName: result.data.spreadsheet.sheetName,
          spreadsheetId: result.data.spreadsheet.spreadsheetId,
          rowDataLength: result.data.spreadsheet.rowData?.length
        });
        
        // 実際のスプレッドシートを確認
        try {
          const spreadsheet = SpreadsheetApp.openById(result.data.spreadsheet.spreadsheetId);
          const sheet = spreadsheet.getSheetByName(result.data.spreadsheet.sheetName);
          
          if (sheet) {
            const lastRow = sheet.getLastRow();
            const lastColumn = sheet.getLastColumn();
            const addedRowData = sheet.getRange(lastRow, 1, 1, lastColumn).getValues()[0];
            
            console.log('📋 追加された行データ:');
            if (result.data.spreadsheet.detectedColumns) {
              result.data.spreadsheet.detectedColumns.forEach(fieldName => {
                const columnMapping = SPREADSHEET_COLUMN_MAPPING[fieldName];
                if (columnMapping) {
                  console.log(`  ${fieldName} (${columnMapping[0]}): データあり`);
                }
              });
            }
            
            console.log('📊 シート統計:', {
              totalRows: lastRow,
              totalColumns: lastColumn,
              sheetName: sheet.getName(),
              detectedColumns: result.data.spreadsheet.detectedColumns?.length || 0
            });
          }
          
        } catch (verifyError) {
          console.error('スプレッドシート確認エラー:', verifyError.toString());
        }
        
        return {
          success: true,
          message: 'スプレッドシート連携テスト完了',
          formData: {
            formId: result.data.formId,
            publicUrl: result.data.publicUrl
          },
          spreadsheetData: {
            sheetName: result.data.spreadsheet.sheetName,
            spreadsheetId: result.data.spreadsheet.spreadsheetId,
            rowAdded: true
          }
        };
        
      } else {
        console.error('❌ スプレッドシート連携失敗:', result.data.spreadsheet?.error);
        return {
          success: false,
          message: 'フォームは作成されましたが、スプレッドシート連携に失敗しました',
          error: result.data.spreadsheet?.error,
          formData: {
            formId: result.data.formId,
            publicUrl: result.data.publicUrl
          }
        };
      }
      
    } else {
      console.error('❌ フォーム作成失敗:', result.error);
      return {
        success: false,
        message: 'フォーム作成に失敗しました',
        error: result.error
      };
    }
    
  } catch (error) {
    console.error('スプレッドシート連携テストエラー:', error.toString());
    return {
      success: false,
      message: 'テストでエラーが発生しました',
      error: error.toString()
    };
  }
}

/**
 * フォーム作成とスプレッドシート連携の統合テスト
 * @param {string} testLocationNumber - テスト用拠点管理番号
 */
function testFormWithSpreadsheet(testLocationNumber = 'IntegrationTest_001') {
  console.log('=== フォーム＋スプレッドシート統合テスト開始 ===');
  
  try {
    // 設定確認
    const spreadsheetSettings = getSpreadsheetSettings();
    const qrCodeSettings = getQRCodeFolderSettings();
    
    console.log('📊 現在の設定状況:');
    console.log('  スプレッドシート設定:', spreadsheetSettings.success);
    console.log('  QRコード設定:', qrCodeSettings.success);
    
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      console.warn('⚠️ スプレッドシートIDが設定されていません');
      console.log('設定コマンド例:');
      console.log('setSpreadsheetDestination("YOUR_SPREADSHEET_ID")');
    }
    
    // 対象拠点をテスト（プリンター関連を除外）
    const allLocations = Object.keys(LOCATION_SHEET_NAMES);
    const testResults = [];
    
    for (const [index, location] of allLocations.entries()) {
      console.log(`\n📍 拠点テスト ${index + 1}/${allLocations.length}: ${location} (${LOCATION_SHEET_NAMES[location]})`);
      
             const locationTestData = {
         title: `統合テスト_${location}_${testLocationNumber}`,
         description: `${LOCATION_SHEET_NAMES[location]}の統合テスト用フォーム`,
         locationNumber: `${testLocationNumber}_${location}`,
         deviceType: 'terminal', // すべてterminalに統一
         location: location,
         // 必要な列のみ設定
         assetNumber: `ASSET-${location.toUpperCase()}-${Date.now().toString().slice(-6)}`,
         modelNumber: `MODEL-${location.toUpperCase()}-TEST`,
         software: 'テストソフトウェア',
         os: 'Windows 11'
       };
      
      try {
        const result = createGoogleForm(locationTestData);
        
        testResults.push({
          location: location,
          locationName: LOCATION_SHEET_NAMES[location],
          formCreated: result.success,
          formId: result.success ? result.data.formId : null,
          spreadsheetLinked: result.success ? result.data.spreadsheet?.success : false,
          qrCodeGenerated: result.success ? result.data.qrCode?.success : false,
          error: result.success ? null : result.error
        });
        
        if (result.success) {
          console.log(`  ✅ 成功: フォーム作成済み`);
          console.log(`  📊 スプレッドシート: ${result.data.spreadsheet?.success ? '成功' : '失敗'}`);
          console.log(`  📱 QRコード: ${result.data.qrCode?.success ? '成功' : '失敗'}`);
          
          // スプレッドシート連携失敗の詳細ログ
          if (!result.data.spreadsheet?.success) {
            console.log(`    スプレッドシートエラー: ${result.data.spreadsheet?.error}`);
          }
        } else {
          console.log(`  ❌ 失敗: ${result.error}`);
        }
        
        // 間隔を空ける（API制限対策）
        if (index < allLocations.length - 1) {
          Utilities.sleep(2000);
        }
        
      } catch (locationError) {
        console.log(`  💥 拠点テスト例外: ${locationError.toString()}`);
        testResults.push({
          location: location,
          locationName: LOCATION_SHEET_NAMES[location],
          formCreated: false,
          error: locationError.toString()
        });
      }
    }
    
    // 結果サマリー
    console.log('\n=== 統合テスト結果サマリー ===');
    const successCount = testResults.filter(r => r.formCreated).length;
    const spreadsheetCount = testResults.filter(r => r.spreadsheetLinked).length;
    const qrCodeCount = testResults.filter(r => r.qrCodeGenerated).length;
    
    console.log(`📊 フォーム作成: ${successCount}/${testResults.length}件成功`);
    console.log(`📊 スプレッドシート連携: ${spreadsheetCount}/${testResults.length}件成功`);
    console.log(`📊 QRコード生成: ${qrCodeCount}/${testResults.length}件成功`);
    
    console.log('\n📋 拠点別結果:');
    testResults.forEach(result => {
      const status = result.formCreated ? '✅' : '❌';
      const sheetStatus = result.spreadsheetLinked ? '📊✅' : '📊❌';
      console.log(`${status} ${sheetStatus} ${result.location} (${result.locationName})`);
      if (result.error) {
        console.log(`    エラー: ${result.error}`);
      }
    });
    
    return {
      success: successCount === testResults.length,
      message: `統合テスト完了: ${successCount}/${testResults.length}件成功`,
      results: testResults,
      summary: {
        total: testResults.length,
        formCreated: successCount,
        spreadsheetLinked: spreadsheetCount,
        qrCodeGenerated: qrCodeCount,
        locations: allLocations
      }
    };
    
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
 * スプレッドシートのシート存在確認テスト関数
 */
function testSpreadsheetSheetExistence() {
  console.log('=== スプレッドシートシート存在確認テスト開始 ===');
  
  try {
    // スプレッドシート設定の確認
    const spreadsheetSettings = getSpreadsheetSettings();
    console.log('スプレッドシート設定確認:', spreadsheetSettings);
    
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      console.warn('⚠️ スプレッドシートIDが設定されていません。');
      return {
        success: false,
        message: 'スプレッドシートIDが設定されていません',
        error: 'SPREADSHEET_ID_DESTINATION が設定されていません'
      };
    }
    
    const spreadsheetId = spreadsheetSettings.settings.spreadsheetId;
    console.log('対象スプレッドシートID:', spreadsheetId);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log('スプレッドシート名:', spreadsheet.getName());
    
    // 既存のシート一覧を取得
    const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    console.log('既存シート一覧:', existingSheets);
    
    // 必要なシートの存在確認
    const requiredSheets = Object.values(LOCATION_SHEET_NAMES);
    const uniqueRequiredSheets = [...new Set(requiredSheets)]; // 重複を削除
    console.log('必要なシート:', uniqueRequiredSheets);
    
    const sheetCheckResults = [];
    
    uniqueRequiredSheets.forEach(sheetName => {
      const exists = existingSheets.includes(sheetName);
      const usedByLocations = Object.keys(LOCATION_SHEET_NAMES).filter(location => LOCATION_SHEET_NAMES[location] === sheetName);
      
      sheetCheckResults.push({
        sheetName: sheetName,
        exists: exists,
        usedByLocations: usedByLocations
      });
      
      const status = exists ? '✅' : '❌';
      console.log(`${status} シート「${sheetName}」 (使用拠点: ${usedByLocations.join(', ')})`);
    });
    
    // 拠点別の確認
    console.log('\n📍 拠点別シート確認:');
    Object.keys(LOCATION_SHEET_NAMES).forEach(location => {
      const sheetName = LOCATION_SHEET_NAMES[location];
      const exists = existingSheets.includes(sheetName);
      const status = exists ? '✅' : '❌';
      console.log(`${status} ${location} → シート「${sheetName}」`);
    });
    
    // 不足しているシート
    const missingSheets = uniqueRequiredSheets.filter(sheetName => !existingSheets.includes(sheetName));
    
    if (missingSheets.length > 0) {
      console.log('\n⚠️ 不足しているシート:');
      missingSheets.forEach(sheetName => {
        const affectedLocations = Object.keys(LOCATION_SHEET_NAMES).filter(location => LOCATION_SHEET_NAMES[location] === sheetName);
        console.log(`- 「${sheetName}」 (影響拠点: ${affectedLocations.join(', ')})`);
      });
      
      console.log('\n📝 シート作成方法:');
      console.log('1. スプレッドシートを手動で開く');
      console.log('2. 以下のシート名で新しいシートを作成:');
      missingSheets.forEach(sheetName => {
        console.log(`   - ${sheetName}`);
      });
      console.log('3. 各シートにヘッダー行を設定（推奨列名）:');
      Object.keys(SPREADSHEET_COLUMN_MAPPING).forEach(fieldName => {
        const columnNames = SPREADSHEET_COLUMN_MAPPING[fieldName];
        console.log(`   ${fieldName}: ${columnNames.join(' または ')}`);
      });
    }
    
    const allSheetsExist = missingSheets.length === 0;
    
    return {
      success: allSheetsExist,
      message: allSheetsExist ? 'すべての必要なシートが存在します' : `${missingSheets.length}個のシートが不足しています`,
      spreadsheetInfo: {
        id: spreadsheetId,
        name: spreadsheet.getName(),
        url: spreadsheet.getUrl()
      },
      existingSheets: existingSheets,
      requiredSheets: uniqueRequiredSheets,
      missingSheets: missingSheets,
      sheetCheckResults: sheetCheckResults,
      locationMapping: LOCATION_SHEET_NAMES
    };
    
  } catch (error) {
    console.error('シート存在確認テストエラー:', error.toString());
    return {
      success: false,
      message: 'シート存在確認テストでエラーが発生しました',
      error: error.toString()
    };
  }
}

/**
 * 自動ヘッダー追加機能のテスト関数
 * @param {string} testSheetName - テスト対象のシート名（オプション）
 */
function testAutoHeaderAddition(testSheetName = null) {
  console.log('=== 自動ヘッダー追加機能テスト開始 ===');
  
  try {
    // スプレッドシート設定の確認
    const spreadsheetSettings = getSpreadsheetSettings();
    
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      console.warn('⚠️ スプレッドシートIDが設定されていません。');
      return {
        success: false,
        message: 'スプレッドシートIDが設定されていません',
        error: 'SPREADSHEET_ID_DESTINATION が設定されていません'
      };
    }
    
    const spreadsheetId = spreadsheetSettings.settings.spreadsheetId;
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    console.log('スプレッドシート名:', spreadsheet.getName());
    
    // テスト対象シートを決定
    const targetSheets = testSheetName ? [testSheetName] : [...new Set(Object.values(LOCATION_SHEET_NAMES))];
    console.log('テスト対象シート:', targetSheets);
    
    const testResults = [];
    
    for (const sheetName of targetSheets) {
      console.log(`\n📊 シート「${sheetName}」の自動ヘッダー追加テスト:`);
      
      try {
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
          console.log('❌ シートが見つかりません');
          testResults.push({
            sheetName: sheetName,
            exists: false,
            error: 'シートが見つかりません'
          });
          continue;
        }
        
        // テスト前の状態を記録
        const beforeState = {
          lastRow: sheet.getLastRow(),
          lastColumn: sheet.getLastColumn(),
          headerInfo: detectHeaderRow(sheet)
        };
        
        console.log('  📋 テスト前の状態:');
        console.log(`    最終行: ${beforeState.lastRow}, 最終列: ${beforeState.lastColumn}`);
        console.log(`    検出されたヘッダー: ${beforeState.headerInfo.found ? Object.keys(beforeState.headerInfo.columnIndexes).length : 0}個`);
        
        // 自動ヘッダー追加機能を実行（テスト用にSVを仮定）
        const headerResult = detectAndEnsureHeaders(sheet, 'SV');
        
        if (headerResult.found) {
          console.log('  ✅ ヘッダー検出・追加成功');
          console.log(`    ヘッダー行: ${headerResult.rowIndex}行目`);
          console.log(`    検出された列: ${Object.keys(headerResult.columnIndexes).length}個`);
          console.log(`    ヘッダー追加: ${headerResult.headersAdded ? 'あり' : 'なし'}`);
          
          if (headerResult.headersAdded && headerResult.addedHeadersInfo) {
            console.log(`    追加されたヘッダー: ${headerResult.addedHeadersInfo.added}個`);
            headerResult.addedHeadersInfo.addedHeaders?.forEach(header => {
              console.log(`      ${header.fieldName}: ${header.headerName} (${header.columnLetter}列)`);
            });
          }
          
          // テスト後の状態を記録
          const afterState = {
            lastRow: sheet.getLastRow(),
            lastColumn: sheet.getLastColumn()
          };
          
          console.log('  📋 テスト後の状態:');
          console.log(`    最終行: ${afterState.lastRow}, 最終列: ${afterState.lastColumn}`);
          
          testResults.push({
            sheetName: sheetName,
            exists: true,
            success: true,
            beforeState: beforeState,
            afterState: afterState,
            headerResult: headerResult,
            headersAdded: headerResult.headersAdded,
            addedCount: headerResult.addedHeadersInfo?.added || 0
          });
          
        } else {
          console.log(`  ❌ ヘッダー検出・追加失敗: ${headerResult.error}`);
          testResults.push({
            sheetName: sheetName,
            exists: true,
            success: false,
            error: headerResult.error,
            beforeState: beforeState
          });
        }
        
      } catch (sheetError) {
        console.log(`❌ シートテストエラー: ${sheetError.toString()}`);
        testResults.push({
          sheetName: sheetName,
          exists: false,
          error: sheetError.toString()
        });
      }
    }
    
    // 結果サマリー
    console.log('\n=== 自動ヘッダー追加テスト結果サマリー ===');
    const successCount = testResults.filter(r => r.success).length;
    const headersAddedCount = testResults.filter(r => r.headersAdded).length;
    const totalAddedHeaders = testResults.reduce((sum, r) => sum + (r.addedCount || 0), 0);
    
    console.log(`📊 テスト成功: ${successCount}/${testResults.length}シート`);
    console.log(`📊 ヘッダー追加発生: ${headersAddedCount}/${testResults.length}シート`);
    console.log(`📊 総追加ヘッダー数: ${totalAddedHeaders}個`);
    
    console.log('\n📋 シート別結果:');
    testResults.forEach(result => {
      const status = result.success ? '✅' : '❌';
      const addedInfo = result.headersAdded ? `(+${result.addedCount}列)` : '';
      console.log(`${status} ${result.sheetName}${addedInfo}`);
      if (result.error) {
        console.log(`    エラー: ${result.error}`);
      }
    });
    
    return {
      success: successCount === testResults.length,
      message: `自動ヘッダー追加テスト完了: ${successCount}/${testResults.length}シート成功`,
      results: testResults,
      summary: {
        total: testResults.length,
        successful: successCount,
        headersAddedSheets: headersAddedCount,
        totalAddedHeaders: totalAddedHeaders
      }
    };
    
  } catch (error) {
    console.error('自動ヘッダー追加テストエラー:', error.toString());
    return {
      success: false,
      message: '自動ヘッダー追加テストでエラーが発生しました',
      error: error.toString()
    };
  }
}

/**
 * スプレッドシートの動的ヘッダー検知テスト関数
 */
function testSpreadsheetHeaderPosition() {
  console.log('=== スプレッドシート動的ヘッダー検知テスト開始 ===');
  
  try {
    // スプレッドシート設定の確認
    const spreadsheetSettings = getSpreadsheetSettings();
    
    if (!spreadsheetSettings.success || !spreadsheetSettings.settings.spreadsheetId) {
      console.warn('⚠️ スプレッドシートIDが設定されていません。');
      return {
        success: false,
        message: 'スプレッドシートIDが設定されていません',
        error: 'SPREADSHEET_ID_DESTINATION が設定されていません'
      };
    }
    
    const spreadsheetId = spreadsheetSettings.settings.spreadsheetId;
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    console.log('スプレッドシート名:', spreadsheet.getName());
    console.log('検索対象列マッピング:', SPREADSHEET_COLUMN_MAPPING);
    
    // 各拠点シートの動的ヘッダー検知を確認
    const headerCheckResults = [];
    const uniqueSheetNames = [...new Set(Object.values(LOCATION_SHEET_NAMES))];
    
    uniqueSheetNames.forEach(sheetName => {
      console.log(`\n📊 シート「${sheetName}」の動的ヘッダー検知:`);
      
      try {
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
          console.log('❌ シートが見つかりません');
          headerCheckResults.push({
            sheetName: sheetName,
            exists: false,
            error: 'シートが見つかりません'
          });
          return;
        }
        
        // 動的ヘッダー検知を実行
        const headerInfo = detectHeaderRow(sheet);
        
        if (headerInfo.found) {
          console.log(`  ✅ ヘッダー行検出成功: ${headerInfo.rowIndex}行目`);
          console.log(`  🎯 検出された列: ${Object.keys(headerInfo.columnIndexes).length}個`);
          
          // 検出された列の詳細を表示
          Object.keys(headerInfo.columnIndexes).forEach(fieldName => {
            const columnIndex = headerInfo.columnIndexes[fieldName];
            const headerValue = headerInfo.rowData[columnIndex];
            console.log(`    ${fieldName}: 列${columnIndex + 1} (${headerValue})`);
          });
          
          headerCheckResults.push({
            sheetName: sheetName,
            exists: true,
            headerDetected: true,
            headerRowIndex: headerInfo.rowIndex,
            detectedColumns: Object.keys(headerInfo.columnIndexes),
            columnMapping: headerInfo.columnIndexes,
            score: headerInfo.score
          });
        } else {
          console.log(`  ❌ ヘッダー行検出失敗: ${headerInfo.error}`);
          headerCheckResults.push({
            sheetName: sheetName,
            exists: true,
            headerDetected: false,
            error: headerInfo.error
          });
        }
        
      } catch (sheetError) {
        console.log(`❌ シート確認エラー: ${sheetError.toString()}`);
        headerCheckResults.push({
          sheetName: sheetName,
          exists: false,
          error: sheetError.toString()
        });
      }
    });
    
    // 結果サマリー
    console.log('\n=== 動的ヘッダー検知結果サマリー ===');
    headerCheckResults.forEach(result => {
      if (result.exists) {
        if (result.headerDetected) {
          console.log(`✅ ${result.sheetName}: ${result.headerRowIndex}行目（${result.detectedColumns.length}列検出）`);
        } else {
          console.log(`⚠️ ${result.sheetName}: ヘッダー未検出 - ${result.error}`);
        }
      } else {
        console.log(`❌ ${result.sheetName}: ${result.error}`);
      }
    });
    
    return {
      success: true,
      message: '動的ヘッダー検知テスト完了',
      spreadsheetInfo: {
        id: spreadsheetId,
        name: spreadsheet.getName(),
        url: spreadsheet.getUrl()
      },
      columnMapping: SPREADSHEET_COLUMN_MAPPING,
      headerCheckResults: headerCheckResults
    };
    
  } catch (error) {
    console.error('ヘッダー行確認テストエラー:', error.toString());
    return {
      success: false,
      message: 'ヘッダー行確認テストでエラーが発生しました',
      error: error.toString()
    };
  }
}


  