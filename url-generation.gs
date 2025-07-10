// ========================================
// 共通フォームURL生成機能
// url-generation.gs
// ========================================

/**
 * 共通フォームURLに拠点管理番号を追加してURLを生成
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} deviceCategory - デバイスカテゴリ（SV/CL/プリンタ/その他）
 * @param {boolean} generateQrUrl - QRコード用URL（中間ページ）を生成するかどうか
 * @return {Object} 生成結果
 */
function generateCommonFormUrl(
  locationNumber,
  deviceCategory,
  generateQrUrl = false
) {
  const startTime = startPerformanceTimer();
  addLog("URL生成開始", { locationNumber, deviceCategory, generateQrUrl });

  try {
    // 共通フォームURL設定を取得
    const settings = getCommonFormsSettings();

    // QRコード用URL（中間ページ）を生成する場合
    if (generateQrUrl && settings.qrRedirectUrl) {
      const qrUrl = `${settings.qrRedirectUrl}?id=${encodeURIComponent(
        locationNumber
      )}`;

      endPerformanceTimer(startTime, "QRコード用URL生成");
      addLog("QRコード用URL生成完了", { locationNumber, qrUrl });

      return {
        success: true,
        url: qrUrl,
        baseUrl: settings.qrRedirectUrl,
        locationNumber: locationNumber,
        deviceCategory: deviceCategory,
        isQrUrl: true,
      };
    }

    // 通常の共通フォームURL生成
    let baseUrl;
    let formType;

    // カテゴリ判定（英語・日本語両方に対応）
    if (
      deviceCategory === "desktop" ||
      deviceCategory === "laptop" ||
      deviceCategory === "server" ||
      deviceCategory === "デスクトップPC" ||
      deviceCategory === "サーバー" ||
      deviceCategory === "ノートPC" ||
      deviceCategory === "SV" ||
      deviceCategory === "CL" ||
      deviceCategory === "端末"
    ) {
      baseUrl = settings.terminalCommonFormUrl;
      formType = "端末";
      if (!baseUrl) {
        throw new Error("端末用共通フォームURLが設定されていません");
      }
    } else if (deviceCategory === "printer" || deviceCategory === "プリンタ") {
      baseUrl = settings.printerCommonFormUrl;
      formType = "プリンタ";
      if (!baseUrl) {
        throw new Error(
          "プリンタ・その他用共通フォームURLが設定されていません"
        );
      }
    } else {
      // その他のカテゴリはプリンタ・その他用として扱う
      baseUrl = settings.printerCommonFormUrl;
      formType = "その他";
      if (!baseUrl) {
        throw new Error(
          "プリンタ・その他用共通フォームURLが設定されていません"
        );
      }
    }

    // URLパラメータとして拠点管理番号を追加（フォームタイプに応じたentry IDを使用）
    const generatedUrl = addLocationNumberParameter(
      baseUrl,
      locationNumber,
      formType,
      settings
    );

    endPerformanceTimer(startTime, "URL生成");
    addLog("URL生成完了", { locationNumber, deviceCategory, generatedUrl });

    return {
      success: true,
      url: generatedUrl,
      baseUrl: baseUrl,
      locationNumber: locationNumber,
      deviceCategory: deviceCategory,
      isQrUrl: false,
    };
  } catch (error) {
    endPerformanceTimer(startTime, "URL生成エラー");
    addLog("URL生成エラー", {
      locationNumber,
      deviceCategory,
      error: error.toString(),
    });

    return {
      success: false,
      error: error.toString(),
      locationNumber: locationNumber,
      deviceCategory: deviceCategory,
    };
  }
}

/**
 * 共通フォームURLに拠点管理番号をパラメータとして追加
 * @param {string} baseUrl - 基本URL
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} formType - フォームタイプ（端末/プリンタ/その他）
 * @param {Object} settings - 設定情報（entry IDを含む）
 * @return {string} パラメータ付きURL
 */
function addLocationNumberParameter(
  baseUrl,
  locationNumber,
  formType,
  settings
) {
  try {
    // URLオブジェクトを作成
    const url = new URL(baseUrl);

    // 既存のパラメータがある場合は追加、ない場合は新規作成
    if (url.search) {
      // 既にパラメータがある場合
      return `${baseUrl}&entry.123456789=${encodeURIComponent(locationNumber)}`;
    } else {
      // パラメータがない場合
      const separator = baseUrl.includes("?") ? "&" : "?";
      return `${baseUrl}${separator}entry.123456789=${encodeURIComponent(
        locationNumber
      )}`;
    }
  } catch (error) {
    // URLパースエラーの場合は単純に追加
    const separator = baseUrl.includes("?") ? "&" : "?";
    return `${baseUrl}${separator}entry.123456789=${encodeURIComponent(
      locationNumber
    )}`;
  }
}

/**
 * 端末マスタに共通フォームURLを保存（自動登録機能付き）
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} generatedUrl - 生成されたURL
 * @param {Object} deviceInfo - デバイス情報（自動登録用）
 * @return {Object} 保存結果
 */
function saveUrlToTerminalMaster(
  locationNumber,
  generatedUrl,
  deviceInfo = null
) {
  const startTime = startPerformanceTimer();
  addLog("端末マスタURL保存開始", { locationNumber, generatedUrl, deviceInfo });

  try {
    // 端末マスタシートを取得
    const spreadsheet = SpreadsheetApp.openById(getMainSpreadsheetId());
    const sheet = spreadsheet.getSheetByName("端末マスタ");

    if (!sheet) {
      throw new Error("端末マスタシートが見つかりません");
    }

    // データ範囲を取得
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];

    // 必要な列のインデックスを取得
    const locationNumberColIndex = headerRow.indexOf("拠点管理番号");
    if (locationNumberColIndex === -1) {
      throw new Error("端末マスタに拠点管理番号列が見つかりません");
    }

    // 共通フォームURL列を検索または作成
    let urlColIndex = headerRow.indexOf("共通フォームURL");
    if (urlColIndex === -1) {
      urlColIndex = headerRow.length;
      sheet.getRange(1, urlColIndex + 1).setValue("共通フォームURL");
    }

    // QRコードURL列を検索または作成
    let qrUrlColIndex = headerRow.indexOf("QRコードURL");
    if (qrUrlColIndex === -1) {
      qrUrlColIndex = headerRow.length;
      sheet.getRange(1, qrUrlColIndex + 1).setValue("QRコードURL");
    }

    // 該当行を検索
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][locationNumberColIndex] === locationNumber) {
        targetRow = i + 1;
        break;
      }
    }

    // 行が見つからない場合は新規作成
    if (targetRow === -1 && deviceInfo) {
      targetRow = sheet.getLastRow() + 1;

      // 新規行のデータを準備
      const today = Utilities.formatDate(
        new Date(),
        "Asia/Tokyo",
        "yyyy/MM/dd"
      );
      const newRowData = [];

      // 各列に対応するデータを設定
      headerRow.forEach((header, index) => {
        switch (header) {
          case "拠点管理番号":
            newRowData[index] = locationNumber;
            break;
          case "機種名":
            newRowData[index] = deviceInfo.modelName || "";
            break;
          case "メーカー":
            newRowData[index] = deviceInfo.manufacturer || "";
            break;
          case "カテゴリ":
            newRowData[index] = deviceInfo.category || "";
            break;
          case "製造番号":
            newRowData[index] = deviceInfo.serialNumber || "";
            break;
          case "作成日":
            newRowData[index] = today;
            break;
          case "更新日":
            newRowData[index] = today;
            break;
          case "共通フォームURL":
            newRowData[index] = generatedUrl;
            break;
          default:
            newRowData[index] = "";
        }
      });

      // 新規行を追加
      sheet.getRange(targetRow, 1, 1, headerRow.length).setValues([newRowData]);
      addLog("新規端末データを追加", { locationNumber, targetRow });
    } else if (targetRow === -1) {
      throw new Error(
        `拠点管理番号 ${locationNumber} が端末マスタに見つからず、デバイス情報も提供されていません`
      );
    } else {
      // 既存行の場合はURLのみ更新
      sheet.getRange(targetRow, urlColIndex + 1).setValue(generatedUrl);

      // 更新日も更新
      const updateDateColIndex = headerRow.indexOf("更新日");
      if (updateDateColIndex !== -1) {
        const today = Utilities.formatDate(
          new Date(),
          "Asia/Tokyo",
          "yyyy/MM/dd"
        );
        sheet.getRange(targetRow, updateDateColIndex + 1).setValue(today);
      }
    }

    endPerformanceTimer(startTime, "端末マスタURL保存");
    addLog("端末マスタURL保存完了", { locationNumber, targetRow, urlColIndex });

    return {
      success: true,
      savedRow: targetRow,
      savedColumn: urlColIndex + 1,
      isNewEntry: targetRow > data.length,
    };
  } catch (error) {
    endPerformanceTimer(startTime, "端末マスタURL保存エラー");
    addLog("端末マスタURL保存エラー", {
      locationNumber,
      error: error.toString(),
    });

    return {
      success: false,
      error: error.toString(),
    };
  }
}

/**
 * プリンタマスタに共通フォームURLを保存（自動登録機能付き）
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} generatedUrl - 生成されたURL
 * @param {Object} deviceInfo - デバイス情報（自動登録用）
 * @return {Object} 保存結果
 */
function saveUrlToPrinterMaster(
  locationNumber,
  generatedUrl,
  deviceInfo = null
) {
  const startTime = startPerformanceTimer();
  addLog("プリンタマスタURL保存開始", {
    locationNumber,
    generatedUrl,
    deviceInfo,
  });

  try {
    // プリンタマスタシートを取得
    const spreadsheet = SpreadsheetApp.openById(getMainSpreadsheetId());
    const sheet = spreadsheet.getSheetByName("プリンタマスタ");

    if (!sheet) {
      throw new Error("プリンタマスタシートが見つかりません");
    }

    // データ範囲を取得
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];

    // 必要な列のインデックスを取得
    const locationNumberColIndex = headerRow.indexOf("拠点管理番号");
    if (locationNumberColIndex === -1) {
      throw new Error("プリンタマスタに拠点管理番号列が見つかりません");
    }

    // 共通フォームURL列を検索または作成
    let urlColIndex = headerRow.indexOf("共通フォームURL");
    if (urlColIndex === -1) {
      urlColIndex = headerRow.length;
      sheet.getRange(1, urlColIndex + 1).setValue("共通フォームURL");
    }

    // 該当行を検索
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][locationNumberColIndex] === locationNumber) {
        targetRow = i + 1;
        break;
      }
    }

    // 行が見つからない場合は新規作成
    if (targetRow === -1 && deviceInfo) {
      targetRow = sheet.getLastRow() + 1;

      // 新規行のデータを準備
      const today = Utilities.formatDate(
        new Date(),
        "Asia/Tokyo",
        "yyyy/MM/dd"
      );
      const newRowData = [];

      // 各列に対応するデータを設定
      headerRow.forEach((header, index) => {
        switch (header) {
          case "拠点管理番号":
            newRowData[index] = locationNumber;
            break;
          case "機種名":
            newRowData[index] = deviceInfo.modelName || "";
            break;
          case "メーカー":
            newRowData[index] = deviceInfo.manufacturer || "";
            break;
          case "カテゴリ":
            newRowData[index] = deviceInfo.category || "";
            break;
          case "製造番号":
            newRowData[index] = deviceInfo.serialNumber || "";
            break;
          case "作成日":
            newRowData[index] = today;
            break;
          case "更新日":
            newRowData[index] = today;
            break;
          case "共通フォームURL":
            newRowData[index] = generatedUrl;
            break;
          default:
            newRowData[index] = "";
        }
      });

      // 新規行を追加
      sheet.getRange(targetRow, 1, 1, headerRow.length).setValues([newRowData]);
      addLog("新規プリンタデータを追加", { locationNumber, targetRow });
    } else if (targetRow === -1) {
      throw new Error(
        `拠点管理番号 ${locationNumber} がプリンタマスタに見つからず、デバイス情報も提供されていません`
      );
    } else {
      // 既存行の場合はURLのみ更新
      sheet.getRange(targetRow, urlColIndex + 1).setValue(generatedUrl);

      // 更新日も更新
      const updateDateColIndex = headerRow.indexOf("更新日");
      if (updateDateColIndex !== -1) {
        const today = Utilities.formatDate(
          new Date(),
          "Asia/Tokyo",
          "yyyy/MM/dd"
        );
        sheet.getRange(targetRow, updateDateColIndex + 1).setValue(today);
      }
    }

    endPerformanceTimer(startTime, "プリンタマスタURL保存");
    addLog("プリンタマスタURL保存完了", {
      locationNumber,
      targetRow,
      urlColIndex,
    });

    return {
      success: true,
      savedRow: targetRow,
      savedColumn: urlColIndex + 1,
      isNewEntry: targetRow > data.length,
    };
  } catch (error) {
    endPerformanceTimer(startTime, "プリンタマスタURL保存エラー");
    addLog("プリンタマスタURL保存エラー", {
      locationNumber,
      error: error.toString(),
    });

    return {
      success: false,
      error: error.toString(),
    };
  }
}

/**
 * その他マスタに共通フォームURLを保存（自動登録機能付き）
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} generatedUrl - 生成されたURL
 * @param {Object} deviceInfo - デバイス情報（自動登録用）
 * @return {Object} 保存結果
 */
function saveUrlToOtherMaster(locationNumber, generatedUrl, deviceInfo = null) {
  const startTime = startPerformanceTimer();
  addLog("その他マスタURL保存開始", { locationNumber, generatedUrl });

  try {
    // その他マスタシートを取得
    const spreadsheet = SpreadsheetApp.openById(getMainSpreadsheetId());
    const sheet = spreadsheet.getSheetByName("その他マスタ");

    if (!sheet) {
      throw new Error("その他マスタシートが見つかりません");
    }

    // 拠点管理番号で該当行を検索
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];

    // 拠点管理番号の列インデックスを取得
    const locationNumberColIndex = headerRow.indexOf("拠点管理番号");
    if (locationNumberColIndex === -1) {
      throw new Error("その他マスタに拠点管理番号列が見つかりません");
    }

    // 共通フォームURL列を検索または作成
    let urlColIndex = headerRow.indexOf("共通フォームURL");
    if (urlColIndex === -1) {
      // 列が存在しない場合は追加
      urlColIndex = headerRow.length;
      sheet.getRange(1, urlColIndex + 1).setValue("共通フォームURL");
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
      throw new Error(
        `拠点管理番号 ${locationNumber} がその他マスタに見つかりません`
      );
    }

    // URLを保存
    sheet.getRange(targetRow, urlColIndex + 1).setValue(generatedUrl);

    endPerformanceTimer(startTime, "その他マスタURL保存");
    addLog("その他マスタURL保存完了", {
      locationNumber,
      targetRow,
      urlColIndex,
    });

    return {
      success: true,
      savedRow: targetRow,
      savedColumn: urlColIndex + 1,
    };
  } catch (error) {
    endPerformanceTimer(startTime, "その他マスタURL保存エラー");
    addLog("その他マスタURL保存エラー", {
      locationNumber,
      error: error.toString(),
    });

    return {
      success: false,
      error: error.toString(),
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
  addLog("URL生成・保存一括処理開始", requestData);

  try {
    const {
      locationNumber,
      deviceCategory,
      generateQrUrl = false,
      deviceInfo = null, // デバイス情報を追加
    } = requestData;

    // バリデーション
    if (!locationNumber || !deviceCategory) {
      throw new Error("拠点管理番号とデバイスカテゴリは必須です");
    }

    // URL生成（QRコード用URLまたは通常のフォームURL）
    const urlResult = generateCommonFormUrl(
      locationNumber,
      deviceCategory,
      generateQrUrl
    );
    if (!urlResult.success) {
      throw new Error("URL生成に失敗しました: " + urlResult.error);
    }

    // QRコード用URLの場合はQRコードURLとしてマスタに保存
    if (generateQrUrl && urlResult.isQrUrl) {
      // QRコード用URLをマスタデータに保存（QRコードURL列）
      let saveResult = saveQrUrlToMaster(
        locationNumber,
        deviceCategory,
        urlResult.url,
        deviceInfo
      );

      if (!saveResult.success) {
        throw new Error(
          "QRコード用URLの保存に失敗しました: " + saveResult.error
        );
      }

      return {
        success: true,
        locationNumber: locationNumber,
        deviceCategory: deviceCategory,
        qrUrl: urlResult.url,
        baseUrl: urlResult.baseUrl,
        savedTo: saveResult.masterType,
        savedRow: saveResult.savedRow,
        isQrUrl: true,
        isNewEntry: saveResult.isNewEntry || false,
      };
    }

    // 通常のフォームURLの場合
    let saveResult;
    let masterType;

    // カテゴリに応じて保存先を決定（英語・日本語両方に対応）
    if (
      deviceCategory === "desktop" ||
      deviceCategory === "laptop" ||
      deviceCategory === "server" ||
      deviceCategory === "デスクトップPC" ||
      deviceCategory === "サーバー" ||
      deviceCategory === "ノートPC" ||
      deviceCategory === "SV" ||
      deviceCategory === "CL" ||
      deviceCategory === "端末"
    ) {
      saveResult = saveUrlToTerminalMaster(
        locationNumber,
        urlResult.url,
        deviceInfo
      );
      masterType = "端末マスタ";
    } else if (deviceCategory === "printer" || deviceCategory === "プリンタ") {
      saveResult = saveUrlToPrinterMaster(
        locationNumber,
        urlResult.url,
        deviceInfo
      );
      masterType = "プリンタマスタ";
    } else {
      // その他のカテゴリはその他マスタに保存
      saveResult = saveUrlToOtherMaster(
        locationNumber,
        urlResult.url,
        deviceInfo
      );
      masterType = "その他マスタ";
    }

    if (!saveResult.success) {
      throw new Error(
        "マスタデータへの保存に失敗しました: " + saveResult.error
      );
    }

    endPerformanceTimer(startTime, "URL生成・保存一括処理");
    addLog("URL生成・保存一括処理完了", {
      locationNumber,
      deviceCategory,
      generatedUrl: urlResult.url,
      savedRow: saveResult.savedRow,
      isNewEntry: saveResult.isNewEntry || false,
    });

    return {
      success: true,
      locationNumber: locationNumber,
      deviceCategory: deviceCategory,
      generatedUrl: urlResult.url,
      baseUrl: urlResult.baseUrl,
      savedTo: masterType,
      savedRow: saveResult.savedRow,
      savedColumn: saveResult.savedColumn,
      isQrUrl: false,
      isNewEntry: saveResult.isNewEntry || false,
    };
  } catch (error) {
    endPerformanceTimer(startTime, "URL生成・保存一括処理エラー");
    addLog("URL生成・保存一括処理エラー", {
      requestData,
      error: error.toString(),
    });

    return {
      success: false,
      error: error.toString(),
      requestData: requestData,
    };
  }
}

/**
 * QRコード用URLをマスタデータに保存
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} deviceCategory - デバイスカテゴリ
 * @param {string} qrUrl - QRコード用URL
 * @param {Object} deviceInfo - デバイス情報（自動登録用）
 * @return {Object} 保存結果
 */
function saveQrUrlToMaster(
  locationNumber,
  deviceCategory,
  qrUrl,
  deviceInfo = null
) {
  const startTime = startPerformanceTimer();
  addLog("QRコード用URL保存開始", { locationNumber, deviceCategory, qrUrl });

  try {
    let masterType;
    let saveResult;

    // カテゴリに応じて保存先を決定（英語・日本語両方に対応）
    if (
      deviceCategory === "desktop" ||
      deviceCategory === "laptop" ||
      deviceCategory === "server" ||
      deviceCategory === "デスクトップPC" ||
      deviceCategory === "サーバー" ||
      deviceCategory === "ノートPC" ||
      deviceCategory === "SV" ||
      deviceCategory === "CL" ||
      deviceCategory === "端末"
    ) {
      saveResult = saveQrUrlToTerminalMaster(locationNumber, qrUrl, deviceInfo);
      masterType = "端末マスタ";
    } else if (deviceCategory === "printer" || deviceCategory === "プリンタ") {
      saveResult = saveQrUrlToPrinterMaster(locationNumber, qrUrl, deviceInfo);
      masterType = "プリンタマスタ";
    } else {
      saveResult = saveQrUrlToOtherMaster(locationNumber, qrUrl, deviceInfo);
      masterType = "その他マスタ";
    }

    if (!saveResult.success) {
      throw new Error("QRコード用URLの保存に失敗しました: " + saveResult.error);
    }

    endPerformanceTimer(startTime, "QRコード用URL保存");
    addLog("QRコード用URL保存完了", { locationNumber, masterType });

    return {
      success: true,
      masterType: masterType,
      savedRow: saveResult.savedRow,
      savedColumn: saveResult.savedColumn,
    };
  } catch (error) {
    endPerformanceTimer(startTime, "QRコード用URL保存エラー");
    addLog("QRコード用URL保存エラー", {
      locationNumber,
      error: error.toString(),
    });

    return {
      success: false,
      error: error.toString(),
    };
  }
}

/**
 * 端末マスタにQRコード用URLを保存
 */
function saveQrUrlToTerminalMaster(locationNumber, qrUrl) {
  return saveQrUrlToMasterSheet("端末マスタ", locationNumber, qrUrl);
}

/**
 * プリンタマスタにQRコード用URLを保存
 */
function saveQrUrlToPrinterMaster(locationNumber, qrUrl) {
  return saveQrUrlToMasterSheet("プリンタマスタ", locationNumber, qrUrl);
}

/**
 * その他マスタにQRコード用URLを保存
 */
function saveQrUrlToOtherMaster(locationNumber, qrUrl) {
  return saveQrUrlToMasterSheet("その他マスタ", locationNumber, qrUrl);
}

/**
 * 指定されたマスタシートにQRコード用URLを保存
 */
function saveQrUrlToMasterSheet(sheetName, locationNumber, qrUrl) {
  try {
    const spreadsheet = SpreadsheetApp.openById(getMainSpreadsheetId());
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`${sheetName}シートが見つかりません`);
    }

    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];

    // 拠点管理番号の列インデックスを取得
    const locationNumberColIndex = headerRow.indexOf("拠点管理番号");
    if (locationNumberColIndex === -1) {
      throw new Error(`${sheetName}に拠点管理番号列が見つかりません`);
    }

    // QRコードURL列を検索または作成
    let qrUrlColIndex = headerRow.indexOf("QRコードURL");
    if (qrUrlColIndex === -1) {
      qrUrlColIndex = headerRow.length;
      sheet.getRange(1, qrUrlColIndex + 1).setValue("QRコードURL");
    }

    // 該当行を検索
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][locationNumberColIndex] === locationNumber) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error(
        `拠点管理番号 ${locationNumber} が${sheetName}に見つかりません`
      );
    }

    // QRコード用URLを保存
    sheet.getRange(targetRow, qrUrlColIndex + 1).setValue(qrUrl);

    return {
      success: true,
      savedRow: targetRow,
      savedColumn: qrUrlColIndex + 1,
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
    };
  }
}

/**
 * メインスプレッドシートIDを取得
 * @return {string} スプレッドシートID
 */
function getMainSpreadsheetId() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty("SPREADSHEET_ID_DESTINATION");

  if (!spreadsheetId) {
    throw new Error("SPREADSHEET_ID_DESTINATIONが設定されていません");
  }

  return spreadsheetId;
}
