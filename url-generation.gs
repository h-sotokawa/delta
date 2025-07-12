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
    if (generateQrUrl) {
      if (settings.qrRedirectUrl) {
        // QRリダイレクトURLが設定されている場合
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
      } else {
        // QRリダイレクトURLが設定されていない場合は、通常のURLを返す
        addLog("QRリダイレクトURLが未設定のため、通常のURLでQRコードを生成", { locationNumber, deviceCategory });
        // generateQrUrlフラグをfalseにして、通常のURL生成処理を継続
        generateQrUrl = false;
      }
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

    // URLパラメータとして拠点管理番号を追加
    const generatedUrl = addLocationNumberParameter(baseUrl, locationNumber);

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
 * @return {string} パラメータ付きURL
 */
function addLocationNumberParameter(baseUrl, locationNumber) {
  try {
    console.log('URL生成:', { baseUrl, locationNumber });
    
    // 既にパラメータを含むURLの場合、拠点管理番号のみを設定
    if (baseUrl.includes('entry.') && baseUrl.includes('=')) {
      // 既存のentry.XXXXパラメータがある場合
      // 例: https://docs.google.com/forms/.../viewform?entry.1372464946=
      // 最後の=の後に拠点管理番号を追加
      if (baseUrl.endsWith('=')) {
        return `${baseUrl}${encodeURIComponent(locationNumber)}`;
      } else {
        // 他のパラメータがある場合は&で区切って追加
        return `${baseUrl}&entry.1372464946=${encodeURIComponent(locationNumber)}`;
      }
    } else {
      // パラメータがない場合は?で追加
      const separator = baseUrl.includes('?') ? '&' : '?';
      return `${baseUrl}${separator}entry.1372464946=${encodeURIComponent(locationNumber)}`;
    }
  } catch (error) {
    console.error('URL生成エラー:', error);
    // エラーの場合は単純に追加
    const separator = baseUrl.includes('?') ? '&' : '?';
    return `${baseUrl}${separator}entry.1372464946=${encodeURIComponent(locationNumber)}`;
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
  deviceInfo = null,
  qrUrl = null
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
          case "資産管理番号":
            newRowData[index] = deviceInfo.assetNumber || "";
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
          case "QRコードURL":
            newRowData[index] = qrUrl || "";
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
      // 既存行の場合はURLを更新
      sheet.getRange(targetRow, urlColIndex + 1).setValue(generatedUrl);
      
      // QRコードURLも更新
      if (qrUrl && qrUrlColIndex !== -1) {
        sheet.getRange(targetRow, qrUrlColIndex + 1).setValue(qrUrl);
      }

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
  deviceInfo = null,
  qrUrl = null
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
          case "資産管理番号":
            newRowData[index] = deviceInfo.assetNumber || "";
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
          case "QRコードURL":
            newRowData[index] = qrUrl || "";
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
      // 既存行の場合はURLを更新
      sheet.getRange(targetRow, urlColIndex + 1).setValue(generatedUrl);
      
      // QRコードURLも更新
      if (qrUrl && qrUrlColIndex !== -1) {
        sheet.getRange(targetRow, qrUrlColIndex + 1).setValue(qrUrl);
      }

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
    
    // QRコード用URLを生成（全カテゴリ共通）
    const settings = getCommonFormsSettings();
    const qrUrl = settings.qrRedirectUrl ? 
      `${settings.qrRedirectUrl}?id=${encodeURIComponent(locationNumber)}` : 
      null;

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
        deviceInfo,
        qrUrl
      );
      masterType = "端末マスタ";
    } else if (deviceCategory === "printer" || deviceCategory === "プリンタ") {
      saveResult = saveUrlToPrinterMaster(
        locationNumber,
        urlResult.url,
        deviceInfo,
        qrUrl
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
      qrUrl: qrUrl, // QRコード用URLを追加
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

/**
 * QRコード画像データを生成（サーバーサイド処理）
 * @param {string} url - QRコードに埋め込むURL
 * @return {Object} 画像データを含む結果オブジェクト
 */
function generateQRCodeImage(url) {
  const startTime = startPerformanceTimer();
  addLog("QRコード画像生成開始", { url });

  try {
    if (!url) {
      throw new Error("URLが指定されていません");
    }

    // URLの長さをチェック（QRコードの制限を考慮）
    if (url.length > 2000) {
      addLog("警告: URLが長すぎます", { 
        urlLength: url.length,
        url: url.substring(0, 100) + "..."
      });
    }

    // 1. 最初にシンプルなGoQR.me APIを試す
    const goQrUrl = 'https://api.qrserver.com/v1/create-qr-code/?' + 
      'size=200x200&' +
      'data=' + encodeURIComponent(url);

    try {
      const response = UrlFetchApp.fetch(goQrUrl, {
        muteHttpExceptions: true,
        validateHttpsCertificates: true,
        headers: {
          'User-Agent': 'Google Apps Script'
        }
      });

      if (response.getResponseCode() === 200) {
        const blob = response.getBlob();
        const base64Data = Utilities.base64Encode(blob.getBytes());
        const contentType = blob.getContentType();

        endPerformanceTimer(startTime, "QRコード画像生成（GoQR.me）");
        addLog("QRコード画像生成成功（GoQR.me）", { 
          url, 
          imageSize: blob.getBytes().length 
        });

        return {
          success: true,
          imageData: 'data:' + contentType + ';base64,' + base64Data,
          provider: 'GoQR.me',
          url: url
        };
      } else {
        addLog("GoQR.me API レスポンスエラー", { 
          statusCode: response.getResponseCode(),
          statusText: response.getContentText().substring(0, 200),
          url: goQrUrl
        });
      }
    } catch (goQrError) {
      addLog("GoQR.me API エラー", { 
        error: goQrError.toString(),
        message: goQrError.message,
        url: goQrUrl
      });
    }

    // 2. フォールバック: QuickChart.io API
    const quickChartUrl = 'https://quickchart.io/qr?' + 
      'text=' + encodeURIComponent(url) + '&' +
      'size=200&' +      // サイズ
      'margin=1&' +      // マージン
      'dark=000000&' +   // 前景色（黒）
      'light=ffffff&' +  // 背景色（白）
      'ecLevel=M';       // エラー訂正レベル（中）

    try {
      const response = UrlFetchApp.fetch(quickChartUrl, {
        muteHttpExceptions: true,
        validateHttpsCertificates: true,
        headers: {
          'User-Agent': 'Google Apps Script'
        }
      });

      if (response.getResponseCode() === 200) {
        const blob = response.getBlob();
        const base64Data = Utilities.base64Encode(blob.getBytes());
        const contentType = blob.getContentType();

        endPerformanceTimer(startTime, "QRコード画像生成（QuickChart）");
        addLog("QRコード画像生成成功（QuickChart）", { 
          url, 
          imageSize: blob.getBytes().length 
        });

        return {
          success: true,
          imageData: 'data:' + contentType + ';base64,' + base64Data,
          provider: 'QuickChart',
          url: url
        };
      } else {
        addLog("QuickChart API レスポンスエラー", { 
          statusCode: response.getResponseCode(),
          statusText: response.getContentText(),
          url: quickChartUrl
        });
      }
    } catch (quickChartError) {
      addLog("QuickChart API エラー", { 
        error: quickChartError.toString(),
        message: quickChartError.message,
        url: quickChartUrl
      });
    }

    // 3. 最終フォールバック: 同じQR Server APIだが異なるパラメータ
    const goQrUrl = 'https://api.qrserver.com/v1/create-qr-code/?' + 
      'data=' + encodeURIComponent(url) + '&' +
      'size=200x200';      // 最小限のパラメータで再試行

    try {
      const response = UrlFetchApp.fetch(goQrUrl, {
        muteHttpExceptions: true,
        validateHttpsCertificates: true,
        headers: {
          'User-Agent': 'Google Apps Script',
          'Accept': 'image/png'
        }
      });

      if (response.getResponseCode() === 200) {
        const blob = response.getBlob();
        const base64Data = Utilities.base64Encode(blob.getBytes());
        const contentType = blob.getContentType() || 'image/png';

        endPerformanceTimer(startTime, "QRコード画像生成（QR Server Alt）");
        addLog("QRコード画像生成成功（QR Server Alt）", { 
          url, 
          imageSize: blob.getBytes().length 
        });

        return {
          success: true,
          imageData: 'data:' + contentType + ';base64,' + base64Data,
          provider: 'QR Server (Alternative)',
          url: url
        };
      } else {
        addLog("QR Server Alt API レスポンスエラー", { 
          statusCode: response.getResponseCode(),
          statusText: response.getContentText(),
          url: goQrUrl
        });
      }
    } catch (goQrError) {
      addLog("QR Server Alt API エラー", { 
        error: goQrError.toString(),
        message: goQrError.message,
        url: goQrUrl
      });
    }

    // すべてのAPIが失敗した場合
    throw new Error("QRコード画像の生成に失敗しました。すべてのAPIがエラーを返しました。");

  } catch (error) {
    endPerformanceTimer(startTime, "QRコード画像生成エラー");
    addLog("QRコード画像生成エラー", {
      url,
      error: error.toString()
    });

    return {
      success: false,
      error: error.toString(),
      url: url
    };
  }
}

/**
 * QRコード用URLを生成し、画像データも含めて返す
 * @param {string} locationNumber - 拠点管理番号
 * @param {string} deviceCategory - デバイスカテゴリ
 * @return {Object} URL生成結果と画像データ
 */
function generateQRCodeWithImage(locationNumber, deviceCategory) {
  const startTime = startPerformanceTimer();
  addLog("QRコード生成（画像付き）開始", { locationNumber, deviceCategory });

  try {
    // まずQRコード用URLを生成
    const urlResult = generateCommonFormUrl(locationNumber, deviceCategory, true);
    
    if (!urlResult.success) {
      throw new Error("URL生成に失敗しました: " + urlResult.error);
    }

    // URLからQRコード画像を生成
    const imageResult = generateQRCodeImage(urlResult.url);
    
    if (!imageResult.success) {
      // 画像生成に失敗してもURLは返す
      return {
        success: true,
        url: urlResult.url,
        locationNumber: locationNumber,
        deviceCategory: deviceCategory,
        imageError: imageResult.error
      };
    }

    endPerformanceTimer(startTime, "QRコード生成（画像付き）");
    addLog("QRコード生成（画像付き）完了", { 
      locationNumber, 
      deviceCategory,
      provider: imageResult.provider 
    });

    // URL生成結果と画像データを統合
    return {
      success: true,
      url: urlResult.url,
      baseUrl: urlResult.baseUrl,
      locationNumber: locationNumber,
      deviceCategory: deviceCategory,
      isQrUrl: urlResult.isQrUrl,
      imageData: imageResult.imageData,
      imageProvider: imageResult.provider
    };

  } catch (error) {
    endPerformanceTimer(startTime, "QRコード生成（画像付き）エラー");
    addLog("QRコード生成（画像付き）エラー", {
      locationNumber,
      deviceCategory,
      error: error.toString()
    });

    return {
      success: false,
      error: error.toString(),
      locationNumber: locationNumber,
      deviceCategory: deviceCategory
    };
  }
}
