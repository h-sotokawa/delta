/**
 * 預り機ステータス管理機能
 * 「0-4.ステータス」が"1.貸出中"かつ「1-4.ユーザー機の預り有無」が"有り"の場合に
 * 設定可能な副ステータスを管理
 */

// ========================================
// 定数定義
// ========================================

// 預り機ステータスの選択肢
const CUSTODY_STATUS_OPTIONS = [
  '返却待ち',
  '返却済み',
  '廃棄予定',
  'その他'
];

// 預り機ステータスを設定可能な条件
const CUSTODY_STATUS_CONDITIONS = {
  mainStatus: '1.貸出中',
  custodyField: '有り'
};

// ========================================
// 預り機ステータス管理関数
// ========================================

/**
 * 預り機ステータスの編集可否を判定
 * @param {Object} rowData - 行データ
 * @param {Object} headers - ヘッダー情報
 * @returns {boolean} 編集可能な場合true
 */
function canEditCustodyStatus(rowData, headers) {
  try {
    // 必要な列のインデックスを取得
    const statusIndex = findColumnIndex(headers, '0-4.ステータス');
    const custodyIndex = findColumnIndex(headers, '1-4.ユーザー機の預り有無');
    
    if (statusIndex === -1 || custodyIndex === -1) {
      debugLog('必要な列が見つかりません', { statusIndex, custodyIndex });
      return false;
    }
    
    // 条件をチェック
    const mainStatus = rowData[statusIndex];
    const custodyStatus = rowData[custodyIndex];
    
    return mainStatus === CUSTODY_STATUS_CONDITIONS.mainStatus && 
           custodyStatus === CUSTODY_STATUS_CONDITIONS.custodyField;
  } catch (error) {
    logError(error, 'canEditCustodyStatus');
    return false;
  }
}

/**
 * 預り機ステータスを更新
 * @param {Object} params - 更新パラメータ
 * @returns {Object} 処理結果
 */
function updateCustodyStatus(params) {
  try {
    validateRequiredFields(params, ['sheetName', 'rowIndex', 'custodyStatus']);
    
    const sheet = getSheetByName(params.sheetName);
    if (!sheet) {
      throw new AppError(
        'シートが見つかりません',
        ErrorTypes.DATA_ACCESS,
        { sheetName: params.sheetName }
      );
    }
    
    // ヘッダー行を取得
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const custodyStatusIndex = findColumnIndex(headers, '1-5.預り機ステータス');
    
    if (custodyStatusIndex === -1) {
      throw new AppError(
        '預り機ステータス列が見つかりません',
        ErrorTypes.CONFIGURATION
      );
    }
    
    // 現在の行データを取得して条件を再チェック
    const rowData = sheet.getRange(params.rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!canEditCustodyStatus(rowData, headers)) {
      throw new AppError(
        '預り機ステータスの編集条件を満たしていません',
        ErrorTypes.VALIDATION
      );
    }
    
    // ステータスを更新
    const actualColumn = custodyStatusIndex + 1; // 1-based index
    sheet.getRange(params.rowIndex, actualColumn).setValue(params.custodyStatus);
    
    // 変更履歴を記録
    recordCustodyStatusChange({
      sheetName: params.sheetName,
      rowIndex: params.rowIndex,
      locationNumber: rowData[findColumnIndex(headers, '0-2.拠点管理番号')],
      oldStatus: rowData[custodyStatusIndex],
      newStatus: params.custodyStatus,
      changeReason: params.changeReason || '',
      changedBy: Session.getActiveUser().getEmail()
    });
    
    logError(null, '預り機ステータスを更新しました', params);
    
    return {
      success: true,
      message: '預り機ステータスを更新しました'
    };
  } catch (error) {
    logError(error, 'updateCustodyStatus', params);
    return createErrorResponse(error, '預り機ステータスの更新に失敗しました');
  }
}

/**
 * 一括で預り機ステータスを更新
 * @param {Object} params - 更新パラメータ
 * @returns {Object} 処理結果
 */
function batchUpdateCustodyStatus(params) {
  try {
    validateRequiredFields(params, ['sheetName', 'updates']);
    
    if (!Array.isArray(params.updates) || params.updates.length === 0) {
      throw new AppError(
        '更新データが指定されていません',
        ErrorTypes.VALIDATION
      );
    }
    
    const sheet = getSheetByName(params.sheetName);
    if (!sheet) {
      throw new AppError(
        'シートが見つかりません',
        ErrorTypes.DATA_ACCESS,
        { sheetName: params.sheetName }
      );
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const custodyStatusIndex = findColumnIndex(headers, '1-5.預り機ステータス');
    
    if (custodyStatusIndex === -1) {
      throw new AppError(
        '預り機ステータス列が見つかりません',
        ErrorTypes.CONFIGURATION
      );
    }
    
    const results = [];
    const actualColumn = custodyStatusIndex + 1;
    
    // 各更新を実行
    params.updates.forEach(update => {
      try {
        const rowData = sheet.getRange(update.rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        if (!canEditCustodyStatus(rowData, headers)) {
          results.push({
            rowIndex: update.rowIndex,
            success: false,
            error: '編集条件を満たしていません'
          });
          return;
        }
        
        // ステータスを更新
        sheet.getRange(update.rowIndex, actualColumn).setValue(update.custodyStatus);
        
        // 変更履歴を記録
        recordCustodyStatusChange({
          sheetName: params.sheetName,
          rowIndex: update.rowIndex,
          locationNumber: rowData[findColumnIndex(headers, '0-2.拠点管理番号')],
          oldStatus: rowData[custodyStatusIndex],
          newStatus: update.custodyStatus,
          changeReason: params.changeReason || '',
          changedBy: Session.getActiveUser().getEmail()
        });
        
        results.push({
          rowIndex: update.rowIndex,
          success: true
        });
      } catch (error) {
        results.push({
          rowIndex: update.rowIndex,
          success: false,
          error: error.message
        });
      }
    });
    
    const successCount = results.filter(r => r.success).length;
    const failureCount = results.filter(r => !r.success).length;
    
    return {
      success: true,
      message: `${successCount}件更新成功、${failureCount}件失敗`,
      results
    };
  } catch (error) {
    logError(error, 'batchUpdateCustodyStatus', params);
    return createErrorResponse(error, '一括更新に失敗しました');
  }
}

/**
 * 預り機ステータスの変更履歴を記録
 * @param {Object} changeData - 変更データ
 */
function recordCustodyStatusChange(changeData) {
  try {
    const logData = {
      timestamp: getCurrentDateTime(),
      sheetName: changeData.sheetName,
      rowIndex: changeData.rowIndex,
      locationNumber: changeData.locationNumber,
      oldStatus: changeData.oldStatus || '',
      newStatus: changeData.newStatus,
      changeReason: changeData.changeReason,
      changedBy: changeData.changedBy,
      changeType: 'CUSTODY_STATUS_CHANGE'
    };
    
    // ログシートに記録
    addLog('預り機ステータス変更', logData);
    
    // 通知が必要な場合は送信
    if (shouldNotifyCustodyStatusChange(changeData)) {
      sendCustodyStatusChangeNotification(changeData);
    }
  } catch (error) {
    logError(error, 'recordCustodyStatusChange', changeData);
  }
}

/**
 * 預り機ステータス変更の通知要否を判定
 * @param {Object} changeData - 変更データ
 * @returns {boolean} 通知が必要な場合true
 */
function shouldNotifyCustodyStatusChange(changeData) {
  // 設定に基づいて通知要否を判定
  // TODO: 設定画面で通知条件を設定できるようにする
  return changeData.newStatus === '返却済み' || changeData.newStatus === '廃棄予定';
}

/**
 * 預り機ステータス変更通知を送信
 * @param {Object} changeData - 変更データ
 */
function sendCustodyStatusChangeNotification(changeData) {
  try {
    const settings = getSettings();
    if (!settings.custodyStatusNotificationEmail) {
      debugLog('預り機ステータス変更通知メールアドレスが設定されていません');
      return;
    }
    
    const subject = `[代替機管理] 預り機ステータス変更通知 - ${changeData.locationNumber}`;
    const body = `
預り機のステータスが変更されました。

■ 変更内容
拠点管理番号: ${changeData.locationNumber}
変更前ステータス: ${changeData.oldStatus || '(未設定)'}
変更後ステータス: ${changeData.newStatus}
変更理由: ${changeData.changeReason || '(理由なし)'}
変更者: ${changeData.changedBy}
変更日時: ${getCurrentDateTime()}

■ 対応のお願い
${getActionMessage(changeData.newStatus)}

このメールは自動送信されています。
`;
    
    MailApp.sendEmail({
      to: settings.custodyStatusNotificationEmail,
      subject: subject,
      body: body
    });
    
    debugLog('預り機ステータス変更通知を送信しました', { to: settings.custodyStatusNotificationEmail });
  } catch (error) {
    logError(error, 'sendCustodyStatusChangeNotification', changeData);
  }
}

/**
 * ステータスに応じたアクションメッセージを取得
 * @param {string} status - ステータス
 * @returns {string} アクションメッセージ
 */
function getActionMessage(status) {
  const messages = {
    '返却済み': 'ユーザー機の返却が完了しました。代替機の返却手続きを進めてください。',
    '廃棄予定': '廃棄手続きの準備をお願いします。',
    'その他': '詳細を確認の上、適切な対応をお願いします。'
  };
  
  return messages[status] || '状況を確認の上、適切な対応をお願いします。';
}

/**
 * 預り機ステータスのサマリーを取得
 * @param {string} location - 拠点ID（省略時は全拠点）
 * @returns {Object} サマリーデータ
 */
function getCustodyStatusSummary(location = null) {
  try {
    const summary = {
      total: 0,
      byStatus: {},
      byLocation: {}
    };
    
    // 各マスタシートからデータを収集
    const sheetNames = [MASTER_SHEET_NAMES.terminal, MASTER_SHEET_NAMES.printer, MASTER_SHEET_NAMES.other];
    
    sheetNames.forEach(sheetName => {
      const sheet = getSheetByName(sheetName);
      if (!sheet) return;
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const statusIndex = findColumnIndex(headers, '0-4.ステータス');
      const custodyIndex = findColumnIndex(headers, '1-4.ユーザー機の預り有無');
      const custodyStatusIndex = findColumnIndex(headers, '1-5.預り機ステータス');
      const locationIndex = findColumnIndex(headers, '拠点');
      
      if (statusIndex === -1 || custodyIndex === -1 || custodyStatusIndex === -1) return;
      
      const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      
      data.forEach(row => {
        // 条件を満たすデータのみ集計
        if (row[statusIndex] === CUSTODY_STATUS_CONDITIONS.mainStatus && 
            row[custodyIndex] === CUSTODY_STATUS_CONDITIONS.custodyField) {
          
          // 拠点フィルタ
          if (location && locationIndex !== -1 && row[locationIndex] !== location) {
            return;
          }
          
          summary.total++;
          
          const custodyStatus = row[custodyStatusIndex] || '未設定';
          summary.byStatus[custodyStatus] = (summary.byStatus[custodyStatus] || 0) + 1;
          
          if (locationIndex !== -1) {
            const loc = row[locationIndex];
            if (!summary.byLocation[loc]) {
              summary.byLocation[loc] = {};
            }
            summary.byLocation[loc][custodyStatus] = (summary.byLocation[loc][custodyStatus] || 0) + 1;
          }
        }
      });
    });
    
    return {
      success: true,
      summary
    };
  } catch (error) {
    logError(error, 'getCustodyStatusSummary', { location });
    return createErrorResponse(error, '預り機ステータスサマリーの取得に失敗しました');
  }
}

/**
 * ヘッダー配列から列インデックスを検索
 * @param {Array} headers - ヘッダー配列
 * @param {string} columnName - 列名
 * @returns {number} インデックス（見つからない場合は-1）
 */
function findColumnIndex(headers, columnName) {
  return headers.findIndex(header => header === columnName);
}

/**
 * シート名からシートを取得
 * @param {string} sheetName - シート名
 * @returns {Sheet|null} シートオブジェクト
 */
function getSheetByName(sheetName) {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  } catch (error) {
    logError(error, 'getSheetByName', { sheetName });
    return null;
  }
}