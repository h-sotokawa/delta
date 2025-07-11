/**
 * エラーハンドリングユーティリティ
 * 統一されたエラー処理とロギング機能を提供
 */

// ========================================
// エラータイプの定義
// ========================================
const ErrorTypes = {
  VALIDATION: 'VALIDATION_ERROR',
  DATA_ACCESS: 'DATA_ACCESS_ERROR',
  PERMISSION: 'PERMISSION_ERROR',
  CONFIGURATION: 'CONFIGURATION_ERROR',
  NETWORK: 'NETWORK_ERROR',
  BUSINESS_LOGIC: 'BUSINESS_LOGIC_ERROR',
  UNKNOWN: 'UNKNOWN_ERROR'
};

// ========================================
// カスタムエラークラス
// ========================================
class AppError extends Error {
  constructor(message, type = ErrorTypes.UNKNOWN, details = null) {
    super(message);
    this.name = 'AppError';
    this.type = type;
    this.details = details;
    this.timestamp = new Date().toISOString();
  }
  
  toJSON() {
    return {
      name: this.name,
      message: this.message,
      type: this.type,
      details: this.details,
      timestamp: this.timestamp,
      stack: this.stack
    };
  }
}

// ========================================
// エラーハンドリング関数
// ========================================

/**
 * エラーをログに記録
 * @param {Error|AppError} error - エラーオブジェクト
 * @param {string} context - エラーが発生したコンテキスト
 * @param {Object} additionalInfo - 追加情報
 */
function logError(error, context = '', additionalInfo = {}) {
  const errorLog = {
    timestamp: new Date().toISOString(),
    context: context,
    error: error instanceof AppError ? error.toJSON() : {
      name: error.name || 'Error',
      message: error.message || error.toString(),
      stack: error.stack
    },
    additionalInfo: additionalInfo
  };
  
  // コンソールに出力
  console.error('エラーログ:', errorLog);
  
  // エラー通知（設定されている場合）
  try {
    const settings = getSettings();
    if (settings.errorNotificationEmail) {
      // エラー通知メールを送信（非同期で実行）
      sendErrorNotification(settings.errorNotificationEmail, errorLog);
    }
  } catch (notificationError) {
    console.error('エラー通知の送信に失敗:', notificationError);
  }
  
  return errorLog;
}

/**
 * 関数を安全に実行
 * @param {Function} fn - 実行する関数
 * @param {string} functionName - 関数名（ログ用）
 * @param {*} defaultValue - エラー時のデフォルト値
 * @returns {*} 実行結果またはデフォルト値
 */
function safeExecute(fn, functionName = 'unknown', defaultValue = null) {
  try {
    return fn();
  } catch (error) {
    logError(error, `Function: ${functionName}`);
    return defaultValue;
  }
}

/**
 * 非同期関数を安全に実行
 * @param {Function} asyncFn - 実行する非同期関数
 * @param {string} functionName - 関数名（ログ用）
 * @param {*} defaultValue - エラー時のデフォルト値
 * @returns {Promise<*>} 実行結果またはデフォルト値
 */
async function safeExecuteAsync(asyncFn, functionName = 'unknown', defaultValue = null) {
  try {
    return await asyncFn();
  } catch (error) {
    logError(error, `AsyncFunction: ${functionName}`);
    return defaultValue;
  }
}

/**
 * エラーレスポンスを生成
 * @param {Error|AppError|string} error - エラーオブジェクトまたはメッセージ
 * @param {string} userMessage - ユーザー向けメッセージ
 * @returns {Object} 標準化されたエラーレスポンス
 */
function createErrorResponse(error, userMessage = null) {
  const errorObj = error instanceof Error ? error : new Error(error);
  
  return {
    success: false,
    error: userMessage || errorObj.message,
    errorType: errorObj instanceof AppError ? errorObj.type : ErrorTypes.UNKNOWN,
    timestamp: new Date().toISOString(),
    // デバッグモードの場合は詳細情報を含める
    ...(DEBUG && {
      details: errorObj instanceof AppError ? errorObj.details : null,
      stack: errorObj.stack
    })
  };
}

/**
 * エラー通知メールを送信
 * @param {string} email - 送信先メールアドレス
 * @param {Object} errorLog - エラーログ
 */
function sendErrorNotification(email, errorLog) {
  try {
    // 送信間隔の制限（同じエラーの連続送信を防ぐ）
    const lastNotificationKey = 'lastErrorNotification_' + errorLog.error.message;
    const lastNotification = PropertiesService.getScriptProperties().getProperty(lastNotificationKey);
    const now = new Date().getTime();
    
    if (lastNotification) {
      const lastTime = parseInt(lastNotification);
      const timeSinceLastNotification = now - lastTime;
      const minInterval = 60 * 60 * 1000; // 1時間
      
      if (timeSinceLastNotification < minInterval) {
        console.log('エラー通知の送信をスキップ（送信間隔制限）');
        return;
      }
    }
    
    // メール本文を作成
    const subject = `[代替機管理システム] エラー通知: ${errorLog.context}`;
    const body = `
以下のエラーが発生しました。

発生日時: ${errorLog.timestamp}
コンテキスト: ${errorLog.context}

エラー詳細:
- 名前: ${errorLog.error.name}
- メッセージ: ${errorLog.error.message}
- タイプ: ${errorLog.error.type || 'N/A'}

追加情報:
${JSON.stringify(errorLog.additionalInfo, null, 2)}

スタックトレース:
${errorLog.error.stack || 'N/A'}

このメールは自動送信されています。
`;
    
    // メール送信
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body
    });
    
    // 最終送信時刻を記録
    PropertiesService.getScriptProperties().setProperty(lastNotificationKey, now.toString());
    
    console.log('エラー通知メールを送信しました:', email);
  } catch (error) {
    console.error('エラー通知メールの送信に失敗:', error);
  }
}

// ========================================
// バリデーションヘルパー
// ========================================

/**
 * 必須フィールドのバリデーション
 * @param {Object} data - チェック対象のデータ
 * @param {Array<string>} requiredFields - 必須フィールド名の配列
 * @throws {AppError} バリデーションエラー
 */
function validateRequiredFields(data, requiredFields) {
  const missingFields = [];
  
  requiredFields.forEach(field => {
    if (!data[field] || (typeof data[field] === 'string' && data[field].trim() === '')) {
      missingFields.push(field);
    }
  });
  
  if (missingFields.length > 0) {
    throw new AppError(
      `必須フィールドが不足しています: ${missingFields.join(', ')}`,
      ErrorTypes.VALIDATION,
      { missingFields }
    );
  }
}

/**
 * メールアドレスのバリデーション
 * @param {string} email - メールアドレス
 * @returns {boolean} 有効な場合true
 */
function validateEmail(email) {
  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
}

/**
 * 日付形式のバリデーション
 * @param {string} dateString - 日付文字列
 * @param {string} format - 期待される形式（例: 'yyyy/MM/dd'）
 * @returns {boolean} 有効な場合true
 */
function validateDateFormat(dateString, format = 'yyyy/MM/dd') {
  if (format === 'yyyy/MM/dd') {
    const pattern = /^\d{4}\/\d{2}\/\d{2}$/;
    if (!pattern.test(dateString)) return false;
    
    const parts = dateString.split('/');
    const year = parseInt(parts[0]);
    const month = parseInt(parts[1]);
    const day = parseInt(parts[2]);
    
    const date = new Date(year, month - 1, day);
    return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day;
  }
  
  return true; // 他の形式は現時点では常にtrue
}

// ========================================
// エクスポート（Google Apps Scriptではグローバルスコープ）
// ========================================
// 関数は自動的にグローバルスコープで利用可能