/**
 * 日時処理ユーティリティ
 * システム全体で統一された日時処理を提供
 */

// ========================================
// 日時フォーマット関数
// ========================================

/**
 * 現在の日付を取得（yyyy/MM/dd形式）
 * @returns {string} フォーマットされた日付文字列
 */
function getCurrentDate() {
  return Utilities.formatDate(new Date(), TIMEZONE, DATE_FORMAT);
}

/**
 * 現在の日時を取得（yyyy/MM/dd HH:mm:ss形式）
 * @returns {string} フォーマットされた日時文字列
 */
function getCurrentDateTime() {
  return Utilities.formatDate(new Date(), TIMEZONE, DATETIME_FORMAT);
}

/**
 * 日付をフォーマット
 * @param {Date|string} dateValue - 日付値
 * @param {string} format - フォーマット文字列（省略時はDATE_FORMAT）
 * @returns {string} フォーマットされた日付文字列
 */
function formatDate(dateValue, format = DATE_FORMAT) {
  if (!dateValue) return '';
  
  try {
    // 既に文字列の場合はそのまま返す
    if (typeof dateValue === 'string') {
      return dateValue;
    }
    
    // Dateオブジェクトの場合
    if (dateValue instanceof Date) {
      return Utilities.formatDate(dateValue, TIMEZONE, format);
    }
    
    // その他の場合は文字列に変換
    return dateValue.toString();
  } catch (error) {
    logError(error, 'formatDate', { dateValue, format });
    return '';
  }
}

/**
 * 文字列を日付オブジェクトに変換
 * @param {string} dateString - 日付文字列
 * @returns {Date|null} 日付オブジェクト
 */
function parseDate(dateString) {
  if (!dateString) return null;
  
  try {
    // yyyy/MM/dd形式の場合
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(dateString)) {
      const parts = dateString.split('/');
      const year = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1; // 月は0-based
      const day = parseInt(parts[2]);
      return new Date(year, month, day);
    }
    
    // yyyy-MM-dd形式の場合
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
      const parts = dateString.split('-');
      const year = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1;
      const day = parseInt(parts[2]);
      return new Date(year, month, day);
    }
    
    // その他の形式は標準のパーサーに任せる
    const date = new Date(dateString);
    return isNaN(date.getTime()) ? null : date;
  } catch (error) {
    logError(error, 'parseDate', { dateString });
    return null;
  }
}

/**
 * 日付の差分を計算（日数）
 * @param {Date|string} date1 - 日付1
 * @param {Date|string} date2 - 日付2
 * @returns {number} 日数差
 */
function getDaysDifference(date1, date2) {
  try {
    const d1 = typeof date1 === 'string' ? parseDate(date1) : date1;
    const d2 = typeof date2 === 'string' ? parseDate(date2) : date2;
    
    if (!d1 || !d2) return 0;
    
    const diffTime = Math.abs(d2 - d1);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  } catch (error) {
    logError(error, 'getDaysDifference', { date1, date2 });
    return 0;
  }
}

/**
 * 日付を加算
 * @param {Date|string} date - 基準日
 * @param {number} days - 加算する日数
 * @returns {Date} 新しい日付
 */
function addDays(date, days) {
  try {
    const d = typeof date === 'string' ? parseDate(date) : new Date(date);
    if (!d) return new Date();
    
    d.setDate(d.getDate() + days);
    return d;
  } catch (error) {
    logError(error, 'addDays', { date, days });
    return new Date();
  }
}

/**
 * 月初日を取得
 * @param {Date|string} date - 基準日（省略時は現在）
 * @returns {Date} 月初日
 */
function getFirstDayOfMonth(date = new Date()) {
  try {
    const d = typeof date === 'string' ? parseDate(date) : new Date(date);
    if (!d) return new Date();
    
    return new Date(d.getFullYear(), d.getMonth(), 1);
  } catch (error) {
    logError(error, 'getFirstDayOfMonth', { date });
    return new Date();
  }
}

/**
 * 月末日を取得
 * @param {Date|string} date - 基準日（省略時は現在）
 * @returns {Date} 月末日
 */
function getLastDayOfMonth(date = new Date()) {
  try {
    const d = typeof date === 'string' ? parseDate(date) : new Date(date);
    if (!d) return new Date();
    
    return new Date(d.getFullYear(), d.getMonth() + 1, 0);
  } catch (error) {
    logError(error, 'getLastDayOfMonth', { date });
    return new Date();
  }
}

/**
 * 営業日を計算（土日を除く）
 * @param {Date|string} startDate - 開始日
 * @param {number} businessDays - 営業日数
 * @returns {Date} 計算後の日付
 */
function addBusinessDays(startDate, businessDays) {
  try {
    let date = typeof startDate === 'string' ? parseDate(startDate) : new Date(startDate);
    if (!date) return new Date();
    
    let daysAdded = 0;
    const direction = businessDays >= 0 ? 1 : -1;
    businessDays = Math.abs(businessDays);
    
    while (daysAdded < businessDays) {
      date = addDays(date, direction);
      const dayOfWeek = date.getDay();
      
      // 土日でない場合はカウント
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        daysAdded++;
      }
    }
    
    return date;
  } catch (error) {
    logError(error, 'addBusinessDays', { startDate, businessDays });
    return new Date();
  }
}

/**
 * 日付の妥当性をチェック
 * @param {*} value - チェック対象の値
 * @returns {boolean} 有効な日付の場合true
 */
function isValidDate(value) {
  if (!value) return false;
  
  if (typeof value === 'string') {
    // 日付形式の文字列かチェック
    if (!validateDateFormat(value)) return false;
    const parsed = parseDate(value);
    return parsed !== null;
  }
  
  if (value instanceof Date) {
    return !isNaN(value.getTime());
  }
  
  return false;
}

/**
 * 日付範囲の妥当性をチェック
 * @param {Date|string} startDate - 開始日
 * @param {Date|string} endDate - 終了日
 * @returns {boolean} 有効な範囲の場合true
 */
function isValidDateRange(startDate, endDate) {
  if (!isValidDate(startDate) || !isValidDate(endDate)) return false;
  
  const start = typeof startDate === 'string' ? parseDate(startDate) : startDate;
  const end = typeof endDate === 'string' ? parseDate(endDate) : endDate;
  
  return start <= end;
}

// ========================================
// スプレッドシート用ヘルパー
// ========================================

/**
 * セルの日付形式を設定
 * @param {Range} range - 対象範囲
 */
function setDateFormat(range) {
  range.setNumberFormat(DATE_FORMAT.replace(/yyyy/g, 'YYYY').replace(/MM/g, 'MM').replace(/dd/g, 'DD'));
}

/**
 * セルの日時形式を設定
 * @param {Range} range - 対象範囲
 */
function setDateTimeFormat(range) {
  range.setNumberFormat(DATETIME_FORMAT.replace(/yyyy/g, 'YYYY').replace(/MM/g, 'MM').replace(/dd/g, 'DD'));
}

/**
 * セルの値を日付として取得
 * @param {Range} range - 対象セル
 * @returns {string} フォーマットされた日付文字列
 */
function getCellDateValue(range) {
  const value = range.getValue();
  return formatDate(value);
}

/**
 * セルに日付を設定（文字列形式）
 * @param {Range} range - 対象セル
 * @param {Date|string} date - 設定する日付
 */
function setCellDateValue(range, date = new Date()) {
  const dateString = formatDate(date);
  range.setValue(dateString);
  range.setNumberFormat('@'); // テキスト形式に設定
}