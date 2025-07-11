/**
 * システム設定と定数管理
 * 全体で使用される定数、設定値を一元管理
 */

// ========================================
// システム設定
// ========================================

// デバッグモード（本番環境では必ずfalseに設定）
const DEBUG = false;

// システムバージョン
const SYSTEM_VERSION = '2.0.0';

// タイムゾーン
const TIMEZONE = 'Asia/Tokyo';

// ========================================
// スプレッドシート関連
// ========================================

// マスタシート名
const MASTER_SHEET_NAMES = {
  terminal: '端末マスタ',
  printer: 'プリンタマスタ',
  other: 'その他マスタ',
  model: '機種マスタ',
  location: '拠点マスタ',
  dataType: 'データタイプマスタ',
  settings: '設定'
};

// 収集シート名
const COLLECTION_SHEET_NAMES = {
  terminalStatus: '端末ステータス収集',
  printerStatus: 'プリンタステータス収集',
  otherStatus: 'その他ステータス収集'
};

// その他のシート名
const OTHER_SHEET_NAMES = {
  log: 'ログ',
  audit: '監査データ',
  summary: 'サマリーデータ'
};

// ========================================
// データ形式と制限
// ========================================

// 日付形式
const DATE_FORMAT = 'yyyy/MM/dd';
const DATETIME_FORMAT = 'yyyy/MM/dd HH:mm:ss';

// データ制限
const DATA_LIMITS = {
  maxRowsPerFetch: 1000,  // 一度に取得する最大行数
  maxCacheAge: 300,       // キャッシュの有効期限（秒）
  maxLogEntries: 10000,   // ログエントリの最大数
  maxEmailsPerHour: 10    // 1時間あたりの最大メール送信数
};

// ========================================
// カテゴリ定義
// ========================================

// デバイスカテゴリ
const DEVICE_CATEGORIES = {
  desktop: {
    id: 'desktop',
    displayName: 'デスクトップ',
    aliases: ['デスクトップPC', 'CL'],
    formType: 'terminal'
  },
  laptop: {
    id: 'laptop',
    displayName: 'ノートパソコン',
    aliases: ['ノートPC'],
    formType: 'terminal'
  },
  server: {
    id: 'server',
    displayName: 'サーバー',
    aliases: ['SV'],
    formType: 'terminal'
  },
  printer: {
    id: 'printer',
    displayName: 'プリンタ',
    aliases: [],
    formType: 'printer'
  },
  other: {
    id: 'other',
    displayName: 'その他',
    aliases: [],
    formType: 'other'
  }
};

// カテゴリマッピング（表示名 → ID）
const CATEGORY_DISPLAY_MAPPING = Object.values(DEVICE_CATEGORIES).reduce((map, cat) => {
  map[cat.displayName] = cat.id;
  cat.aliases.forEach(alias => {
    map[alias] = cat.id;
  });
  return map;
}, {});

// ========================================
// ステータス定義
// ========================================

// 機器ステータス
const DEVICE_STATUS = {
  LENDING: '1.貸出中',
  RETURNING: '2.返却済み',
  STORAGE: '3.社内にて保管中',
  MAINTENANCE: '4.メンテナンス中',
  DISPOSAL: '5.廃棄済み'
};

// 預り機ステータス
const CUSTODY_STATUS = {
  WAITING_RETURN: '返却待ち',
  RETURNED: '返却済み',
  DISPOSAL_PLANNED: '廃棄予定',
  OTHER: 'その他'
};

// アクティブステータス
const ACTIVE_STATUS = {
  ACTIVE: 'active',
  INACTIVE: 'inactive'
};

// ========================================
// バリデーションパターン
// ========================================

const VALIDATION_PATTERNS = {
  locationId: /^[a-z0-9_-]+$/,
  locationCode: /^[A-Z0-9]+$/,
  email: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
  locationNumber: /^[A-Za-z0-9_-]+$/,
  phoneNumber: /^[\d-]+$/,
  postalCode: /^\d{3}-?\d{4}$/
};

// ========================================
// UI設定
// ========================================

const UI_SETTINGS = {
  defaultPageSize: 20,
  maxPageSize: 100,
  toastDuration: 3000,
  errorToastDuration: 5000,
  loadingDelay: 300,
  debounceDelay: 500
};

// ========================================
// API設定
// ========================================

const API_SETTINGS = {
  timeout: 30000,  // API タイムアウト（ミリ秒）
  retryCount: 3,   // リトライ回数
  retryDelay: 1000 // リトライ間隔（ミリ秒）
};

// ========================================
// メール設定
// ========================================

const EMAIL_SETTINGS = {
  fromName: '代替機管理システム',
  replyTo: 'noreply@example.com',
  errorNotificationSubject: '[代替機管理システム] エラー通知',
  statusChangeSubject: '[代替機管理システム] ステータス変更通知'
};

// ========================================
// ヘルパー関数
// ========================================

/**
 * カテゴリIDを取得
 * @param {string} categoryName - カテゴリ名（表示名またはエイリアス）
 * @returns {string|null} カテゴリID
 */
function getCategoryId(categoryName) {
  return CATEGORY_DISPLAY_MAPPING[categoryName] || null;
}

/**
 * カテゴリ表示名を取得
 * @param {string} categoryId - カテゴリID
 * @returns {string|null} カテゴリ表示名
 */
function getCategoryDisplayName(categoryId) {
  const category = DEVICE_CATEGORIES[categoryId];
  return category ? category.displayName : null;
}

/**
 * デバッグログを出力
 * @param {string} message - ログメッセージ
 * @param {*} data - ログデータ
 */
function debugLog(message, data = null) {
  if (DEBUG) {
    console.log(`[DEBUG] ${message}`, data);
  }
}

/**
 * パフォーマンス計測開始
 * @param {string} label - 計測ラベル
 * @returns {number} 開始時刻
 */
function startPerformanceTimer(label) {
  if (DEBUG) {
    console.time(label);
  }
  return new Date().getTime();
}

/**
 * パフォーマンス計測終了
 * @param {string} label - 計測ラベル
 * @param {number} startTime - 開始時刻
 */
function endPerformanceTimer(label, startTime) {
  if (DEBUG) {
    console.timeEnd(label);
    const duration = new Date().getTime() - startTime;
    console.log(`${label}: ${duration}ms`);
  }
}

// ========================================
// 設定値の検証
// ========================================

/**
 * 設定値を検証
 * @throws {Error} 設定値が不正な場合
 */
function validateConfig() {
  // 必須のシート名が定義されているか確認
  const requiredSheets = ['terminal', 'printer', 'other', 'model', 'location'];
  requiredSheets.forEach(key => {
    if (!MASTER_SHEET_NAMES[key]) {
      throw new Error(`必須シート名が未定義: ${key}`);
    }
  });
  
  // デバイスカテゴリの整合性確認
  Object.keys(DEVICE_CATEGORIES).forEach(key => {
    const category = DEVICE_CATEGORIES[key];
    if (!category.id || !category.displayName || !category.formType) {
      throw new Error(`カテゴリ定義が不完全: ${key}`);
    }
  });
}

// 起動時に設定を検証
try {
  validateConfig();
} catch (error) {
  console.error('設定エラー:', error);
}