/**
 * 列操作ヘルパー関数
 * Google Sheetsの列名を動的に取得するためのユーティリティ関数群
 */

/**
 * 列インデックスを動的に取得するヘルパー関数
 * @param {Array} headers - ヘッダー行の配列
 * @param {string} columnName - 検索する列名
 * @return {number} 列インデックス（見つからない場合は-1）
 */
function getColumnIndex(headers, columnName) {
  return headers.indexOf(columnName);
}

/**
 * 配列から値を安全に取得するヘルパー関数
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー行の配列
 * @param {string} columnName - 取得する列名
 * @return {*} 値（見つからない場合は空文字）
 */
function getValueByColumnName(row, headers, columnName) {
  const index = getColumnIndex(headers, columnName);
  return index >= 0 ? (row[index] || '') : '';
}

/**
 * 列番号を動的に取得するヘルパー関数（1ベース）
 * @param {Array} headers - ヘッダー行の配列
 * @param {string} columnName - 検索する列名
 * @return {number} 列番号（見つからない場合は0）
 */
function getColumnNumber(headers, columnName) {
  const index = getColumnIndex(headers, columnName);
  return index >= 0 ? index + 1 : 0;
}

/**
 * 複数の列名から最初に見つかった列のインデックスを取得
 * @param {Array} headers - ヘッダー行の配列
 * @param {Array<string>} columnNames - 検索する列名の配列（優先順）
 * @return {number} 列インデックス（見つからない場合は-1）
 */
function getColumnIndexMultiple(headers, columnNames) {
  for (const columnName of columnNames) {
    const index = getColumnIndex(headers, columnName);
    if (index >= 0) return index;
  }
  return -1;
}

/**
 * オブジェクトからヘッダーに対応する値の配列を作成
 * @param {Object} data - データオブジェクト
 * @param {Array} headers - ヘッダー行の配列
 * @return {Array} ヘッダーに対応する値の配列
 */
function createRowFromObject(data, headers) {
  return headers.map(header => data[header] || '');
}

/**
 * 行データをオブジェクトに変換
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー行の配列
 * @return {Object} ヘッダーをキーとしたオブジェクト
 */
function rowToObject(row, headers) {
  const obj = {};
  headers.forEach((header, index) => {
    obj[header] = row[index] || '';
  });
  return obj;
}