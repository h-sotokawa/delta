// 拠点名の日本語表示用マッピング
const LOCATION_NAMES = {
  'osaka_desktop': '大阪(デスクトップ)',
  'osaka_notebook': '大阪(ノート、サーバー)',
  'kobe': '神戸',
  'himeji': '姫路'
};

const TARGET_SHEET_NAME = 'main';

// ログを保持する配列
let serverLogs = [];

// ログを追加する関数
function addLog(message, data = null) {
  const timestamp = new Date().toISOString();
  const log = {
    timestamp: timestamp,
    message: message,
    data: data
  };
  serverLogs.push(log);
  console.log(`[${timestamp}] ${message}`, data);
  return log;
}

function doGet(e) {
  serverLogs = []; // ログをリセット
  addLog('doGet関数が呼び出されました', e);
  addLog('現在のユーザー', Session.getActiveUser().getEmail());
  addLog('実行環境', Session.getEffectiveUser().getEmail());
  
  try {
    addLog('テンプレート作成開始');
    const template = HtmlService.createTemplateFromFile('Index');
    addLog('テンプレート作成成功');
    
    addLog('HTML評価開始');
    const html = template.evaluate()
        .setTitle('スプレッドシートビューアー')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setFaviconUrl('https://www.google.com/favicon.ico');
    
    addLog('HTMLの評価成功', {
      title: html.getTitle(),
      content: html.getContent().substring(0, 100) + '...' // 最初の100文字のみ表示
    });
    
    return html;
  } catch (error) {
    addLog('doGetでエラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack,
      name: error.name
    });
    
    return HtmlService.createHtmlOutput(`
      <html>
        <body>
          <h1>エラーが発生しました</h1>
          <pre>${error.toString()}</pre>
          <h2>デバッグ情報</h2>
          <pre>${JSON.stringify(serverLogs, null, 2)}</pre>
        </body>
      </html>
    `);
  }
}

function include(filename) {
  addLog('include関数が呼び出されました', filename);
  try {
    const content = HtmlService.createHtmlOutputFromFile(filename).getContent();
    addLog('ファイルの読み込み成功', filename);
    return content;
  } catch (error) {
    addLog('includeでエラーが発生', error);
    throw error;
  }
}

// 拠点ごとのスプレッドシートIDをスクリプトプロパティから取得
function getSpreadsheetIdFromProperty(location) {
  const propertyKeys = {
    'osaka_desktop': 'SPREADSHEET_ID_SOURCE_OSAKA_DESKTOP',
    'osaka_notebook': 'SPREADSHEET_ID_SOURCE_OSAKA_LAPTOP',
    'kobe': 'SPREADSHEET_ID_SOURCE_KOBE',
    'himeji': 'SPREADSHEET_ID_SOURCE_HIMEJI',
  };
  const key = propertyKeys[location];
  if (!key) return null;
  const id = PropertiesService.getScriptProperties().getProperty(key);
  return id || null;
}

function getSpreadsheetData(location, queryType) {
  addLog('getSpreadsheetData関数が呼び出されました', { location, queryType });
  
  try {
    // 選択された拠点のスプレッドシートIDをスクリプトプロパティから取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにスプレッドシートIDが設定されていません: ' + location);
    }

    addLog('使用するスプレッドシートID', spreadsheetId);
    addLog('対象シート名', TARGET_SHEET_NAME);
    
    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
    
    if (!sheet) {
      throw new Error('シート「' + TARGET_SHEET_NAME + '」が見つかりません。');
    }

    // データを取得
    const range = sheet.getDataRange();
    const data = range.getValues();
    
    // 日付データの処理
    for (let i = 1; i < data.length; i++) {
      for (let j = 0; j < data[i].length; j++) {
        const cell = data[i][j];
        if (cell instanceof Date) {
          // 日付オブジェクトを文字列に変換
          const year = cell.getFullYear();
          const month = String(cell.getMonth() + 1).padStart(2, '0');
          const day = String(cell.getDate()).padStart(2, '0');
          const hours = String(cell.getHours()).padStart(2, '0');
          const minutes = String(cell.getMinutes()).padStart(2, '0');
          const seconds = String(cell.getSeconds()).padStart(2, '0');
          data[i][j] = `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
        }
      }
    }
    
    const response = {
      success: true,
      data: data,
      logs: serverLogs,
      metadata: {
        location: LOCATION_NAMES[location],
        queryType: queryType,
        spreadsheetName: spreadsheet.getName(),
        sheetName: sheet.getName(),
        lastRow: sheet.getLastRow(),
        lastColumn: sheet.getLastColumn()
      }
    };
    
    addLog('返却するレスポンス', response);
    return response;
    
  } catch (error) {
    addLog('エラーが発生', {
      error: error.toString(),
      message: error.message,
      stack: error.stack,
      name: error.name
    });
    
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack,
        name: error.name
      },
      logs: serverLogs
    };
  }
}

// 拠点一覧を取得する関数
function getLocations() {
  return {
    success: true,
    locations: LOCATION_NAMES
  };
}

// ステータス更新用の関数
function updateMachineStatus(rowIndex, newStatus, location) {
  try {
    // location引数で指定された拠点のスプレッドシートIDをスクリプトプロパティから取得
    const spreadsheetId = getSpreadsheetIdFromProperty(location);
    addLog('updateMachineStatus: ステータスを変更するスプレッドシートID', spreadsheetId);
    if (!spreadsheetId) {
      throw new Error('スクリプトプロパティにスプレッドシートIDが設定されていません: ' + location);
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    
    // 行インデックスを1から始まる形式に変換（ヘッダー行を考慮）
    const actualRow = rowIndex + 1;
    
    // 1列目（A列）の値を更新
    sheet.getRange(actualRow, 1).setValue(newStatus);

    // 15列目（O列）にユーザー、16列目（P列）に変更日時を記録
    const user = Session.getActiveUser().getEmail();
    const now = new Date();
    sheet.getRange(actualRow, 15).setValue(user);
    sheet.getRange(actualRow, 16).setValue(now);
    
    return {
      success: true,
      message: 'ステータスを更新しました',
      data: {
        rowIndex: rowIndex,
        newStatus: newStatus,
        user: user,
        updatedAt: now,
        spreadsheetId: spreadsheetId
      }
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      errorDetails: {
        message: error.message,
        stack: error.stack
      }
    };
  }
} 