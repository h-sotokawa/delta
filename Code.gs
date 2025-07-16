// ========================================
// Code.gs - メインエントリーポイント
// ========================================

/**
 * Google Apps Script Webアプリケーションのエントリーポイント
 * この関数はWebアプリケーションへのGETリクエストを処理します
 * 
 * @param {Object} e - イベントオブジェクト
 * @return {HtmlService.HtmlOutput} HTMLレスポンス
 */
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

// ========================================
// 以下の関数は他のモジュールに移動されました
// ========================================

/*
移動先一覧：

1. location-master-operations.gs
   - 拠点マスタ関連の全機能
   - getLocationMaster, addLocation, updateLocation, deleteLocation等

2. model-master-operations.gs
   - 機種マスタ関連の全機能
   - generateNextModelId, getModelMaster, addModelMasterData等

3. spreadsheet-operations.gs
   - スプレッドシート操作関連
   - getSpreadsheetData, updateMachineStatus, checkDataConsistency等

4. form-storage-operations.gs
   - フォーム・ストレージ管理
   - getFormStorageSettings, saveFormStorageSettings, validateDriveFolderId等

5. utility-functions.gs
   - ユーティリティ関数
   - addLog, startPerformanceTimer, formatDateFast, include等

6. test-diagnostic-functions.gs
   - テスト・診断関数
   - testJurisdictionFeatures, testDynamicSummaryDisplay, diagnoseSummarySheet等
*/