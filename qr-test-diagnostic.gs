/**
 * QRコード生成テスト診断機能
 * Google Apps Scriptのネットワーク環境でQRコード生成APIの接続性をテストする
 */

/**
 * APIの接続性をテストする汎用関数
 */
function testApiConnectivity(apiName, apiUrl, options = {}) {
  console.log(`\n===== ${apiName} テスト開始 =====`);
  console.log(`URL: ${apiUrl}`);
  
  try {
    const defaultOptions = {
      method: 'GET',
      muteHttpExceptions: true,
      validateHttpsCertificates: false,
      followRedirects: true
    };
    
    const requestOptions = Object.assign({}, defaultOptions, options);
    console.log('リクエストオプション:', JSON.stringify(requestOptions));
    
    const response = UrlFetchApp.fetch(apiUrl, requestOptions);
    const responseCode = response.getResponseCode();
    const contentType = response.getContentType();
    const headers = response.getAllHeaders();
    
    console.log(`ステータスコード: ${responseCode}`);
    console.log(`Content-Type: ${contentType}`);
    console.log('レスポンスヘッダー:', JSON.stringify(headers));
    
    if (responseCode === 200) {
      const blob = response.getBlob();
      const bytes = blob.getBytes();
      console.log(`✅ 成功！ データサイズ: ${bytes.length} bytes`);
      
      if (bytes.length > 0) {
        // 最初の数バイトをチェックしてPNG画像か確認
        const pngHeader = [0x89, 0x50, 0x4E, 0x47]; // PNG signature
        const isPng = pngHeader.every((byte, index) => bytes[index] === byte);
        console.log(`PNG画像チェック: ${isPng ? '✅ 有効なPNG' : '❌ PNG形式ではない'}`);
      }
      
      return { success: true, size: bytes.length };
    } else {
      console.log(`❌ エラー: HTTP ${responseCode}`);
      const errorContent = response.getContentText();
      console.log('エラー内容:', errorContent.substring(0, 200));
      return { success: false, error: `HTTP ${responseCode}` };
    }
    
  } catch (error) {
    console.log(`❌ 例外エラー: ${error.toString()}`);
    console.log('エラー詳細:', error.message);
    return { success: false, error: error.toString() };
  }
}

/**
 * すべてのQRコードAPIをテストする診断関数
 */
function runQRCodeDiagnostics() {
  console.log('QRコード生成API診断を開始します...');
  console.log('実行時刻:', new Date().toISOString());
  console.log('Google Apps Script環境からの接続テスト');
  
  const testUrl = 'https://www.google.com';
  const results = [];
  
  // 1. QR Server (api.qrserver.com) - 最もシンプルな形式
  results.push(testApiConnectivity(
    'QR Server (最小パラメータ)',
    `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(testUrl)}`
  ));
  
  // 2. QR Server with size
  results.push(testApiConnectivity(
    'QR Server (サイズ指定)',
    `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(testUrl)}&size=200x200`
  ));
  
  // 3. QuickChart.io
  results.push(testApiConnectivity(
    'QuickChart.io',
    `https://quickchart.io/qr?text=${encodeURIComponent(testUrl)}`
  ));
  
  // 4. Google Charts (deprecated but might work)
  results.push(testApiConnectivity(
    'Google Charts API',
    `https://chart.googleapis.com/chart?chs=200x200&cht=qr&chl=${encodeURIComponent(testUrl)}`
  ));
  
  // 5. 基本的なHTTPS接続テスト
  console.log('\n===== 基本的なHTTPS接続テスト =====');
  try {
    const testResponse = UrlFetchApp.fetch('https://www.google.com');
    console.log('✅ Google.comへの接続: 成功');
  } catch (e) {
    console.log('❌ Google.comへの接続: 失敗', e.toString());
  }
  
  // 6. Proxyやファイアウォールの確認
  console.log('\n===== プロキシ/ファイアウォール設定の確認 =====');
  const scriptProperties = PropertiesService.getScriptProperties();
  const userProperties = PropertiesService.getUserProperties();
  
  console.log('スクリプトプロパティ:', scriptProperties.getKeys());
  console.log('ユーザープロパティ:', userProperties.getKeys());
  
  // 結果のサマリー
  console.log('\n===== 診断結果サマリー =====');
  const successCount = results.filter(r => r && r.success).length;
  console.log(`テスト済みAPI: ${results.length}`);
  console.log(`成功: ${successCount}`);
  console.log(`失敗: ${results.length - successCount}`);
  
  return {
    timestamp: new Date().toISOString(),
    results: results,
    summary: {
      total: results.length,
      success: successCount,
      failed: results.length - successCount
    }
  };
}

/**
 * 内部ネットワークからのQRコード生成テスト
 */
function testInternalQRGeneration() {
  console.log('\n===== 内部QRコード生成テスト =====');
  
  // GASの内部機能を使ってQRコードを生成する代替案
  const testText = 'https://www.example.com';
  
  // 1. Google SlidesのQRコード機能を試す
  try {
    console.log('Google Slides APIを使用したQRコード生成を試行...');
    // この方法は実際にはGoogle Apps Scriptでは直接利用できませんが、
    // 将来的な拡張のためのプレースホルダーとして記載
    console.log('❌ Google Slides API: 現在のGAS環境では利用不可');
  } catch (e) {
    console.log('Google Slides エラー:', e.toString());
  }
  
  // 2. HTMLサービスを使用したクライアントサイド生成の可能性を確認
  console.log('\n代替案: HTMLサービスでクライアントサイドQRコード生成');
  console.log('この方法では、クライアント側のJavaScriptライブラリを使用します');
  console.log('例: qrcode.js, QRCode.js など');
  
  return {
    recommendation: 'クライアントサイドでのQRコード生成を推奨',
    alternatives: [
      'HTMLサービスでqrcode.jsを使用',
      'Google ChartsのQRコードAPIをiframeで埋め込み',
      '外部サービスへのプロキシ経由でのアクセス'
    ]
  };
}

/**
 * 完全な診断を実行
 */
function runFullDiagnostics() {
  console.log('========== QRコード生成 完全診断 ==========');
  
  // API接続性テスト
  const apiResults = runQRCodeDiagnostics();
  
  // 内部生成テスト
  const internalResults = testInternalQRGeneration();
  
  // 診断結果をスプレッドシートに記録
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('診断ログ');
    if (sheet) {
      sheet.appendRow([
        new Date(),
        'QR診断',
        JSON.stringify(apiResults),
        JSON.stringify(internalResults)
      ]);
    }
  } catch (e) {
    console.log('診断結果の記録に失敗:', e.toString());
  }
  
  return {
    apiDiagnostics: apiResults,
    internalAlternatives: internalResults,
    timestamp: new Date().toISOString()
  };
}