/**
 * QRコード生成のテスト関数
 */
function testQRCodeGeneration() {
  console.log("===== QRコード生成テスト開始 =====");
  
  // テスト用のURL
  const testUrl = "https://example.com/test?id=12345";
  
  // QRコード画像生成をテスト
  const result = generateQRCodeImage(testUrl);
  
  console.log("生成結果:", {
    success: result.success,
    hasImageData: !!result.imageData,
    hasImageUrl: !!result.imageUrl,
    provider: result.provider,
    method: result.method,
    error: result.error || result.imageError
  });
  
  if (result.imageUrl) {
    console.log("画像URL:", result.imageUrl);
    
    // URLが実際にアクセス可能かテスト
    try {
      const testFetch = UrlFetchApp.fetch(result.imageUrl, {
        method: 'HEAD',
        muteHttpExceptions: true
      });
      console.log("画像URLアクセステスト:", {
        statusCode: testFetch.getResponseCode(),
        contentType: testFetch.getContentType()
      });
    } catch (e) {
      console.log("画像URLアクセステストエラー:", e.toString());
    }
  }
  
  return result;
}

/**
 * 統合テスト - 拠点管理番号からQRコード生成まで
 */
function testFullQRCodeFlow() {
  console.log("===== 統合QRコードフローテスト開始 =====");
  
  // テスト用データ
  const testLocationNumber = "HIMEJI_desktop_TestModel_12345_001";
  const testDeviceCategory = "desktop";
  
  // QRコード付きURL生成
  const result = generateQRCodeWithImage(testLocationNumber, testDeviceCategory);
  
  console.log("統合テスト結果:", {
    success: result.success,
    url: result.url,
    hasImageUrl: !!result.imageUrl,
    hasImageData: !!result.imageData,
    provider: result.provider,
    locationNumber: result.locationNumber,
    deviceCategory: result.deviceCategory
  });
  
  return result;
}

/**
 * 診断実行のメイン関数
 */
function runQRDiagnostics() {
  console.log("========== QRコード診断開始 ==========");
  console.log("実行時刻:", new Date().toISOString());
  
  // 1. 基本的なQRコード生成テスト
  console.log("\n----- 基本QRコード生成テスト -----");
  const basicTest = testQRCodeGeneration();
  
  // 2. 統合フローテスト
  console.log("\n----- 統合フローテスト -----");
  const flowTest = testFullQRCodeFlow();
  
  // 3. 診断結果サマリー
  console.log("\n===== 診断結果サマリー =====");
  console.log("基本テスト:", basicTest.success ? "✅ 成功" : "❌ 失敗");
  console.log("統合テスト:", flowTest.success ? "✅ 成功" : "❌ 失敗");
  
  if (basicTest.success && basicTest.imageUrl) {
    console.log("\n✅ QRコード生成は正常に動作しています");
    console.log("使用方法: ブラウザで直接画像URLを使用してください");
    console.log("例: <img src=\"" + basicTest.imageUrl + "\" />");
  } else {
    console.log("\n❌ QRコード生成に問題があります");
    console.log("エラー詳細:", basicTest.error || basicTest.imageError);
  }
  
  return {
    timestamp: new Date().toISOString(),
    basicTest: basicTest,
    flowTest: flowTest,
    summary: {
      allTestsPassed: basicTest.success && flowTest.success
    }
  };
}