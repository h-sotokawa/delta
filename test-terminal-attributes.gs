/**
 * 端末属性機能のテスト関数
 */
function testTerminalAttributes() {
  console.log("===== 端末属性機能テスト開始 =====");
  
  // テストデータ
  const testData = {
    locationNumber: "OSAKA_server_ThinkPad_ABC123_001",
    deviceCategory: "server",
    deviceInfo: {
      modelName: "ThinkPad X1 Carbon",
      manufacturer: "Lenovo",
      category: "server",
      serialNumber: "ABC123",
      sequenceNumber: "1",
      software: "Office 2021",      // 新規追加
      os: "Windows 11",            // 新規追加
      assetNumber: "IT-2024-0001"  // 新規追加
    }
  };
  
  console.log("テストデータ:", testData);
  
  // マスタシートの確認
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("端末マスタ");
  if (!sheet) {
    console.error("端末マスタシートが見つかりません");
    return;
  }
  
  // ヘッダー行の確認
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  console.log("現在のヘッダー:", headers);
  
  // 必要な列のインデックスを確認
  const softwareIndex = headers.indexOf("ソフトウェア");
  const osIndex = headers.indexOf("OS");
  const assetIndex = headers.indexOf("資産管理番号");
  
  console.log("列インデックス:", {
    software: softwareIndex,
    os: osIndex,
    assetNumber: assetIndex
  });
  
  // 端末属性が正しく保存されるかテスト
  console.log("\n===== 保存テスト =====");
  try {
    const result = generateAndSaveCommonFormUrl(testData);
    console.log("保存結果:", result);
    
    if (result.success) {
      // 保存されたデータを確認
      const savedRow = result.savedRow;
      const savedData = sheet.getRange(savedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      console.log("\n保存されたデータ:");
      console.log("ソフトウェア:", savedData[softwareIndex]);
      console.log("OS:", savedData[osIndex]);
      console.log("資産管理番号:", savedData[assetIndex]);
      
      // 検証
      const verification = {
        software: savedData[softwareIndex] === testData.deviceInfo.software,
        os: savedData[osIndex] === testData.deviceInfo.os,
        assetNumber: savedData[assetIndex] === testData.deviceInfo.assetNumber
      };
      
      console.log("\n検証結果:", verification);
      
      if (verification.software && verification.os && verification.assetNumber) {
        console.log("✅ 端末属性が正しく保存されました");
      } else {
        console.log("❌ 端末属性の保存に問題があります");
      }
    }
  } catch (error) {
    console.error("テスト中にエラーが発生しました:", error);
  }
  
  return {
    timestamp: new Date().toISOString(),
    testPassed: true
  };
}

/**
 * カテゴリ別の表示制御テスト
 */
function testCategoryDisplay() {
  console.log("===== カテゴリ別表示制御テスト =====");
  
  const categories = ["server", "desktop", "laptop", "printer", "other"];
  
  categories.forEach(category => {
    console.log(`\nカテゴリ: ${category}`);
    
    // 端末系カテゴリかどうか判定
    const isTerminal = ["server", "desktop", "laptop"].includes(category);
    
    console.log(`端末系カテゴリ: ${isTerminal ? "はい" : "いいえ"}`);
    console.log(`端末属性セクション: ${isTerminal ? "表示" : "非表示"}`);
    console.log(`必須項目:`);
    if (isTerminal) {
      console.log("  - ソフトウェア: 必須");
      console.log("  - OS: 必須");
      console.log("  - 資産管理番号: 必須");
    } else {
      console.log("  - 端末属性なし");
    }
  });
}