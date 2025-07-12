/**
 * URL生成機能の統合テスト
 */
function testUrlGenerationComplete() {
  console.log("===== URL生成機能統合テスト開始 =====");
  
  // テストケース1: 端末系カテゴリ（全属性あり）
  console.log("\n----- テストケース1: 端末系カテゴリ -----");
  const terminalData = {
    locationNumber: "OSAKA_server_ThinkPad_ABC123_001",
    deviceCategory: "server",
    deviceInfo: {
      modelName: "ThinkPad X1 Carbon",
      manufacturer: "Lenovo",
      category: "server",
      serialNumber: "ABC123",
      sequenceNumber: "1",
      software: "Office 2021",      // 端末系で必須
      os: "Windows 11",            // 端末系で必須
      assetNumber: "IT-2024-0001"  // 端末系で必須
    }
  };
  
  try {
    const result1 = generateAndSaveCommonFormUrl(terminalData);
    console.log("端末系結果:", {
      success: result1.success,
      savedTo: result1.savedTo,
      savedRow: result1.savedRow,
      url: result1.generatedUrl ? "生成成功" : "生成失敗"
    });
  } catch (error) {
    console.error("端末系エラー:", error);
  }
  
  // テストケース2: プリンタカテゴリ（属性なし）
  console.log("\n----- テストケース2: プリンタカテゴリ -----");
  const printerData = {
    locationNumber: "OSAKA_printer_MFP_XYZ789_001",
    deviceCategory: "printer",
    deviceInfo: {
      modelName: "LaserJet Pro",
      manufacturer: "HP",
      category: "printer",
      serialNumber: "XYZ789",
      sequenceNumber: "1"
      // software, os, assetNumber は含まない（プリンタでは不要）
    }
  };
  
  try {
    const result2 = generateAndSaveCommonFormUrl(printerData);
    console.log("プリンタ結果:", {
      success: result2.success,
      savedTo: result2.savedTo,
      savedRow: result2.savedRow,
      url: result2.generatedUrl ? "生成成功" : "生成失敗"
    });
  } catch (error) {
    console.error("プリンタエラー:", error);
  }
  
  // データ保存確認
  console.log("\n----- 保存データ確認 -----");
  verifyMasterData();
}

/**
 * マスタデータの保存状態を確認
 */
function verifyMasterData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 端末マスタの確認
  const terminalSheet = spreadsheet.getSheetByName("端末マスタ");
  if (terminalSheet) {
    const headers = terminalSheet.getRange(1, 1, 1, terminalSheet.getLastColumn()).getValues()[0];
    console.log("\n端末マスタのヘッダー:", headers);
    
    // 必要な列が存在するか確認
    const requiredColumns = ["OS", "ソフトウェア", "資産管理番号"];
    requiredColumns.forEach(col => {
      const index = headers.indexOf(col);
      console.log(`${col}列: ${index >= 0 ? "存在（列" + (index + 1) + "）" : "不足"}`);
    });
    
    // 最新のデータを確認
    const lastRow = terminalSheet.getLastRow();
    if (lastRow > 1) {
      const latestData = terminalSheet.getRange(lastRow, 1, 1, terminalSheet.getLastColumn()).getValues()[0];
      console.log("\n最新の端末データ:");
      headers.forEach((header, index) => {
        if (["拠点管理番号", "機種名", "OS", "ソフトウェア", "資産管理番号"].includes(header)) {
          console.log(`${header}: ${latestData[index]}`);
        }
      });
    }
  }
  
  // プリンタマスタの確認
  const printerSheet = spreadsheet.getSheetByName("プリンタマスタ");
  if (printerSheet) {
    const headers = printerSheet.getRange(1, 1, 1, printerSheet.getLastColumn()).getValues()[0];
    console.log("\n\nプリンタマスタのヘッダー:", headers);
    
    // 端末属性列が含まれていないことを確認
    const terminalColumns = ["OS", "ソフトウェア", "資産管理番号"];
    terminalColumns.forEach(col => {
      const index = headers.indexOf(col);
      console.log(`${col}列: ${index >= 0 ? "存在（不要）" : "なし（正常）"}`);
    });
    
    // 最新のデータを確認
    const lastRow = printerSheet.getLastRow();
    if (lastRow > 1) {
      const latestData = printerSheet.getRange(lastRow, 1, 1, printerSheet.getLastColumn()).getValues()[0];
      console.log("\n最新のプリンタデータ:");
      headers.forEach((header, index) => {
        if (["拠点管理番号", "機種名", "製造番号"].includes(header)) {
          console.log(`${header}: ${latestData[index]}`);
        }
      });
    }
  }
}

/**
 * 端末属性フィールドの動作確認
 */
function testTerminalAttributesValidation() {
  console.log("\n===== 端末属性バリデーションテスト =====");
  
  // カテゴリごとの属性要件
  const categories = {
    "server": { requiresAttributes: true, label: "サーバー" },
    "desktop": { requiresAttributes: true, label: "デスクトップ" },
    "laptop": { requiresAttributes: true, label: "ノートPC" },
    "printer": { requiresAttributes: false, label: "プリンタ" },
    "other": { requiresAttributes: false, label: "その他" }
  };
  
  Object.entries(categories).forEach(([category, config]) => {
    console.log(`\n${config.label} (${category}):`);
    console.log(`端末属性必須: ${config.requiresAttributes ? "はい" : "いいえ"}`);
    
    if (config.requiresAttributes) {
      console.log("必須フィールド:");
      console.log("  - ソフトウェア");
      console.log("  - OS");
      console.log("  - 資産管理番号");
    }
  });
}