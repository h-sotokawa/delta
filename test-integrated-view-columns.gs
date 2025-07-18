// ========================================
// 統合ビューの列構造テスト
// ========================================

/**
 * 統合ビューの列構造と動的列取得をテストする関数
 */
function testIntegratedViewColumns() {
  console.log('=== 統合ビューの列構造テスト開始 ===');
  
  try {
    // 1. 端末系統合ビューシートのヘッダーを確認
    const terminalSheet = getIntegratedViewTerminalSheet();
    const terminalHeaders = terminalSheet.getRange(1, 1, 1, terminalSheet.getLastColumn()).getValues()[0];
    
    console.log('\n端末系統合ビューのヘッダー数:', terminalHeaders.length);
    console.log('端末系統合ビューのヘッダー:');
    terminalHeaders.forEach((header, index) => {
      console.log(`  ${index + 1}列目: "${header}"`);
    });
    
    // 2. プリンタ・その他系統合ビューシートのヘッダーを確認
    const printerSheet = getIntegratedViewPrinterOtherSheet();
    const printerHeaders = printerSheet.getRange(1, 1, 1, printerSheet.getLastColumn()).getValues()[0];
    
    console.log('\nプリンタ・その他系統合ビューのヘッダー数:', printerHeaders.length);
    console.log('プリンタ・その他系統合ビューのヘッダー:');
    printerHeaders.forEach((header, index) => {
      console.log(`  ${index + 1}列目: "${header}"`);
    });
    
    // 3. マスタデータの取得テスト
    console.log('\n=== マスタデータ取得テスト ===');
    const terminalData = getTerminalMasterData();
    console.log('端末マスタデータ件数:', terminalData.length);
    if (terminalData.length > 0) {
      console.log('端末マスタデータのサンプル:');
      const sample = terminalData[0];
      for (const key in sample) {
        console.log(`  ${key}: ${sample[key]}`);
      }
    }
    
    // 4. ステータス収集データの取得テスト
    console.log('\n=== ステータス収集データ取得テスト ===');
    const statusData = getLatestStatusCollectionData();
    const statusKeys = Object.keys(statusData);
    console.log('ステータスデータ件数:', statusKeys.length);
    if (statusKeys.length > 0) {
      console.log('ステータスデータのサンプル:');
      const sampleKey = statusKeys[0];
      const sampleStatus = statusData[sampleKey];
      console.log(`  管理番号: ${sampleKey}`);
      for (const key in sampleStatus) {
        console.log(`    ${key}: ${sampleStatus[key]}`);
      }
    }
    
    // 5. 拠点マスタデータの取得テスト
    console.log('\n=== 拠点マスタデータ取得テスト ===');
    const locationData = getLocationMasterData();
    const locationKeys = Object.keys(locationData);
    console.log('拠点マスタデータ件数:', locationKeys.length);
    if (locationKeys.length > 0) {
      console.log('拠点マスタデータのサンプル:');
      const sampleLocationKey = locationKeys[0];
      const sampleLocation = locationData[sampleLocationKey];
      console.log(`  拠点コード: ${sampleLocationKey}`);
      console.log(`    拠点名: ${sampleLocation.locationName}`);
      console.log(`    管轄: ${sampleLocation.jurisdiction}`);
    }
    
    // 6. integrateDeviceData関数のテスト
    console.log('\n=== integrateDeviceData関数テスト ===');
    if (terminalData.length > 0) {
      const integratedData = integrateDeviceData(terminalData.slice(0, 3), statusData, locationData, 'terminal');
      console.log('統合データ件数:', integratedData.length);
      if (integratedData.length > 0) {
        console.log('統合データの列数:', integratedData[0].length);
        console.log('統合データのサンプル（最初の5列）:');
        const sample = integratedData[0];
        for (let i = 0; i < Math.min(5, sample.length); i++) {
          console.log(`  ${i + 1}列目: ${sample[i]}`);
        }
      }
    }
    
    console.log('\n=== テスト完了 ===');
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

/**
 * 統合ビューの更新をテストする関数
 */
function testUpdateIntegratedViews() {
  console.log('=== 統合ビュー更新テスト開始 ===');
  
  try {
    // 更新前のデータ数を確認
    const terminalSheet = getIntegratedViewTerminalSheet();
    const printerSheet = getIntegratedViewPrinterOtherSheet();
    
    const beforeTerminalRows = terminalSheet.getLastRow() - 1; // ヘッダーを除く
    const beforePrinterRows = printerSheet.getLastRow() - 1;
    
    console.log('更新前の端末系データ行数:', beforeTerminalRows);
    console.log('更新前のプリンタ系データ行数:', beforePrinterRows);
    
    // 統合ビューを更新
    const result = updateAllIntegratedViews();
    console.log('\n更新結果:', result);
    
    // 更新後のデータ数を確認
    const afterTerminalRows = terminalSheet.getLastRow() - 1;
    const afterPrinterRows = printerSheet.getLastRow() - 1;
    
    console.log('\n更新後の端末系データ行数:', afterTerminalRows);
    console.log('更新後のプリンタ系データ行数:', afterPrinterRows);
    
    // 更新後のデータサンプルを確認
    if (afterTerminalRows > 0) {
      const sampleData = terminalSheet.getRange(2, 1, 1, terminalSheet.getLastColumn()).getValues()[0];
      console.log('\n端末系統合ビューのデータサンプル（最初の行）:');
      const headers = terminalSheet.getRange(1, 1, 1, terminalSheet.getLastColumn()).getValues()[0];
      for (let i = 0; i < Math.min(10, sampleData.length); i++) {
        console.log(`  ${headers[i]}: ${sampleData[i]}`);
      }
    }
    
    console.log('\n=== テスト完了 ===');
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

/**
 * 統合ビューのデータ検証テスト
 */
function testValidateIntegratedViewData() {
  console.log('=== 統合ビューデータ検証テスト開始 ===');
  
  try {
    // 端末系統合ビューの検証
    const terminalSheet = getIntegratedViewTerminalSheet();
    const terminalHeaders = terminalSheet.getRange(1, 1, 1, terminalSheet.getLastColumn()).getValues()[0];
    
    // 期待されるヘッダー数をチェック
    console.log('端末系統合ビューの列数:', terminalHeaders.length);
    console.log('期待される列数: 46');
    
    // 重要な列が存在するか確認
    const requiredColumns = [
      '拠点管理番号', 'カテゴリ', '機種名', '製造番号', '資産管理番号',
      'ソフトウェア', 'OS', 'タイムスタンプ', '0-4.ステータス',
      '貸出日数', '要注意フラグ', '拠点名', '管轄', 'formURL', 'QRコードURL'
    ];
    
    console.log('\n必須列の存在確認:');
    for (const colName of requiredColumns) {
      const index = terminalHeaders.indexOf(colName);
      if (index >= 0) {
        console.log(`  ✓ ${colName} (列${index + 1})`);
      } else {
        console.log(`  ✗ ${colName} が見つかりません`);
      }
    }
    
    // データサンプルの検証
    if (terminalSheet.getLastRow() > 1) {
      const sampleData = terminalSheet.getRange(2, 1, 1, terminalSheet.getLastColumn()).getValues()[0];
      console.log('\nデータサンプルの検証:');
      
      // 拠点管理番号の形式確認
      const mgmtNumIndex = terminalHeaders.indexOf('拠点管理番号');
      if (mgmtNumIndex >= 0 && sampleData[mgmtNumIndex]) {
        const mgmtNum = sampleData[mgmtNumIndex];
        console.log(`  拠点管理番号: ${mgmtNum}`);
        const parts = mgmtNum.split('_');
        console.log(`    - 拠点コード: ${parts[0] || '(空)'}`);
        console.log(`    - カテゴリ: ${parts[1] || '(空)'}`);
        console.log(`    - 機種: ${parts[2] || '(空)'}`);
      }
      
      // ステータスの確認
      const statusIndex = terminalHeaders.indexOf('0-4.ステータス');
      if (statusIndex >= 0) {
        console.log(`  ステータス: ${sampleData[statusIndex] || '(空)'}`);
      }
      
      // 貸出日数と要注意フラグの確認
      const loanDaysIndex = terminalHeaders.indexOf('貸出日数');
      const cautionIndex = terminalHeaders.indexOf('要注意フラグ');
      if (loanDaysIndex >= 0 && cautionIndex >= 0) {
        console.log(`  貸出日数: ${sampleData[loanDaysIndex]}`);
        console.log(`  要注意フラグ: ${sampleData[cautionIndex]}`);
      }
    }
    
    console.log('\n=== テスト完了 ===');
    
  } catch (error) {
    console.error('テストエラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}