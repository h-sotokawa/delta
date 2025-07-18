/**
 * 統合ビュー リアルタイム更新トリガー関数
 * フォーム送信時、マスタ変更時、定期実行時の処理
 */

/**
 * フォーム送信時の統合ビュー更新
 * onFormSubmitトリガーで実行される
 */
function updateIntegratedViewOnSubmit(e) {
  try {
    console.log('=== フォーム送信による統合ビュー更新開始 ===');
    
    // 送信されたデータから拠点管理番号を取得
    const submittedData = e.values;
    const locationNumber = submittedData[1]; // 拠点管理番号は通常2列目
    
    if (!locationNumber) {
      console.error('拠点管理番号が見つかりません');
      return;
    }
    
    console.log('更新対象拠点管理番号:', locationNumber);
    
    // カテゴリを判定して適切な統合ビューを更新
    const category = extractCategoryFromLocationNumber(locationNumber);
    const updateResult = updateSpecificIntegratedView(locationNumber, category);
    
    // 検索インデックスも更新
    updateSearchIndex(locationNumber);
    
    console.log('=== フォーム送信による統合ビュー更新完了 ===');
    return updateResult;
    
  } catch (error) {
    console.error('フォーム送信時の更新エラー:', error);
    throw error;
  }
}

/**
 * マスタシート変更時の統合ビュー更新
 * onChangeトリガーで実行される
 */
function updateIntegratedViewOnChange(e) {
  try {
    // 変更されたシートを特定
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    
    console.log('変更されたシート:', sheetName);
    
    // マスタシートの変更の場合のみ処理
    if (sheetName.includes('マスタ') && !sheetName.includes('データタイプ')) {
      console.log('=== マスタ変更による統合ビュー更新開始 ===');
      
      // 変更された行の拠点管理番号を取得
      const range = sheet.getActiveRange();
      const row = range.getRow();
      
      if (row > 1) { // ヘッダー行以外
        const locationNumber = sheet.getRange(row, 1).getValue(); // A列が拠点管理番号
        
        if (locationNumber) {
          const category = extractCategoryFromLocationNumber(locationNumber);
          updateSpecificIntegratedView(locationNumber, category);
          updateSearchIndex(locationNumber);
        }
      }
      
      console.log('=== マスタ変更による統合ビュー更新完了 ===');
    }
    
  } catch (error) {
    console.error('マスタ変更時の更新エラー:', error);
    // エラーログは記録するが、処理は継続
  }
}

/**
 * 全統合ビューの再構築（日次バッチ）
 * timeBasedトリガーで深夜2:00に実行
 */
function rebuildAllIntegratedViews() {
  try {
    console.log('=== 全統合ビュー再構築開始 ===');
    const startTime = new Date();
    
    // 端末系統合ビューの再構築
    rebuildTerminalIntegratedView();
    
    // プリンタ・その他系統合ビューの再構築
    rebuildPrinterOtherIntegratedView();
    
    // 検索インデックスの再構築
    rebuildSearchIndex();
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    console.log(`=== 全統合ビュー再構築完了 (処理時間: ${duration}秒) ===`);
    
    // 完了通知（必要に応じて）
    notifyRebuildCompletion(duration);
    
  } catch (error) {
    console.error('全統合ビュー再構築エラー:', error);
    notifyRebuildError(error);
  }
}

/**
 * 拠点管理番号からカテゴリを抽出
 */
function extractCategoryFromLocationNumber(locationNumber) {
  if (!locationNumber || typeof locationNumber !== 'string') {
    return 'unknown';
  }
  
  // フォーマット: 拠点_カテゴリ_モデル_製造番号_連番
  const parts = locationNumber.split('_');
  if (parts.length >= 2) {
    const categoryCode = parts[1];
    
    // カテゴリコードの判定
    if (['SV', 'CL', 'Server', 'Desktop', 'Laptop', 'Tablet'].includes(categoryCode)) {
      return 'terminal';
    } else if (['PR', 'OT', 'Printer', 'Router', 'Hub', 'Other'].includes(categoryCode)) {
      return 'printer_other';
    }
  }
  
  return 'unknown';
}

/**
 * 特定の統合ビューを更新
 */
function updateSpecificIntegratedView(locationNumber, category) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  if (category === 'terminal') {
    const sheet = spreadsheet.getSheetByName('integrated_view_terminal');
    return updateIntegratedViewRow(sheet, locationNumber, 'terminal');
  } else if (category === 'printer_other') {
    const sheet = spreadsheet.getSheetByName('integrated_view_printer_other');
    return updateIntegratedViewRow(sheet, locationNumber, 'printer_other');
  }
  
  console.warn('不明なカテゴリ:', category);
  return false;
}

/**
 * 統合ビューの行を更新
 */
function updateIntegratedViewRow(sheet, locationNumber, viewType) {
  if (!sheet) {
    console.error('統合ビューシートが見つかりません:', viewType);
    return false;
  }
  
  try {
    // 既存の行を検索
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    let targetRow = -1;
    
    // ヘッダーを取得
    const viewHeaders = values[0];
    const locationNumberIndex = getColumnIndex(viewHeaders, '拠点管理番号');
    
    if (locationNumberIndex < 0) {
      console.error('拠点管理番号列が見つかりません');
      return false;
    }
    
    // 拠点管理番号で行を検索
    for (let i = 1; i < values.length; i++) {
      if (values[i][locationNumberIndex] === locationNumber) {
        targetRow = i + 1; // シートの行番号は1から始まる
        break;
      }
    }
    
    // データを収集
    const integratedData = collectIntegratedData(locationNumber, viewType);
    
    if (!integratedData || integratedData.length === 0) {
      console.warn('統合データが見つかりません:', locationNumber);
      return false;
    }
    
    if (targetRow > 0) {
      // 既存行を更新
      const updateRange = sheet.getRange(targetRow, 1, 1, integratedData.length);
      updateRange.setValues([integratedData]);
      console.log(`更新完了: ${locationNumber} (行${targetRow})`);
    } else {
      // 新規行を追加
      sheet.appendRow(integratedData);
      console.log(`新規追加: ${locationNumber}`);
    }
    
    return true;
    
  } catch (error) {
    console.error('統合ビュー行更新エラー:', error);
    return false;
  }
}

/**
 * 統合データを収集
 */
function collectIntegratedData(locationNumber, viewType) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const data = [];
  
  try {
    if (viewType === 'terminal') {
      // 端末マスタからデータ取得
      const terminalMaster = getDataFromMaster('端末マスタ', locationNumber);
      if (!terminalMaster) return null;
      
      // 端末ステータス収集から最新データ取得
      const statusData = getLatestStatusData('端末ステータス収集', locationNumber);
      
      // 拠点マスタから拠点情報取得
      const locationData = getLocationData(terminalMaster[1]); // 拠点コード
      
      // データ結合（端末系: A-AT 46列）
      data.push(...terminalMaster); // A-G: 基本情報
      data.push(...(statusData || Array(36).fill(''))); // H-AN: ステータス情報
      data.push(locationData.locationName || ''); // AO: 拠点名
      data.push(locationData.jurisdiction || ''); // AP: 管轄
      
      // 計算フィールド追加
      if (statusData && statusData[2]) { // 貸出開始日時
        const loanDays = calculateLoanDays(statusData[2]);
        data[11] = loanDays; // L列: 貸出日数
      }
      
    } else if (viewType === 'printer_other') {
      // プリンタマスタまたはその他マスタからデータ取得
      let masterData = getDataFromMaster('プリンタマスタ', locationNumber);
      if (!masterData) {
        masterData = getDataFromMaster('その他マスタ', locationNumber);
      }
      if (!masterData) return null;
      
      // プリンタステータス収集から最新データ取得
      const statusData = getLatestStatusData('プリンタステータス収集', locationNumber);
      
      // 拠点マスタから拠点情報取得
      const locationData = getLocationData(masterData[1]); // 拠点コード
      
      // データ結合（プリンタ・その他系: A-AU 47列）
      data.push(...masterData); // A-D/E: 基本情報
      data.push(...(statusData || Array(38).fill(''))); // F-AN: ステータス情報
      data.push(locationData.locationName || ''); // AO: 拠点名
      data.push(locationData.jurisdiction || ''); // AP: 管轄
    }
    
    return data;
    
  } catch (error) {
    console.error('データ収集エラー:', error);
    return null;
  }
}

/**
 * マスタシートからデータ取得
 */
function getDataFromMaster(sheetName, locationNumber) {
  const result = getDataFromMasterWithHeaders(sheetName, locationNumber);
  return result ? result.data : null;
}

/**
 * マスタシートからデータとヘッダーを取得
 */
function getDataFromMasterWithHeaders(sheetName, locationNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return null;
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  
  // 拠点管理番号の列を動的に取得
  const locationNumberIndex = getColumnIndex(headers, '拠点管理番号');
  if (locationNumberIndex < 0) return null;
  
  // 拠点管理番号で検索
  for (let i = 1; i < values.length; i++) {
    if (values[i][locationNumberIndex] === locationNumber) {
      const columnCount = sheetName === '端末マスタ' ? 7 : 
                         sheetName === 'その他マスタ' ? 5 : 4;
      return {
        data: values[i].slice(0, columnCount),
        headers: headers.slice(0, columnCount)
      };
    }
  }
  
  return null;
}

/**
 * 収集シートから最新ステータスデータ取得
 */
function getLatestStatusData(sheetName, locationNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return null;
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  
  // 拠点管理番号の列を動的に取得
  const locationNumberIndex = getColumnIndex(headers, '拠点管理番号');
  const timestampIndex = getColumnIndex(headers, 'タイムスタンプ');
  
  if (locationNumberIndex < 0) return null;
  
  // 拠点管理番号で該当するすべての行を取得
  const matchingRows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][locationNumberIndex] === locationNumber) {
      matchingRows.push({
        row: values[i],
        timestamp: timestampIndex >= 0 ? new Date(values[i][timestampIndex]) : new Date()
      });
    }
  }
  
  // タイムスタンプでソートして最新を取得
  if (matchingRows.length > 0) {
    matchingRows.sort((a, b) => b.timestamp - a.timestamp);
    const latestRow = matchingRows[0].row;
    
    // ステータスデータ部分を返す
    if (sheetName === '端末ステータス収集') {
      return latestRow.slice(0, 40); // タイムスタンプから必要な列まで
    } else {
      return latestRow.slice(0, 41); // プリンタの場合
    }
  }
  
  return null;
}

/**
 * 拠点情報を取得
 */
function getLocationData(locationCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('拠点マスタ');
  if (!sheet) return { locationName: '', jurisdiction: '' };
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  
  // 列を動的に取得
  const locationCodeIndex = getColumnIndex(headers, '拠点コード');
  const locationNameIndex = getColumnIndex(headers, '拠点名');
  const jurisdictionIndex = getColumnIndex(headers, '管轄');
  
  if (locationCodeIndex < 0) return { locationName: '', jurisdiction: '' };
  
  // 拠点コードで検索
  for (let i = 1; i < values.length; i++) {
    if (values[i][locationCodeIndex] === locationCode) {
      return {
        locationName: locationNameIndex >= 0 ? (values[i][locationNameIndex] || '') : '',
        jurisdiction: jurisdictionIndex >= 0 ? (values[i][jurisdictionIndex] || '') : ''
      };
    }
  }
  
  return { locationName: '', jurisdiction: '' };
}

/**
 * 貸出日数を計算
 */
function calculateLoanDays(loanStartDate) {
  if (!loanStartDate) return 0;
  
  const start = new Date(loanStartDate);
  const today = new Date();
  const diffTime = Math.abs(today - start);
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  
  return diffDays;
}

/**
 * 検索インデックスを更新
 */
function updateSearchIndex(locationNumber) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('search_index');
    if (!sheet) return;
    
    // 機器情報を取得
    const category = extractCategoryFromLocationNumber(locationNumber);
    const integratedData = collectIntegratedData(locationNumber, 
                                                 category === 'terminal' ? 'terminal' : 'printer_other');
    
    if (!integratedData) return;
    
    // 検索キー生成
    const searchKey = generateSearchKey(integratedData, locationNumber);
    
    // インデックス更新
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    let targetRow = -1;
    
    // ヘッダーを取得
    const indexHeaders = values[0];
    const indexLocationNumberIndex = getColumnIndex(indexHeaders, '拠点管理番号');
    
    if (indexLocationNumberIndex < 0) return;
    
    // 既存エントリを検索
    for (let i = 1; i < values.length; i++) {
      if (values[i][indexLocationNumberIndex] === locationNumber) {
        targetRow = i + 1;
        break;
      }
    }
    
    const indexData = [
      locationNumber,
      searchKey,
      integratedData[2] || '', // カテゴリ
      integratedData[1] || '', // 拠点
      integratedData[8] || '', // 状態
      new Date() // 最終更新日時
    ];
    
    if (targetRow > 0) {
      // 更新
      sheet.getRange(targetRow, 1, 1, indexData.length).setValues([indexData]);
    } else {
      // 新規追加
      sheet.appendRow(indexData);
    }
    
  } catch (error) {
    console.error('検索インデックス更新エラー:', error);
  }
}

/**
 * 検索キーを生成
 */
function generateSearchKey(data, locationNumber) {
  const parts = [
    locationNumber,
    data[1] || '', // 拠点
    data[2] || '', // カテゴリ
    data[3] || '', // 機種名
    data[8] || '', // 状態
    data[data.length - 1] || '' // 管轄
  ];
  
  return parts.filter(p => p).join(' ');
}

/**
 * 再構築完了通知
 */
function notifyRebuildCompletion(duration) {
  const email = PropertiesService.getScriptProperties().getProperty('ALERT_NOTIFICATION_EMAIL');
  if (!email) return;
  
  const subject = '統合ビュー再構築完了';
  const body = `統合ビューの日次再構築が正常に完了しました。

処理時間: ${duration}秒
実行日時: ${new Date().toLocaleString('ja-JP')}

システムは正常に稼働しています。`;
  
  try {
    GmailApp.sendEmail(email, subject, body);
  } catch (error) {
    console.error('通知メール送信エラー:', error);
  }
}

/**
 * 再構築エラー通知
 */
function notifyRebuildError(error) {
  const email = PropertiesService.getScriptProperties().getProperty('ERROR_NOTIFICATION_EMAIL');
  if (!email) return;
  
  const subject = '統合ビュー再構築エラー';
  const body = `統合ビューの日次再構築でエラーが発生しました。

エラー内容: ${error.toString()}
実行日時: ${new Date().toLocaleString('ja-JP')}

システム管理者による確認が必要です。`;
  
  try {
    GmailApp.sendEmail(email, subject, body);
  } catch (error) {
    console.error('エラー通知メール送信失敗:', error);
  }
}