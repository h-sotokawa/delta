// 実際のintegrateDeviceData関数の詳細デバッグ
function debugActualIntegration() {
  console.log('=== 実際のintegrateDeviceData関数デバッグ ===');
  
  try {
    // 実際のデータを取得
    const terminalData = getTerminalMasterData();
    const statusData = getLatestStatusCollectionData();
    const locationData = getLocationMasterData();
    
    console.log('端末マスタデータ件数:', terminalData.length);
    console.log('ステータスデータ件数:', Object.keys(statusData).length);
    console.log('拠点マスタデータ件数:', Object.keys(locationData).length);
    
    if (terminalData.length > 0) {
      const testDevice = terminalData[0];
      const managementNumber = testDevice['拠点管理番号'];
      
      console.log('\n=== デバッグ情報 ===');
      console.log('テスト対象の拠点管理番号:', managementNumber);
      console.log('statusDataのキー:', Object.keys(statusData));
      console.log('管理番号がstatusDataに存在するか:', statusData.hasOwnProperty(managementNumber));
      
      // 実際のintegrateDeviceData関数の処理をステップごとに確認
      const latestStatus = statusData[managementNumber] || {};
      console.log('latestStatus取得結果:', latestStatus);
      console.log('latestStatusが空オブジェクトか:', Object.keys(latestStatus).length === 0);
      
      if (Object.keys(latestStatus).length > 0) {
        console.log('latestStatusのフィールド数:', Object.keys(latestStatus).length);
        console.log('タイムスタンプ:', latestStatus['タイムスタンプ']);
        console.log('9999.管理ID:', latestStatus['9999.管理ID']);
        console.log('0-4.ステータス:', latestStatus['0-4.ステータス']);
      } else {
        console.log('ERROR: latestStatusが空です');
      }
      
      // 実際のintegrateDeviceData関数を呼び出し
      console.log('\n=== 実際の関数呼び出し ===');
      const result = integrateDeviceData([testDevice], statusData, locationData, 'terminal');
      console.log('結果の行数:', result.length);
      if (result.length > 0) {
        console.log('結果の列数:', result[0].length);
        console.log('最初の15列:');
        for (let i = 0; i < Math.min(15, result[0].length); i++) {
          console.log(`  ${i+1}: ${result[0][i] || '(空)'}`);
        }
      }
    }
    
  } catch (error) {
    console.error('エラー:', error);
    console.error('スタックトレース:', error.stack);
  }
}

// integrateDeviceData関数の修正版
function integrateDeviceDataFixed(deviceData, statusData, locationMap, deviceType) {
  console.log('=== 修正版integrateDeviceData関数 ===');
  const integratedRows = [];
  
  for (const device of deviceData) {
    const managementNumber = device['拠点管理番号'];
    if (!managementNumber) continue;
    
    console.log(`処理中の拠点管理番号: ${managementNumber}`);
    
    // ステータスデータから最新情報を取得
    const latestStatus = statusData[managementNumber];
    console.log(`latestStatus取得結果: ${latestStatus ? 'あり' : 'なし'}`);
    
    if (latestStatus) {
      console.log(`latestStatusのフィールド数: ${Object.keys(latestStatus).length}`);
    }
    
    // 拠点情報を取得
    const locationCode = managementNumber.split('_')[0] || '';
    const locationInfo = locationMap[locationCode] || {};
    
    // 貸出日数を計算
    let loanDays = 0;
    if (latestStatus && latestStatus['0-4.ステータス'] === '1.貸出中' && latestStatus['タイムスタンプ']) {
      const loanDate = new Date(latestStatus['タイムスタンプ']);
      const today = new Date();
      loanDays = Math.floor((today - loanDate) / (1000 * 60 * 60 * 24));
    }
    
    // 要注意フラグを判定
    const cautionFlag = loanDays >= 90 || (latestStatus && latestStatus['3-0.社内ステータス'] === '1.修理中');
    
    if (deviceType === 'terminal') {
      // 端末系統合ビュー（46列）
      const integratedRow = [
        // マスタシートデータ（A-G列）
        managementNumber,
        device['カテゴリ'] || '',
        device['機種名'] || '',
        device['製造番号'] || '',
        device['資産管理番号'] || '',
        device['ソフトウェア'] || '',
        device['OS'] || '',
        
        // 収集シートデータ（H-AN列）
        (latestStatus && latestStatus['タイムスタンプ']) || '',
        (latestStatus && latestStatus['9999.管理ID']) || '',
        (latestStatus && latestStatus['0-0.拠点管理番号']) || managementNumber,
        (latestStatus && latestStatus['0-1.担当者']) || '',
        (latestStatus && latestStatus['0-2.EMシステムズの社員ですか？']) || '',
        (latestStatus && latestStatus['0-3.所属会社']) || '',
        (latestStatus && latestStatus['0-4.ステータス']) || '',
        (latestStatus && latestStatus['1-1.顧客名または貸出先']) || '',
        (latestStatus && latestStatus['1-2.顧客番号']) || '',
        (latestStatus && latestStatus['1-3.住所']) || '',
        (latestStatus && latestStatus['1-4.ユーザー機の預り有無']) || '',
        (latestStatus && latestStatus['1-5.依頼者']) || '',
        (latestStatus && latestStatus['1-6.備考']) || '',
        (latestStatus && latestStatus['1-7.預りユーザー機のシリアルNo.(製造番号)']) || '',
        (latestStatus && latestStatus['1-8.お預かり証No.']) || '',
        (latestStatus && latestStatus['2-1.預り機返却の有無']) || '',
        (latestStatus && latestStatus['2-2.依頼者']) || '',
        (latestStatus && latestStatus['2-3.備考']) || '',
        (latestStatus && latestStatus['3-0.社内ステータス']) || '',
        (latestStatus && latestStatus['3-0-1.棚卸しフラグ']) || '',
        (latestStatus && latestStatus['3-0-2.拠点']) || locationCode,
        (latestStatus && latestStatus['3-1-1.ソフト']) || '',
        (latestStatus && latestStatus['3-1-2.備考']) || '',
        (latestStatus && latestStatus['3-2-1.端末初期化の引継ぎ']) || '',
        (latestStatus && latestStatus['3-2-2.備考']) || '',
        (latestStatus && latestStatus['3-2-3.引継ぎ担当者']) || '',
        (latestStatus && latestStatus['3-2-4.初期化作業の引継ぎ']) || '',
        (latestStatus && latestStatus['4-1.所在']) || '',
        (latestStatus && latestStatus['4-2.持ち出し理由']) || '',
        (latestStatus && latestStatus['4-3.備考']) || '',
        (latestStatus && latestStatus['5-1.内容']) || '',
        (latestStatus && latestStatus['5-2.所在']) || '',
        (latestStatus && latestStatus['5-3.備考']) || '',
        
        // 計算フィールド（AO-AP列）
        loanDays,
        cautionFlag,
        
        // 参照データ（AQ-AT列）
        locationInfo.locationName || locationCode,
        locationInfo.jurisdiction || '',
        device['formURL'] || '',
        device['QRコードURL'] || ''
      ];
      
      console.log(`統合行の列数: ${integratedRow.length}`);
      integratedRows.push(integratedRow);
    }
  }
  
  return integratedRows;
}