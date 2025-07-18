// latestStatusの取得をデバッグする関数
function debugLatestStatusRetrieval() {
  console.log('=== latestStatus取得デバッグ ===');
  
  try {
    // 1. ステータス収集データを取得
    const statusData = getLatestStatusCollectionData();
    console.log('ステータスデータ件数:', Object.keys(statusData).length);
    
    // 2. 端末マスタデータを取得
    const terminalData = getTerminalMasterData();
    console.log('端末マスタデータ件数:', terminalData.length);
    
    if (terminalData.length > 0) {
      const testDevice = terminalData[0];
      const managementNumber = testDevice['拠点管理番号'];
      console.log('テスト対象の拠点管理番号:', managementNumber);
      
      // 3. statusDataから対応するデータを取得
      const latestStatus = statusData[managementNumber];
      console.log('latestStatusが取得できるか:', latestStatus ? 'はい' : 'いいえ');
      
      if (latestStatus) {
        console.log('latestStatusのフィールド数:', Object.keys(latestStatus).length);
        console.log('latestStatusの主要フィールド:');
        console.log('  タイムスタンプ:', latestStatus['タイムスタンプ']);
        console.log('  9999.管理ID:', latestStatus['9999.管理ID']);
        console.log('  0-4.ステータス:', latestStatus['0-4.ステータス']);
        console.log('  0-1.担当者:', latestStatus['0-1.担当者']);
        
        // 4. 実際のintegrateDeviceData関数の処理をシミュレート
        console.log('\n=== integrateDeviceData関数の処理シミュレート ===');
        const locationData = getLocationMasterData();
        const locationCode = managementNumber.split('_')[0] || '';
        const locationInfo = locationData[locationCode] || {};
        
        console.log('拠点コード:', locationCode);
        console.log('拠点情報:', locationInfo);
        
        // 貸出日数計算
        let loanDays = 0;
        if (latestStatus['0-4.ステータス'] === '1.貸出中' && latestStatus['タイムスタンプ']) {
          const loanDate = new Date(latestStatus['タイムスタンプ']);
          const today = new Date();
          loanDays = Math.floor((today - loanDate) / (1000 * 60 * 60 * 24));
        }
        console.log('計算された貸出日数:', loanDays);
        
        // 要注意フラグ
        const cautionFlag = loanDays >= 90 || latestStatus['3-0.社内ステータス'] === '1.修理中';
        console.log('要注意フラグ:', cautionFlag);
        
        // 統合行を作成してみる
        const integratedRow = [
          // マスタシートデータ（A-G列）
          managementNumber,
          testDevice['カテゴリ'] || '',
          testDevice['機種名'] || '',
          testDevice['製造番号'] || '',
          testDevice['資産管理番号'] || '',
          testDevice['ソフトウェア'] || '',
          testDevice['OS'] || '',
          
          // 収集シートデータ（H-AN列）
          latestStatus['タイムスタンプ'] || '',
          latestStatus['9999.管理ID'] || '',
          latestStatus['0-0.拠点管理番号'] || managementNumber,
          latestStatus['0-1.担当者'] || '',
          latestStatus['0-2.EMシステムズの社員ですか？'] || '',
          latestStatus['0-3.所属会社'] || '',
          latestStatus['0-4.ステータス'] || '',
          latestStatus['1-1.顧客名または貸出先'] || '',
          latestStatus['1-2.顧客番号'] || '',
          latestStatus['1-3.住所'] || '',
          latestStatus['1-4.ユーザー機の預り有無'] || '',
          latestStatus['1-5.依頼者'] || '',
          latestStatus['1-6.備考'] || '',
          latestStatus['1-7.預りユーザー機のシリアルNo.(製造番号)'] || '',
          latestStatus['1-8.お預かり証No.'] || '',
          latestStatus['2-1.預り機返却の有無'] || '',
          latestStatus['2-2.依頼者'] || '',
          latestStatus['2-3.備考'] || '',
          latestStatus['3-0.社内ステータス'] || '',
          latestStatus['3-0-1.棚卸しフラグ'] || '',
          latestStatus['3-0-2.拠点'] || locationCode,
          latestStatus['3-1-1.ソフト'] || '',
          latestStatus['3-1-2.備考'] || '',
          latestStatus['3-2-1.端末初期化の引継ぎ'] || '',
          latestStatus['3-2-2.備考'] || '',
          latestStatus['3-2-3.引継ぎ担当者'] || '',
          latestStatus['3-2-4.初期化作業の引継ぎ'] || '',
          latestStatus['4-1.所在'] || '',
          latestStatus['4-2.持ち出し理由'] || '',
          latestStatus['4-3.備考'] || '',
          latestStatus['5-1.内容'] || '',
          latestStatus['5-2.所在'] || '',
          latestStatus['5-3.備考'] || '',
          
          // 計算フィールド（AO-AP列）
          loanDays,
          cautionFlag,
          
          // 参照データ（AQ-AT列）
          locationInfo.locationName || locationCode,
          locationInfo.jurisdiction || '',
          testDevice['formURL'] || '',
          testDevice['QRコードURL'] || ''
        ];
        
        console.log('\n統合行の列数:', integratedRow.length);
        console.log('統合行の最初の15列:');
        for (let i = 0; i < Math.min(15, integratedRow.length); i++) {
          console.log(`  ${i+1}: ${integratedRow[i] || '(空)'}`);
        }
        
      } else {
        console.log('ERROR: latestStatusが取得できません');
        console.log('statusDataのキー:', Object.keys(statusData));
        console.log('探している拠点管理番号:', managementNumber);
      }
    }
    
  } catch (error) {
    console.error('エラー:', error);
  }
}