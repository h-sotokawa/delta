/**
 * ハイブリッドQRコード生成ソリューション
 * サーバーサイドとクライアントサイドの両方のアプローチを組み合わせる
 */

/**
 * 改良版QRコード画像生成関数
 * 複数の方法を試して、最も信頼性の高い方法でQRコードを生成する
 */
function generateQRCodeImageHybrid(url) {
  const startTime = startPerformanceTimer();
  addLog("ハイブリッドQRコード生成開始", { url: url });

  try {
    if (!url) {
      throw new Error("URLが指定されていません");
    }

    // 方法1: 直接画像URLを返す（ブラウザで処理）
    // この方法は最も単純で、ブラウザ側でCORSの制約を受けにくい
    const directImageUrls = [
      {
        name: 'QR Server Direct',
        url: `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(url)}&size=200x200`,
        provider: 'QR Server (Direct Link)'
      },
      {
        name: 'QuickChart Direct',
        url: `https://quickchart.io/qr?text=${encodeURIComponent(url)}&size=200`,
        provider: 'QuickChart (Direct Link)'
      },
      {
        name: 'QR Code API Direct',
        url: `https://qr-api.izo.dev/v1/create?data=${encodeURIComponent(url)}&size=200`,
        provider: 'QR Code API (Direct Link)'
      }
    ];

    // 最初に直接URLアプローチを試す
    for (const api of directImageUrls) {
      try {
        addLog(`${api.name} - 直接URL生成`, { url: api.url });
        
        // URLの妥当性をチェック（実際にfetchはしない）
        if (api.url && api.url.startsWith('https://')) {
          return {
            success: true,
            imageUrl: api.url,  // Base64ではなく直接URLを返す
            provider: api.provider,
            url: url,
            method: 'direct-url'
          };
        }
      } catch (e) {
        addLog(`${api.name} エラー`, { error: e.toString() });
      }
    }

    // 方法2: サーバーサイドでBase64変換を試みる（以前の方法）
    const serverSideApis = [
      {
        name: 'QR Server',
        url: `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(url)}`,
        options: {
          method: 'GET',
          muteHttpExceptions: true,
          validateHttpsCertificates: false
        }
      }
    ];

    for (const api of serverSideApis) {
      try {
        addLog(`${api.name} - サーバーサイド試行`, { url: api.url });
        
        const response = UrlFetchApp.fetch(api.url, api.options);
        
        if (response.getResponseCode() === 200) {
          const blob = response.getBlob();
          const bytes = blob.getBytes();
          
          if (bytes.length > 0) {
            const base64 = Utilities.base64Encode(bytes);
            
            addLog(`${api.name} - 成功`, { size: bytes.length });
            
            return {
              success: true,
              imageData: `data:image/png;base64,${base64}`,
              provider: api.name,
              url: url,
              method: 'server-base64'
            };
          }
        }
      } catch (e) {
        addLog(`${api.name} サーバーサイドエラー`, { error: e.toString() });
      }
    }

    // 方法3: クライアントサイド生成を指示
    addLog("サーバーサイド生成失敗 - クライアントサイド生成を推奨");
    
    return {
      success: true,
      url: url,
      imageUrl: directImageUrls[0].url, // フォールバックとして最初のURLを使用
      provider: 'Client-side generation recommended',
      method: 'client-fallback',
      clientInstructions: {
        message: 'サーバーサイドでの生成に失敗しました。ブラウザで直接生成します。',
        fallbackUrl: directImageUrls[0].url
      }
    };

  } catch (error) {
    addLog("ハイブリッドQRコード生成エラー", { error: error.toString() });
    endPerformanceTimer(startTime, "QRコード生成エラー");
    
    // エラーでも最低限のURLは返す
    return {
      success: false,
      error: error.toString(),
      url: url,
      imageUrl: `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(url)}`,
      method: 'error-fallback'
    };
  }
}

/**
 * クライアントから画像データを受け取る関数
 */
function receiveQRCodeImage(imageData, originalUrl) {
  addLog("クライアントからQRコード画像受信", { 
    url: originalUrl,
    dataLength: imageData ? imageData.length : 0
  });
  
  // 受信したデータを保存または処理
  const cache = CacheService.getScriptCache();
  const cacheKey = `qr_image_${Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, originalUrl)}`;
  
  try {
    // キャッシュに保存（最大100KB）
    if (imageData && imageData.length < 100000) {
      cache.put(cacheKey, imageData, 600); // 10分間キャッシュ
      return { success: true, cached: true };
    }
  } catch (e) {
    addLog("キャッシュ保存エラー", { error: e.toString() });
  }
  
  return { success: true, cached: false };
}

/**
 * クライアントからのエラー報告を受け取る
 */
function reportQRGenerationError(errorMessage) {
  addLog("クライアントQRコード生成エラー", { error: errorMessage });
  return { received: true };
}

/**
 * 既存のgenerateQRCodeImage関数を置き換える
 */
function generateQRCodeImage(url) {
  // ハイブリッドソリューションを使用
  return generateQRCodeImageHybrid(url);
}