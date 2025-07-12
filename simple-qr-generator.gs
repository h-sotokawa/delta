/**
 * シンプルなQRコード画像生成（テスト用）
 * @param {string} text - QRコードに埋め込むテキスト
 * @return {Object} 画像データを含む結果オブジェクト
 */
function generateSimpleQRCode(text) {
  try {
    // 最もシンプルなQRコード生成URL
    const qrUrl = 'https://chart.googleapis.com/chart?chs=200x200&cht=qr&chl=' + encodeURIComponent(text);
    
    console.log('QRコード生成URL:', qrUrl);
    
    // 直接画像を取得
    const response = UrlFetchApp.fetch(qrUrl);
    
    if (response.getResponseCode() === 200) {
      const blob = response.getBlob();
      const base64 = Utilities.base64Encode(blob.getBytes());
      
      return {
        success: true,
        imageData: 'data:image/png;base64,' + base64,
        url: text
      };
    }
    
    return {
      success: false,
      error: 'Failed to generate QR code'
    };
    
  } catch (error) {
    console.error('QRコード生成エラー:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * QRコード画像データを生成（改善版）
 * @param {string} url - QRコードに埋め込むURL
 * @return {Object} 画像データを含む結果オブジェクト
 */
function generateQRCodeImageV2(url) {
  const startTime = Date.now();
  
  try {
    if (!url) {
      throw new Error("URLが指定されていません");
    }

    console.log("QRコード画像生成開始 V2:", url);

    // 方法1: QR Serverの最もシンプルなエンドポイント
    try {
      const qrServerUrl = 'https://api.qrserver.com/v1/create-qr-code/?data=' + 
        encodeURIComponent(url) + '&size=200x200';
      
      console.log("QR Server URL:", qrServerUrl);
      
      const options = {
        method: 'get',
        contentType: 'application/octet-stream',
        muteHttpExceptions: false
      };
      
      const response = UrlFetchApp.fetch(qrServerUrl, options);
      const blob = response.getBlob();
      
      if (blob.getBytes().length > 0) {
        const base64Data = Utilities.base64Encode(blob.getBytes());
        
        console.log("QRコード画像生成成功 (QR Server):", blob.getBytes().length, "bytes");
        
        return {
          success: true,
          imageData: 'data:image/png;base64,' + base64Data,
          provider: 'QR Server',
          url: url
        };
      }
    } catch (e1) {
      console.error("QR Server エラー:", e1.toString());
    }

    // 方法2: QuickChartの最もシンプルな方法
    try {
      const quickChartUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(url);
      
      console.log("QuickChart URL:", quickChartUrl);
      
      const response = UrlFetchApp.fetch(quickChartUrl);
      const blob = response.getBlob();
      
      if (blob.getBytes().length > 0) {
        const base64Data = Utilities.base64Encode(blob.getBytes());
        
        console.log("QRコード画像生成成功 (QuickChart):", blob.getBytes().length, "bytes");
        
        return {
          success: true,
          imageData: 'data:image/png;base64,' + base64Data,
          provider: 'QuickChart',
          url: url
        };
      }
    } catch (e2) {
      console.error("QuickChart エラー:", e2.toString());
    }

    // 方法3: Google Charts API（廃止予定だが、まだ動く可能性）
    try {
      const googleChartsUrl = 'https://chart.googleapis.com/chart?chs=200x200&cht=qr&chl=' + 
        encodeURIComponent(url);
      
      console.log("Google Charts URL:", googleChartsUrl);
      
      const response = UrlFetchApp.fetch(googleChartsUrl);
      const blob = response.getBlob();
      
      if (blob.getBytes().length > 0) {
        const base64Data = Utilities.base64Encode(blob.getBytes());
        
        console.log("QRコード画像生成成功 (Google Charts):", blob.getBytes().length, "bytes");
        
        return {
          success: true,
          imageData: 'data:image/png;base64,' + base64Data,
          provider: 'Google Charts',
          url: url
        };
      }
    } catch (e3) {
      console.error("Google Charts エラー:", e3.toString());
    }

    // 全て失敗した場合
    return {
      success: false,
      error: "すべてのQRコード生成APIが失敗しました",
      url: url
    };

  } catch (error) {
    console.error("QRコード画像生成エラー:", error);
    return {
      success: false,
      error: error.toString(),
      url: url
    };
  }
}