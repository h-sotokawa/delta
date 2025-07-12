/**
 * QRコード画像データを生成（シンプル版）
 * @param {string} url - QRコードに埋め込むURL
 * @return {Object} 画像データを含む結果オブジェクト
 */
function generateQRCodeImage(url) {
  const startTime = startPerformanceTimer();
  addLog("QRコード画像生成開始（修正版）", { url: url });

  try {
    if (!url) {
      throw new Error("URLが指定されていません");
    }

    // QR Serverの最もシンプルなURL
    const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?data=' + encodeURIComponent(url);
    
    addLog("QRコード生成URL", { url: qrUrl });
    
    // シンプルにfetch
    const response = UrlFetchApp.fetch(qrUrl);
    const blob = response.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    addLog("QRコード生成成功", { 
      size: blob.getBytes().length,
      contentType: blob.getContentType()
    });
    
    return {
      success: true,
      imageData: 'data:image/png;base64,' + base64,
      provider: 'QR Server',
      url: url
    };
    
  } catch (error) {
    addLog("QRコード生成エラー", { 
      error: error.toString(),
      url: url 
    });
    
    // エラーでも最低限URLは返す
    return {
      success: true,
      url: url,
      imageData: null,
      provider: 'None',
      imageError: error.toString()
    };
  }
}