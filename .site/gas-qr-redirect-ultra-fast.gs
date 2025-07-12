// QRコード中間ページ - 超高速版
// 302リダイレクトを使用して即座に遷移

function doGet(e) {
  const locationNumber = e.parameter.id || '';
  
  if (!locationNumber) {
    // エラーページのみHTMLを返す
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>エラー</title>
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, sans-serif;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
            margin: 0;
            background: #f5f5f5;
          }
          .error {
            background: white;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 2px 20px rgba(0,0,0,0.1);
            text-align: center;
            max-width: 400px;
          }
          h2 { color: #dc2626; margin-bottom: 20px; }
          p { color: #666; line-height: 1.5; }
        </style>
      </head>
      <body>
        <div class="error">
          <h2>⚠️ エラー</h2>
          <p>拠点管理番号が指定されていません。</p>
          <p>QRコードを正しく読み取ってください。</p>
        </div>
      </body>
      </html>
    `).setTitle('エラー').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // フォームURLを高速に構築
  const redirectUrl = buildFormUrl(locationNumber);
  
  // 最速のリダイレクト方法：ContentServiceを使用
  return ContentService
    .createTextOutput(`<html><head><meta http-equiv="refresh" content="0; url=${redirectUrl}"></head><body></body></html>`)
    .setMimeType(ContentService.MimeType.HTML);
}

// 高速化のため関数を最適化
function buildFormUrl(locationNumber) {
  // カテゴリを高速判定
  const parts = locationNumber.split('_');
  const category = parts[1] ? parts[1].toLowerCase() : 'desktop';
  
  // プリンタ系かどうかを判定
  const isPrinter = category === 'printer' || category === 'other';
  
  // URLを直接構築（条件分岐を最小化）
  const baseUrl = isPrinter 
    ? 'https://docs.google.com/forms/d/e/1FAIpQLSfkxqPsHQpeKhWTBO6y3dUyvHoUKw365Wrdzsi1NNMbJ6C5mQ/viewform'
    : 'https://docs.google.com/forms/d/e/1FAIpQLSe53HZRJiAZM7VCZUejMTgQ34RxfV9K4Kn-20uHqFiUQKvKaQ/viewform';
  
  const entryField = isPrinter ? 'entry.1109208984' : 'entry.1372464946';
  
  return `${baseUrl}?${entryField}=${encodeURIComponent(locationNumber)}`;
}

// 別の実装方法：URL Shortenerサービスとの連携
function createShortUrl(longUrl) {
  // Google URL Shortenerは廃止されたため、別のサービスを使用
  // または、事前に短縮URLを生成してマッピングを保存
  const urlMappings = {
    'OSAKA_server_ThinkPad_ABC123_001': 'https://short.link/abc123',
    // ... 他のマッピング
  };
  
  return urlMappings[longUrl] || longUrl;
}