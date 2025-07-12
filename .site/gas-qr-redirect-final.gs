// QRコード中間ページ - Google Apps Script Web App
// プロジェクト名: QRコードリダイレクト

function doGet(e) {
  // URLパラメータから拠点管理番号を取得
  const locationNumber = e.parameter.id || '';
  
  // エラーページ
  if (!locationNumber) {
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <title>エラー - 代替機管理システム</title>
          <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
          <style>
            * {
              box-sizing: border-box;
            }
            
            body {
              font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
              text-align: center;
              padding: 20px;
              background-color: #f5f5f5;
              margin: 0;
              min-height: 100vh;
              display: flex;
              align-items: center;
              justify-content: center;
            }
            
            .error-container {
              background-color: white;
              padding: 40px 25px;
              border-radius: 15px;
              box-shadow: 0 4px 20px rgba(0,0,0,0.1);
              max-width: 500px;
              width: 100%;
              margin: 0 auto;
            }
            
            .error-icon {
              font-size: 4rem;
              margin-bottom: 20px;
            }
            
            h2 {
              color: #dc2626;
              margin-bottom: 20px;
              font-size: 1.8rem;
              font-weight: 600;
            }
            
            .error-message {
              color: #374151;
              font-size: 1.1rem;
              line-height: 1.6;
              margin-bottom: 30px;
            }
            
            .suggestion {
              background-color: #fef3c7;
              border: 2px solid #fbbf24;
              border-radius: 10px;
              padding: 20px;
              margin-top: 25px;
            }
            
            .suggestion-title {
              color: #92400e;
              font-weight: 600;
              margin-bottom: 10px;
              font-size: 1.1rem;
            }
            
            .suggestion-list {
              text-align: left;
              margin: 0;
              padding-left: 20px;
              color: #78350f;
            }
            
            .suggestion-list li {
              margin: 8px 0;
              line-height: 1.5;
            }
            
            .retry-button {
              display: inline-block;
              margin-top: 30px;
              padding: 14px 28px;
              background-color: #3b82f6;
              color: white;
              text-decoration: none;
              border-radius: 50px;
              font-size: 1rem;
              font-weight: 500;
              transition: all 0.3s ease;
              box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
            }
            
            .retry-button:hover {
              background-color: #2563eb;
              transform: translateY(-2px);
              box-shadow: 0 6px 20px rgba(59, 130, 246, 0.4);
            }
            
            /* スマートフォン対応 */
            @media (max-width: 480px) {
              body {
                padding: 10px;
              }
              
              .error-container {
                padding: 30px 20px;
                border-radius: 12px;
              }
              
              .error-icon {
                font-size: 3rem;
              }
              
              h2 {
                font-size: 1.5rem;
              }
              
              .error-message {
                font-size: 1rem;
              }
              
              .suggestion {
                padding: 15px;
              }
              
              .suggestion-title {
                font-size: 1rem;
              }
              
              .suggestion-list {
                font-size: 0.9rem;
              }
              
              .retry-button {
                width: 100%;
                max-width: 250px;
              }
            }
          </style>
        </head>
        <body>
          <div class="error-container">
            <div class="error-icon">⚠️</div>
            <h2>エラーが発生しました</h2>
            
            <div class="error-message">
              <p>拠点管理番号が指定されていません。</p>
              <p>QRコードを正しく読み取ってください。</p>
            </div>
            
            <div class="suggestion">
              <div class="suggestion-title">確認事項：</div>
              <ul class="suggestion-list">
                <li>QRコードが完全に読み取れているか</li>
                <li>カメラが正常に動作しているか</li>
                <li>QRコードが破損していないか</li>
              </ul>
            </div>
            
            <a href="javascript:history.back()" class="retry-button">
              もう一度試す
            </a>
          </div>
        </body>
      </html>
    `);
  }
  
  // カテゴリを判別
  const category = getCategoryFromLocationNumber(locationNumber);
  const categoryNames = {
    'desktop': 'デスクトップ',
    'laptop': 'ノートパソコン',
    'server': 'サーバー',
    'printer': 'プリンタ',
    'other': 'その他'
  };
  const displayName = categoryNames[category] || '端末';
  
  // フォーム設定
  const forms = {
    terminal: {
      formUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSe53HZRJiAZM7VCZUejMTgQ34RxfV9K4Kn-20uHqFiUQKvKaQ/viewform',
      entryField: 'entry.1372464946'
    },
    printer: {
      formUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSfkxqPsHQpeKhWTBO6y3dUyvHoUKw365Wrdzsi1NNMbJ6C5mQ/viewform',
      entryField: 'entry.1109208984'
    }
  };
  
  // カテゴリに応じたフォームを選択
  const formType = (category === 'printer' || category === 'other') ? 'printer' : 'terminal';
  const form = forms[formType];
  
  // リダイレクトURLを構築
  const redirectUrl = `${form.formUrl}?${form.entryField}=${encodeURIComponent(locationNumber)}`;
  
  // 即座にリダイレクト（meta refreshを0秒に設定）
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <meta http-equiv="refresh" content="0;url=${redirectUrl}">
        <title>リダイレクト中...</title>
      </head>
      <body>
        <script>
          // JavaScriptでも即座にリダイレクト（念のため）
          window.location.href = '${redirectUrl}';
        </script>
      </body>
    </html>
  `);
  
  return htmlOutput;
}

function getCategoryFromLocationNumber(locationNumber) {
  if (!locationNumber) return 'desktop';
  
  // 拠点管理番号の形式: 拠点_カテゴリ_モデル_製造番号_連番
  const parts = locationNumber.split('_');
  if (parts.length >= 2) {
    const category = parts[1].toLowerCase();
    const validCategories = ['desktop', 'laptop', 'server', 'printer', 'other'];
    if (validCategories.includes(category)) {
      return category;
    }
  }
  return 'desktop';
}