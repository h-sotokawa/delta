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
          <title>エラー</title>
        </head>
        <body style="text-align: center; padding: 50px; font-family: Arial, sans-serif;">
          <h2 style="color: red;">エラー</h2>
          <p>拠点管理番号が指定されていません。</p>
          <p>QRコードを正しく読み取ってください。</p>
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
  
  // HTMLテンプレートを作成（自動リダイレクト付き）
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>代替機ステータス登録</title>
        <meta http-equiv="refresh" content="2;url=${redirectUrl}">
        <style>
          body {
            font-family: Arial, sans-serif;
            text-align: center;
            padding: 50px;
            background-color: #f5f5f5;
          }
          .container {
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            max-width: 500px;
            margin: 0 auto;
          }
          h2 {
            color: #333;
            margin-bottom: 20px;
          }
          .info-box {
            background-color: #f0f7ff;
            border: 1px solid #d0e3ff;
            padding: 20px;
            border-radius: 5px;
            margin: 20px 0;
          }
          .location-number {
            font-size: 1.2em;
            font-weight: bold;
            color: #1a73e8;
            word-break: break-all;
          }
          .device-type {
            color: #666;
            margin-top: 10px;
          }
          .loading {
            margin: 20px 0;
            color: #666;
          }
          .manual-link {
            display: inline-block;
            margin-top: 20px;
            padding: 10px 20px;
            background-color: #1a73e8;
            color: white;
            text-decoration: none;
            border-radius: 5px;
          }
          .manual-link:hover {
            background-color: #1557b0;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>代替機ステータス登録</h2>
          
          <div class="info-box">
            <div>拠点管理番号：</div>
            <div class="location-number">${locationNumber}</div>
            <div class="device-type">デバイスタイプ: ${displayName}</div>
          </div>
          
          <div class="loading">
            <p>フォームへ移動します...</p>
            <p>2秒後に自動的に移動します</p>
          </div>
          
          <a href="${redirectUrl}" class="manual-link">
            今すぐフォームへ移動
          </a>
        </div>
        
        <script>
          // 念のためJavaScriptでもリダイレクト
          setTimeout(function() {
            window.location.href = '${redirectUrl}';
          }, 2000);
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