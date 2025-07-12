// QRコード即座リダイレクト版 - Google Apps Script Web App
// 中間ページを表示せず、即座にGoogle Formsへリダイレクト

function doGet(e) {
  // URLパラメータから拠点管理番号を取得
  const locationNumber = e.parameter.id || '';
  
  // エラーページ（拠点管理番号がない場合のみ表示）
  if (!locationNumber) {
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <title>エラー - 代替機管理システム</title>
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <style>
            body {
              font-family: Arial, sans-serif;
              text-align: center;
              padding: 50px;
              background-color: #f5f5f5;
            }
            .error-container {
              background-color: white;
              padding: 30px;
              border-radius: 10px;
              box-shadow: 0 2px 10px rgba(0,0,0,0.1);
              max-width: 400px;
              margin: 0 auto;
            }
            h2 { color: #dc2626; }
          </style>
        </head>
        <body>
          <div class="error-container">
            <h2>エラー</h2>
            <p>拠点管理番号が指定されていません。</p>
            <p>QRコードを正しく読み取ってください。</p>
          </div>
        </body>
      </html>
    `);
  }
  
  // カテゴリを判別
  const category = getCategoryFromLocationNumber(locationNumber);
  
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
  return HtmlService.createHtmlOutput(`
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