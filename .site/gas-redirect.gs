// Google Apps Script Web App
// これをGASプロジェクトにコピーして、Webアプリとしてデプロイしてください

function doGet(e) {
  // URLパラメータから拠点管理番号を取得
  const locationNumber = e.parameter.id || '';
  
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
  
  // HTMLテンプレートを作成
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>リダイレクト中...</title>
        <script>
          window.location.href = '${redirectUrl}';
        </script>
      </head>
      <body>
        <div style="text-align: center; padding: 50px; font-family: Arial, sans-serif;">
          <h2>代替機ステータス登録</h2>
          <p>拠点管理番号: ${locationNumber}</p>
          <p>フォームへ移動中...</p>
          <p>自動的に移動しない場合は、<a href="${redirectUrl}">こちらをクリック</a>してください。</p>
        </div>
      </body>
    </html>
  `);
  
  return htmlOutput;
}

function getCategoryFromLocationNumber(locationNumber) {
  if (!locationNumber) return 'desktop';
  
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