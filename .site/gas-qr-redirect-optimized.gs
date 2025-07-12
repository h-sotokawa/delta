// QRコード中間ページ - 最適化版
// 最速リダイレクトを実現

function doGet(e) {
  const locationNumber = e.parameter.id || '';
  
  if (!locationNumber) {
    return HtmlService.createHtmlOutput(getErrorPage());
  }
  
  const category = getCategoryFromLocationNumber(locationNumber);
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
  
  const formType = (category === 'printer' || category === 'other') ? 'printer' : 'terminal';
  const form = forms[formType];
  const redirectUrl = `${form.formUrl}?${form.entryField}=${encodeURIComponent(locationNumber)}`;
  
  // 最小限のHTMLで最速リダイレクト
  return HtmlService.createHtmlOutput(
    '<script>location.replace("' + redirectUrl + '");</script>'
  ).setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getErrorPage() {
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>エラー</title>
<style>
body{font-family:sans-serif;text-align:center;padding:50px;background:#f5f5f5}
.error{background:white;padding:40px;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,0.1);max-width:400px;margin:0 auto}
h2{color:#dc2626}
</style>
</head>
<body>
<div class="error">
<h2>エラー</h2>
<p>拠点管理番号が指定されていません。</p>
<p>QRコードを正しく読み取ってください。</p>
</div>
</body>
</html>`;
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