// QRコード中間ページ - モバイルフレンドリー版
// プロジェクト名: QRコードリダイレクト（モバイル最適化）

function doGet(e) {
  // URLパラメータから拠点管理番号を取得
  const locationNumber = e.parameter.id || '';
  
  // エラーページ
  if (!locationNumber) {
    return HtmlService.createHtmlOutput(getErrorPage());
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
  
  // モードを判定（iframe表示 or リダイレクト）
  const useIframe = e.parameter.mode === 'iframe';
  
  if (useIframe) {
    // iframe版の中間ページ
    return HtmlService.createHtmlOutput(getIframePage(redirectUrl, locationNumber, category, displayName))
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    // 通常のリダイレクト版
    return HtmlService.createHtmlOutput(getRedirectPage(redirectUrl, locationNumber, category, displayName));
  }
}

function getIframePage(formUrl, locationNumber, category, displayName) {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>代替機ステータス登録</title>
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
        <style>
          * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
          }
          
          body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background-color: #f5f5f5;
            overflow: hidden;
          }
          
          .header {
            background-color: white;
            padding: 15px 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 100;
          }
          
          .header-content {
            max-width: 1200px;
            margin: 0 auto;
            display: flex;
            align-items: center;
            justify-content: space-between;
          }
          
          .header-title {
            font-size: 1.2rem;
            font-weight: 600;
            color: #1e40af;
          }
          
          .location-info {
            background-color: #f0f7ff;
            padding: 8px 16px;
            border-radius: 20px;
            font-size: 0.9rem;
            color: #1e40af;
            display: flex;
            align-items: center;
            gap: 8px;
          }
          
          .device-icon {
            font-size: 1.2rem;
          }
          
          .iframe-container {
            position: fixed;
            top: 70px;
            left: 0;
            right: 0;
            bottom: 0;
            width: 100%;
            height: calc(100vh - 70px);
            background-color: white;
          }
          
          iframe {
            width: 100%;
            height: 100%;
            border: none;
          }
          
          .loading-overlay {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: rgba(255, 255, 255, 0.9);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 50;
          }
          
          .loading-content {
            text-align: center;
          }
          
          .loading-spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #1a73e8;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          .loading-text {
            color: #666;
            font-size: 1.1rem;
          }
          
          .open-new-tab {
            position: fixed;
            bottom: 20px;
            right: 20px;
            background-color: #1a73e8;
            color: white;
            padding: 12px 24px;
            border-radius: 50px;
            text-decoration: none;
            font-size: 0.9rem;
            font-weight: 500;
            box-shadow: 0 4px 15px rgba(26, 115, 232, 0.3);
            display: flex;
            align-items: center;
            gap: 8px;
            z-index: 100;
          }
          
          .open-new-tab:hover {
            background-color: #1557b0;
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(26, 115, 232, 0.4);
          }
          
          /* モバイル対応 */
          @media (max-width: 768px) {
            .header {
              padding: 12px 15px;
            }
            
            .header-content {
              flex-direction: column;
              gap: 10px;
            }
            
            .header-title {
              font-size: 1.1rem;
            }
            
            .location-info {
              font-size: 0.85rem;
              padding: 6px 12px;
            }
            
            .iframe-container {
              top: 90px;
              height: calc(100vh - 90px);
            }
            
            .open-new-tab {
              bottom: 15px;
              right: 15px;
              padding: 10px 20px;
              font-size: 0.85rem;
            }
          }
          
          /* アイコンスタイル */
          .icon-desktop::before { content: "🖥️"; }
          .icon-laptop::before { content: "💻"; }
          .icon-server::before { content: "🖲️"; }
          .icon-printer::before { content: "🖨️"; }
          .icon-other::before { content: "📱"; }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="header-content">
            <h1 class="header-title">代替機ステータス登録</h1>
            <div class="location-info">
              <span class="device-icon icon-${category}"></span>
              <span>${locationNumber}</span>
            </div>
          </div>
        </div>
        
        <div class="iframe-container">
          <div class="loading-overlay" id="loadingOverlay">
            <div class="loading-content">
              <div class="loading-spinner"></div>
              <div class="loading-text">フォームを読み込んでいます...</div>
            </div>
          </div>
          <iframe 
            id="formFrame" 
            src="${formUrl}"
            onload="hideLoading()"
            allow="camera; microphone"
          ></iframe>
        </div>
        
        <a href="${formUrl}" target="_blank" class="open-new-tab">
          新しいタブで開く 
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path>
            <polyline points="15 3 21 3 21 9"></polyline>
            <line x1="10" y1="14" x2="21" y2="3"></line>
          </svg>
        </a>
        
        <script>
          function hideLoading() {
            const overlay = document.getElementById('loadingOverlay');
            if (overlay) {
              overlay.style.display = 'none';
            }
          }
          
          // iframeのエラーハンドリング
          window.addEventListener('message', function(e) {
            if (e.data && e.data.type === 'form-submitted') {
              // フォーム送信完了時の処理
              alert('フォームが送信されました');
            }
          });
        </script>
      </body>
    </html>
  `;
}

function getRedirectPage(redirectUrl, locationNumber, category, displayName) {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>代替機ステータス登録</title>
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
        <meta http-equiv="refresh" content="3;url=${redirectUrl}">
        <style>
          * {
            box-sizing: border-box;
            -webkit-tap-highlight-color: transparent;
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
          
          .container {
            background-color: white;
            padding: 30px 20px;
            border-radius: 15px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            max-width: 500px;
            width: 100%;
            margin: 0 auto;
          }
          
          h2 {
            color: #333;
            margin-bottom: 25px;
            font-size: 1.8rem;
            font-weight: 600;
          }
          
          .info-box {
            background-color: #f0f7ff;
            border: 2px solid #d0e3ff;
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
          }
          
          .info-label {
            font-size: 0.9rem;
            color: #666;
            margin-bottom: 8px;
          }
          
          .location-number {
            font-size: 1.3rem;
            font-weight: bold;
            color: #1a73e8;
            word-break: break-all;
            line-height: 1.4;
            margin-bottom: 12px;
          }
          
          .device-type {
            color: #666;
            font-size: 1rem;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
          }
          
          .device-icon {
            font-size: 1.2rem;
          }
          
          .loading {
            margin: 30px 0;
            color: #666;
          }
          
          .loading-spinner {
            display: inline-block;
            width: 40px;
            height: 40px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #1a73e8;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin-bottom: 15px;
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          .loading-text {
            font-size: 1.1rem;
            line-height: 1.5;
          }
          
          .countdown {
            font-size: 2rem;
            font-weight: bold;
            color: #1a73e8;
            margin: 10px 0;
          }
          
          .manual-link {
            display: inline-block;
            margin-top: 20px;
            padding: 16px 32px;
            background-color: #1a73e8;
            color: white;
            text-decoration: none;
            border-radius: 50px;
            font-size: 1.1rem;
            font-weight: 500;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(26, 115, 232, 0.3);
            touch-action: manipulation;
          }
          
          .manual-link:hover {
            background-color: #1557b0;
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(26, 115, 232, 0.4);
          }
          
          .manual-link:active {
            transform: translateY(0);
            box-shadow: 0 2px 10px rgba(26, 115, 232, 0.3);
          }
          
          .tips {
            background-color: #fef3c7;
            border: 2px solid #fbbf24;
            border-radius: 10px;
            padding: 15px;
            margin-top: 25px;
            text-align: left;
          }
          
          .tips-title {
            color: #92400e;
            font-weight: 600;
            margin-bottom: 10px;
            font-size: 1rem;
          }
          
          .tips-content {
            color: #78350f;
            font-size: 0.9rem;
            line-height: 1.5;
          }
          
          /* スマートフォン対応 */
          @media (max-width: 480px) {
            body {
              padding: 10px;
            }
            
            .container {
              padding: 25px 15px;
              border-radius: 12px;
            }
            
            h2 {
              font-size: 1.5rem;
              margin-bottom: 20px;
            }
            
            .info-box {
              padding: 15px;
            }
            
            .location-number {
              font-size: 1.1rem;
            }
            
            .device-type {
              font-size: 0.9rem;
            }
            
            .loading-text {
              font-size: 1rem;
            }
            
            .manual-link {
              padding: 14px 28px;
              font-size: 1rem;
              width: 100%;
              max-width: 280px;
            }
            
            .tips {
              padding: 12px;
            }
            
            .tips-title {
              font-size: 0.95rem;
            }
            
            .tips-content {
              font-size: 0.85rem;
            }
          }
          
          /* タブレット対応 */
          @media (min-width: 481px) and (max-width: 768px) {
            .container {
              padding: 35px 25px;
            }
            
            h2 {
              font-size: 1.7rem;
            }
          }
          
          /* アイコンスタイル */
          .icon-desktop::before { content: "🖥️"; }
          .icon-laptop::before { content: "💻"; }
          .icon-server::before { content: "🖲️"; }
          .icon-printer::before { content: "🖨️"; }
          .icon-other::before { content: "📱"; }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>代替機ステータス登録</h2>
          
          <div class="info-box">
            <div class="info-label">拠点管理番号：</div>
            <div class="location-number">${locationNumber}</div>
            <div class="device-type">
              <span class="device-icon icon-${category}"></span>
              <span>デバイスタイプ: ${displayName}</span>
            </div>
          </div>
          
          <div class="loading">
            <div class="loading-spinner"></div>
            <div class="loading-text">
              フォームへ移動します...
            </div>
            <div class="countdown" id="countdown">3</div>
            <div style="font-size: 0.9rem; color: #999;">
              秒後に自動的に移動します
            </div>
          </div>
          
          <a href="${redirectUrl}" class="manual-link">
            今すぐフォームへ移動 →
          </a>
          
          <div class="tips">
            <div class="tips-title">💡 ヒント</div>
            <div class="tips-content">
              スマートフォンでフォームが見づらい場合は、画面を横向きにするか、ピンチアウト（2本指で広げる）で拡大してご利用ください。
            </div>
          </div>
        </div>
        
        <script>
          // カウントダウン表示
          let count = 3;
          const countdownEl = document.getElementById('countdown');
          
          const countdownInterval = setInterval(function() {
            count--;
            if (count > 0) {
              countdownEl.textContent = count;
            } else {
              countdownEl.textContent = '0';
              clearInterval(countdownInterval);
            }
          }, 1000);
          
          // 自動リダイレクト
          setTimeout(function() {
            window.location.href = '${redirectUrl}';
          }, 3000);
          
          // タッチデバイスでのフィードバック改善
          const link = document.querySelector('.manual-link');
          link.addEventListener('touchstart', function() {
            this.style.transform = 'scale(0.95)';
          });
          link.addEventListener('touchend', function() {
            this.style.transform = 'scale(1)';
          });
        </script>
      </body>
    </html>
  `;
}

function getErrorPage() {
  return `
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
  `;
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