# テスト戦略書

## 1. テスト戦略概要

### 1.1 テストの目的
- **機能品質**: 要求仕様通りの機能実現確認
- **非機能品質**: パフォーマンス・セキュリティ・使いやすさの確認
- **回帰品質**: 変更による既存機能への影響確認
- **リリース品質**: 本番環境での安定動作確認

### 1.2 テスト方針
- **段階的テスト**: 単体→統合→システム→受入の順次実施
- **自動化優先**: 繰り返し実行可能な自動テストの構築
- **早期発見**: 開発工程での品質問題の早期検出
- **継続改善**: テスト結果に基づく品質向上

### 1.3 品質基準
- **機能テスト**: 100% 通過
- **パフォーマンステスト**: 要件定義の性能基準クリア
- **セキュリティテスト**: 脆弱性検査クリア
- **ユーザビリティテスト**: ユーザー満足度80%以上

## 2. テストレベル定義

### 2.1 単体テスト（Unit Test）

#### 2.1.1 対象範囲
- **Google Apps Script関数**: 各バックエンド関数の個別テスト
- **JavaScript関数**: フロントエンド関数の個別テスト
- **バリデーション関数**: 入力検証ロジックのテスト
- **ユーティリティ関数**: 共通処理関数のテスト

#### 2.1.2 テスト観点
- **正常系**: 期待される入力に対する正しい出力
- **異常系**: 不正入力に対する適切なエラーハンドリング
- **境界値**: 限界値・境界値での動作確認
- **性能**: 個別関数の処理時間・メモリ使用量

#### 2.1.3 実装方法
```javascript
// Google Apps Script単体テスト例
function testGenerateLocationNumber() {
  // テストデータ準備
  const testCases = [
    {
      input: { location: 'OSA', category: 'SV', model: 'ThinkPad', serial: 'ABC123', seq: 1 },
      expected: 'OSA_SV_ThinkPad_ABC123_001'
    },
    {
      input: { location: '', category: 'SV', model: 'ThinkPad', serial: 'ABC123', seq: 1 },
      expected: null  // エラーケース
    }
  ];
  
  // テスト実行
  testCases.forEach((testCase, index) => {
    try {
      const result = generateLocationNumber(testCase.input);
      
      if (testCase.expected === null) {
        // エラーが期待される場合
        console.log(`Test ${index + 1}: Expected error - ${result ? 'FAILED' : 'PASSED'}`);
      } else {
        // 正常結果が期待される場合
        const passed = result === testCase.expected;
        console.log(`Test ${index + 1}: ${passed ? 'PASSED' : 'FAILED'} - Expected: ${testCase.expected}, Got: ${result}`);
      }
    } catch (error) {
      if (testCase.expected === null) {
        console.log(`Test ${index + 1}: PASSED - Expected error: ${error.message}`);
      } else {
        console.log(`Test ${index + 1}: FAILED - Unexpected error: ${error.message}`);
      }
    }
  });
}

// データ種類マスタ単体テスト例
function testDataTypeMaster() {
  // テストデータ準備
  const testCases = [
    {
      name: "正常なデータ種類追加",
      input: { 
        dataTypeId: 'USER_INFO', 
        dataTypeName: 'ユーザー情報', 
        description: 'ユーザーに関する基本情報',
        displayOrder: 1,
        status: 'active'
      },
      expected: { success: true }
    },
    {
      name: "重複IDでのデータ種類追加",
      input: { 
        dataTypeId: 'USER_INFO', 
        dataTypeName: '重複ユーザー情報', 
        description: '重複テスト',
        displayOrder: 2,
        status: 'active'
      },
      expected: { success: false, error: 'データ種類IDが既に存在します' }
    },
    {
      name: "必須項目欠落",
      input: { 
        dataTypeId: '', 
        dataTypeName: 'テストデータ', 
        description: 'テスト',
        displayOrder: 3,
        status: 'active'
      },
      expected: { success: false, error: 'データ種類IDは必須です' }
    },
    {
      name: "無効なステータス",
      input: { 
        dataTypeId: 'TEST_DATA', 
        dataTypeName: 'テストデータ', 
        description: 'テスト',
        displayOrder: 4,
        status: 'invalid'
      },
      expected: { success: false, error: 'ステータスはactive/inactiveのみ許可されます' }
    }
  ];
  
  // テスト実行
  testCases.forEach((testCase) => {
    console.log(`実行中: ${testCase.name}`);
    try {
      const result = addDataTypeMaster(testCase.input);
      
      if (testCase.expected.success) {
        console.log(result.success ? '✅ PASSED' : `❌ FAILED - ${result.error}`);
      } else {
        console.log(result.error === testCase.expected.error ? '✅ PASSED' : `❌ FAILED - Expected: ${testCase.expected.error}, Got: ${result.error}`);
      }
    } catch (error) {
      console.log(`❌ FAILED - Unexpected error: ${error.message}`);
    }
  });
}

// データ種類表示順序テスト
function testDataTypeDisplayOrder() {
  const testCases = [
    {
      name: "表示順序による並び替え",
      input: [
        { dataTypeId: 'TYPE_3', displayOrder: 3, status: 'active' },
        { dataTypeId: 'TYPE_1', displayOrder: 1, status: 'active' },
        { dataTypeId: 'TYPE_2', displayOrder: 2, status: 'active' }
      ],
      expected: ['TYPE_1', 'TYPE_2', 'TYPE_3']
    },
    {
      name: "同一表示順序の処理",
      input: [
        { dataTypeId: 'TYPE_A', displayOrder: 1, status: 'active' },
        { dataTypeId: 'TYPE_B', displayOrder: 1, status: 'active' },
        { dataTypeId: 'TYPE_C', displayOrder: 2, status: 'active' }
      ],
      expectedOrder: [1, 1, 2]  // 順序値の確認
    }
  ];
  
  testCases.forEach((testCase) => {
    console.log(`実行中: ${testCase.name}`);
    const sorted = sortDataTypesByDisplayOrder(testCase.input);
    
    if (testCase.expected) {
      const actualOrder = sorted.map(item => item.dataTypeId);
      const passed = JSON.stringify(actualOrder) === JSON.stringify(testCase.expected);
      console.log(passed ? '✅ PASSED' : `❌ FAILED - Expected: ${testCase.expected}, Got: ${actualOrder}`);
    } else {
      const actualOrder = sorted.map(item => item.displayOrder);
      console.log(`結果: ${actualOrder}`);
    }
  });
}

// データ種類バリデーションテスト
function testDataTypeValidation() {
  const validationTests = [
    {
      name: "データ種類ID形式チェック",
      testCases: [
        { input: 'USER_INFO', expected: true },
        { input: 'USER-INFO', expected: false },
        { input: 'user_info', expected: false },
        { input: 'USER_INFO_123', expected: true },
        { input: '123_USER', expected: false },
        { input: '', expected: false }
      ]
    },
    {
      name: "データ種類名長さチェック",
      testCases: [
        { input: 'あ', expected: true },  // 最小長
        { input: 'あ'.repeat(50), expected: true },  // 最大長
        { input: 'あ'.repeat(51), expected: false },  // 超過
        { input: '', expected: false }
      ]
    },
    {
      name: "説明文長さチェック",
      testCases: [
        { input: '', expected: true },  // 空文字OK
        { input: 'あ'.repeat(200), expected: true },  // 最大長
        { input: 'あ'.repeat(201), expected: false }  // 超過
      ]
    }
  ];
  
  validationTests.forEach((test) => {
    console.log(`\n${test.name}:`);
    test.testCases.forEach((testCase, index) => {
      const result = validateDataTypeField(test.name, testCase.input);
      const passed = result === testCase.expected;
      console.log(`  Test ${index + 1}: ${passed ? '✅' : '❌'} Input: "${testCase.input}" Expected: ${testCase.expected}`);
    });
  });
}
```

### 2.2 統合テスト（Integration Test）

#### 2.2.1 対象範囲
- **フロントエンド-バックエンド連携**: `google.script.run`を使用したAPI呼び出し
- **バックエンド-Google Sheets連携**: Sheets APIを使用したデータ操作
- **マスタデータ連携**: 複数マスタ間の整合性確認
- **外部サービス連携**: Google Forms・QRコード生成との連携

#### 2.2.2 テストシナリオ
```javascript
// 統合テストシナリオ例: URL生成フロー
async function integrationTestUrlGeneration() {
  const testScenarios = [
    {
      name: "正常なURL生成フロー",
      steps: [
        // 1. 拠点マスタからデータ取得
        () => getLocationMaster(),
        // 2. 機種マスタからデータ取得
        () => getModelsByCategory('server'),
        // 3. URL生成
        (data) => generateCommonFormUrl('OSA_SV_ThinkPad_ABC123_001', 'server'),
        // 4. マスタ保存
        (result) => saveUrlToTerminalMaster('OSA_SV_ThinkPad_ABC123_001', result.url)
      ]
    }
  ];
  
  for (const scenario of testScenarios) {
    console.log(`実行中: ${scenario.name}`);
    
    try {
      let previousResult = null;
      for (const step of scenario.steps) {
        previousResult = await step(previousResult);
        if (!previousResult.success) {
          throw new Error(`Step failed: ${previousResult.error}`);
        }
      }
      console.log(`✅ ${scenario.name}: PASSED`);
    } catch (error) {
      console.log(`❌ ${scenario.name}: FAILED - ${error.message}`);
    }
  }
}

// 統合テストシナリオ例: データ種類マスタフロー
async function integrationTestDataTypeMaster() {
  const testScenarios = [
    {
      name: "データ種類マスタ完全フロー",
      steps: [
        // 1. 新規データ種類追加
        () => addDataTypeMaster({
          dataTypeId: 'TEST_TYPE',
          dataTypeName: 'テストデータ種類',
          description: '統合テスト用',
          displayOrder: 99,
          status: 'active'
        }),
        // 2. データ種類一覧取得（アクティブのみ）
        () => getActiveDataTypes(),
        // 3. データ種類更新
        () => updateDataTypeMaster('TEST_TYPE', {
          dataTypeName: '更新済みテストデータ',
          displayOrder: 50
        }),
        // 4. スプレッドシートビューアーで確認
        () => getSpreadsheetData('データ種類マスタ'),
        // 5. データ種類無効化
        () => updateDataTypeMaster('TEST_TYPE', { status: 'inactive' }),
        // 6. アクティブデータ確認
        () => getActiveDataTypes()
      ]
    },
    {
      name: "データ種類とスプレッドシート連携",
      steps: [
        // 1. データ種類マスタ取得
        () => getDataTypeMaster(),
        // 2. スプレッドシートビューアーでフィルタ
        (dataTypes) => filterSpreadsheetByDataType(dataTypes[0].dataTypeId),
        // 3. フィルタ結果の検証
        (filtered) => validateFilteredData(filtered),
        // 4. 表示順序でソート
        () => getDataTypesOrderedByDisplay()
      ]
    }
  ];
  
  for (const scenario of testScenarios) {
    console.log(`実行中: ${scenario.name}`);
    
    try {
      let previousResult = null;
      for (const step of scenario.steps) {
        previousResult = await step(previousResult);
        if (!previousResult.success) {
          throw new Error(`Step failed: ${previousResult.error}`);
        }
      }
      console.log(`✅ ${scenario.name}: PASSED`);
    } catch (error) {
      console.log(`❌ ${scenario.name}: FAILED - ${error.message}`);
    }
  }
}

// データ種類マスタとGoogle Sheets連携テスト
function integrationTestDataTypeSheets() {
  const tests = [
    {
      name: "シート作成・初期化テスト",
      test: () => {
        // データ種類マスタシート作成
        const sheet = createDataTypeMasterSheet();
        // ヘッダー検証
        const headers = sheet.getRange(1, 1, 1, 7).getValues()[0];
        const expectedHeaders = ['データ種類ID', 'データ種類名', '説明', '表示順序', 'ステータス', '作成日', '更新日'];
        return JSON.stringify(headers) === JSON.stringify(expectedHeaders);
      }
    },
    {
      name: "データ保存・取得テスト",
      test: () => {
        // データ保存
        const testData = {
          dataTypeId: 'INT_TEST',
          dataTypeName: '統合テスト',
          description: 'シート連携テスト',
          displayOrder: 1,
          status: 'active'
        };
        const saveResult = saveDataTypeToSheet(testData);
        
        // データ取得
        const retrievedData = getDataTypeFromSheet('INT_TEST');
        
        // 検証
        return retrievedData.dataTypeId === testData.dataTypeId &&
               retrievedData.dataTypeName === testData.dataTypeName;
      }
    }
  ];
  
  tests.forEach(test => {
    console.log(`実行中: ${test.name}`);
    try {
      const result = test.test();
      console.log(result ? '✅ PASSED' : '❌ FAILED');
    } catch (error) {
      console.log(`❌ FAILED - ${error.message}`);
    }
  });
}
```

### 2.3 システムテスト（System Test）

#### 2.3.1 機能テスト
- **ダッシュボード機能**: 全機能カードの動作確認
- **スプレッドシートビューアー**: データ表示・検索・フィルタ機能
- **URL生成機能**: 拠点管理番号生成・URL作成・保存
- **マスタ管理機能**: CRUD操作・バリデーション
- **設定機能**: 設定変更・保存・検証

#### 2.3.2 テストケース例

| テストID | 機能 | テストケース | 期待結果 |
|----------|------|-------------|----------|
| SYS-001 | URL生成 | 全項目正常入力 | URL正常生成・マスタ保存 |
| SYS-002 | URL生成 | 拠点未選択 | エラーメッセージ表示 |
| SYS-003 | URL生成 | 製造番号不正形式 | バリデーションエラー |
| SYS-004 | マスタ管理 | 拠点新規追加 | データ正常保存・一覧更新 |
| SYS-005 | マスタ管理 | 重複拠点ID追加 | 重複エラー表示 |
| SYS-006 | データ種類マスタ | 新規データ種類追加 | データ正常保存・作成日時自動設定 |
| SYS-007 | データ種類マスタ | 重複データ種類ID追加 | 重複エラーメッセージ表示 |
| SYS-008 | データ種類マスタ | データ種類情報更新 | 更新成功・更新日時自動更新 |
| SYS-009 | データ種類マスタ | 無効なデータ種類ID形式 | バリデーションエラー表示 |
| SYS-010 | データ種類マスタ | アクティブデータ種類のみ表示 | inactiveステータスのデータ非表示 |
| SYS-011 | データ種類マスタ | 表示順序による並び替え | displayOrder昇順で表示 |
| SYS-012 | データ種類マスタ | データ種類削除（無効化） | ステータスをinactiveに変更 |
| SYS-013 | スプレッドシートビューアー | データ種類でフィルタ | 選択データ種類のみ表示 |
| SYS-014 | スプレッドシートビューアー | データ種類マスタ表示 | 全データ種類情報正常表示 |
| SYS-015 | データ種類マスタ | 必須項目未入力 | エラーメッセージ表示・保存失敗 |
| SYS-016 | システム設定 | debugMode有効時のURL設定変更 | 正常に保存・検証成功 |
| SYS-017 | システム設定 | debugMode無効時のURL設定変更 | エラー表示・変更不可 |
| SYS-018 | システム設定 | エラー通知メール設定（有効） | 複数メール正常保存 |
| SYS-019 | システム設定 | エラー通知メール設定（無効） | 不正形式エラー表示 |
| SYS-020 | システム設定 | アラート通知メール設定 | 正常保存・テスト送信成功 |
| SYS-021 | システム設定 | debugMode無効時の通知設定変更 | 読み取り専用・変更不可 |
| SYS-022 | システム設定 | 複数メールアドレス検証 | 有効/無効アドレス分離表示 |
| SYS-023 | システム設定 | テスト通知送信 | 指定アドレスへ送信成功 |

### 2.4 受入テスト（Acceptance Test）

#### 2.4.1 ユーザーシナリオテスト
```
シナリオ1: 新しい代替機の管理番号生成
1. ダッシュボードからURL生成画面へ遷移
2. 大阪営業所を選択
3. サーバーカテゴリを選択
4. ThinkPad X1 Carbonを選択
5. 製造番号ABC123456を入力
6. 連番001を確認
7. URL生成・保存ボタンをクリック
8. 生成されたURLを確認
9. 端末マスタに保存されたことを確認

期待結果:
- OSA_SV_ThinkPadX1Carbon_ABC123456_001が生成される
- Google FormsのURLに事前入力パラメータが含まれる
- 端末マスタに新しい行が追加される

シナリオ2: 新しいデータ種類の追加と管理
1. ダッシュボードからデータ種類マスタ管理画面へ遷移
2. 「新規追加」ボタンをクリック
3. データ種類ID「USER_INFO」を入力
4. データ種類名「ユーザー情報」を入力
5. 説明「ユーザーに関する基本情報」を入力
6. 表示順序「1」を入力
7. ステータス「active」を選択
8. 「保存」ボタンをクリック
9. 一覧画面に新しいデータ種類が表示されることを確認
10. スプレッドシートビューアーでデータ種類フィルタに表示されることを確認

期待結果:
- データ種類が正常に保存される
- 作成日時が自動設定される
- アクティブなデータ種類として一覧に表示される
- スプレッドシートビューアーのフィルタオプションに追加される

シナリオ3: データ種類の更新と表示順序変更
1. データ種類マスタ一覧画面を開く
2. 既存のデータ種類「USER_INFO」の編集ボタンをクリック
3. データ種類名を「ユーザー基本情報」に変更
4. 表示順序を「5」に変更
5. 「更新」ボタンをクリック
6. 一覧画面で表示順序が変更されていることを確認
7. 更新日時が自動更新されていることを確認

期待結果:
- データ種類情報が正常に更新される
- 表示順序に従って一覧が並び替えられる
- 更新日時が現在時刻に更新される

シナリオ4: データ種類の無効化とフィルタリング
1. データ種類マスタ一覧画面を開く
2. 使用しなくなったデータ種類の編集ボタンをクリック
3. ステータスを「inactive」に変更
4. 「更新」ボタンをクリック
5. アクティブデータのみ表示フィルタを適用
6. 無効化したデータが非表示になることを確認
7. スプレッドシートビューアーのフィルタから消えることを確認

期待結果:
- ステータスがinactiveに更新される
- アクティブフィルタ適用時に非表示になる
- スプレッドシートビューアーのフィルタオプションから除外される
```

## 3. テストタイプ別戦略

### 3.1 パフォーマンステスト

#### 3.1.1 負荷テスト
```javascript
// パフォーマンステスト実装例
function performanceTestUrlGeneration() {
  const testCases = 100;  // 100回実行
  const results = [];
  
  for (let i = 0; i < testCases; i++) {
    const startTime = Date.now();
    
    try {
      const result = generateCommonFormUrl(
        `OSA_SV_ThinkPad_ABC${i.toString().padStart(3, '0')}_001`,
        'server'
      );
      
      const endTime = Date.now();
      const responseTime = endTime - startTime;
      
      results.push({
        testNumber: i + 1,
        success: result.success,
        responseTime: responseTime
      });
      
    } catch (error) {
      results.push({
        testNumber: i + 1,
        success: false,
        error: error.message,
        responseTime: null
      });
    }
  }
  
  // 結果分析
  const successCount = results.filter(r => r.success).length;
  const responseTimes = results.filter(r => r.responseTime).map(r => r.responseTime);
  const avgResponseTime = responseTimes.reduce((a, b) => a + b, 0) / responseTimes.length;
  const maxResponseTime = Math.max(...responseTimes);
  
  console.log('パフォーマンステスト結果:');
  console.log(`成功率: ${(successCount / testCases * 100).toFixed(2)}%`);
  console.log(`平均応答時間: ${avgResponseTime.toFixed(2)}ms`);
  console.log(`最大応答時間: ${maxResponseTime}ms`);
}
```

#### 3.1.2 性能基準
- **URL生成**: 平均1秒以内、最大2秒以内
- **データ取得**: 平均3秒以内、最大5秒以内
- **マスタ更新**: 平均2秒以内、最大4秒以内

### 3.2 セキュリティテスト

#### 3.2.1 認証・認可テスト
- **未認証アクセス**: 認証なしでのAPI呼び出し拒否確認
- **権限外アクセス**: 管理者機能への一般ユーザーアクセス拒否
- **セッション管理**: セッション有効期限・無効化の確認

#### 3.2.2 入力検証テスト
```javascript
// セキュリティテスト例: SQLインジェクション対策
function securityTestSqlInjection() {
  const maliciousInputs = [
    "'; DROP TABLE users; --",
    "<script>alert('XSS')</script>",
    "../../etc/passwd",
    "../../../../windows/system32",
    "' OR '1'='1"
  ];
  
  maliciousInputs.forEach((input, index) => {
    try {
      const result = validateLocationNumber(input);
      
      if (result.valid) {
        console.log(`❌ Security Test ${index + 1}: FAILED - Malicious input accepted: ${input}`);
      } else {
        console.log(`✅ Security Test ${index + 1}: PASSED - Malicious input rejected: ${input}`);
      }
    } catch (error) {
      console.log(`✅ Security Test ${index + 1}: PASSED - Exception thrown for: ${input}`);
    }
  });
}
```

### 3.3 ユーザビリティテスト

#### 3.3.1 タスク完了テスト
```
タスク1: 新規代替機の管理番号を生成して保存する
- 想定時間: 3分以内
- 成功基準: エラーなしでURL生成・保存完了
- 測定項目: 完了時間、エラー回数、ヘルプ参照回数

タスク2: 大阪営業所の端末一覧を確認する
- 想定時間: 1分以内
- 成功基準: 目的のデータを正しく表示
- 測定項目: 完了時間、検索操作回数

タスク3: 新しい機種をマスタに追加する
- 想定時間: 2分以内
- 成功基準: 機種情報の正常保存
- 測定項目: 完了時間、入力エラー回数
```

#### 3.3.2 満足度調査
- **System Usability Scale (SUS)**: 標準的なユーザビリティ評価
- **Net Promoter Score (NPS)**: システム推奨度調査
- **Custom Survey**: 機能別満足度調査

### 3.4 アクセシビリティテスト

#### 3.4.1 自動テスト
```javascript
// アクセシビリティ自動チェック（疑似コード）
function accessibilityTest() {
  const tests = [
    checkAltTextOnImages,
    checkHeadingStructure,
    checkColorContrast,
    checkKeyboardNavigation,
    checkAriaLabels,
    checkFocusIndicators
  ];
  
  tests.forEach(test => {
    const result = test();
    console.log(`${test.name}: ${result.passed ? 'PASSED' : 'FAILED'}`);
    if (!result.passed) {
      console.log(`Issues: ${result.issues.join(', ')}`);
    }
  });
}
```

#### 3.4.2 手動テスト
- **キーボード操作**: マウスなしでの全機能操作確認
- **スクリーンリーダー**: 音声読み上げでの操作確認
- **拡大表示**: 200%拡大での表示・操作確認

## 4. テスト環境・データ

### 4.1 テスト環境構成

#### 4.1.1 開発テスト環境
- **プラットフォーム**: Google Apps Script Test Deployment
- **データ**: テスト用Google Sheets
- **ユーザー**: 開発者個人アカウント
- **設定**: DEBUG_MODE = true

#### 4.1.2 統合テスト環境
- **プラットフォーム**: Google Apps Script Staging Deployment
- **データ**: 本番類似テストデータ
- **ユーザー**: テスト用組織アカウント
- **設定**: DEBUG_MODE = false

#### 4.1.3 本番テスト環境
- **プラットフォーム**: Google Apps Script Production Deployment
- **データ**: 本番データのコピー
- **ユーザー**: 実際の利用者アカウント
- **設定**: 本番設定

### 4.2 テストデータ設計

#### 4.2.1 基本テストデータ
```javascript
// テスト用拠点マスタ
const testLocations = [
  { locationId: 'OSA', locationName: '大阪営業所テスト', status: 'active' },
  { locationId: 'KOB', locationName: '神戸営業所テスト', status: 'active' },
  { locationId: 'HIM', locationName: '姫路営業所テスト', status: 'inactive' }
];

// テスト用機種マスタ
const testModels = [
  { modelName: 'TestThinkPad', manufacturer: 'TestLenovo', category: 'laptop' },
  { modelName: 'TestOptiPlex', manufacturer: 'TestDell', category: 'desktop' },
  { modelName: 'TestPrinter', manufacturer: 'TestCanon', category: 'printer' }
];

// テスト用データ種類マスタ
const testDataTypes = [
  { 
    dataTypeId: 'USER_INFO', 
    dataTypeName: 'ユーザー情報テスト', 
    description: 'ユーザーに関する基本情報のテストデータ',
    displayOrder: 1,
    status: 'active'
  },
  { 
    dataTypeId: 'DEVICE_INFO', 
    dataTypeName: 'デバイス情報テスト', 
    description: 'デバイスに関する詳細情報のテストデータ',
    displayOrder: 2,
    status: 'active'
  },
  { 
    dataTypeId: 'MAINTENANCE_LOG', 
    dataTypeName: 'メンテナンスログテスト', 
    description: 'メンテナンス履歴のテストデータ',
    displayOrder: 3,
    status: 'active'
  },
  { 
    dataTypeId: 'OLD_DATA', 
    dataTypeName: '旧データテスト', 
    description: '使用されなくなったテストデータ',
    displayOrder: 99,
    status: 'inactive'
  }
];
```

#### 4.2.2 境界値テストデータ
```javascript
// 境界値テストケース
const boundaryTestCases = [
  {
    name: '最小値テスト',
    locationNumber: 'A_A_A_A_001',  // 最短の有効な拠点管理番号
  },
  {
    name: '最大値テスト',
    locationNumber: 'ABC_DESKTOP_' + 'A'.repeat(50) + '_' + 'A'.repeat(20) + '_999'  // 最長の有効な拠点管理番号
  },
  {
    name: '空文字テスト',
    locationNumber: '',  // 無効ケース
  }
];

// データ種類マスタ境界値テストケース
const dataTypeBoundaryTests = [
  {
    name: 'データ種類ID最小値',
    dataType: {
      dataTypeId: 'A',  // 最短の有効なID
      dataTypeName: 'テスト',
      displayOrder: 1
    }
  },
  {
    name: 'データ種類ID最大値',
    dataType: {
      dataTypeId: 'A'.repeat(50) + '_' + 'B'.repeat(49),  // 最長の有効なID（100文字）
      dataTypeName: 'テスト',
      displayOrder: 1
    }
  },
  {
    name: 'データ種類名最大値',
    dataType: {
      dataTypeId: 'TEST_MAX_NAME',
      dataTypeName: 'あ'.repeat(50),  // 最大50文字
      displayOrder: 1
    }
  },
  {
    name: '説明文最大値',
    dataType: {
      dataTypeId: 'TEST_MAX_DESC',
      dataTypeName: 'テスト',
      description: 'あ'.repeat(200),  // 最大200文字
      displayOrder: 1
    }
  },
  {
    name: '表示順序最小値',
    dataType: {
      dataTypeId: 'TEST_MIN_ORDER',
      dataTypeName: 'テスト',
      displayOrder: 0  // 最小値
    }
  },
  {
    name: '表示順序最大値',
    dataType: {
      dataTypeId: 'TEST_MAX_ORDER',
      dataTypeName: 'テスト',
      displayOrder: 999999  // 最大値
    }
  }
];
```

### 4.3 テストデータ管理

#### 4.3.1 テストデータ作成
```javascript
function setupTestData() {
  // テスト用スプレッドシートの初期化
  const testSpreadsheetId = createTestSpreadsheet();
  
  // マスタデータの投入
  populateTestMasterData(testSpreadsheetId);
  
  // 拠点別データの投入
  populateTestLocationData(testSpreadsheetId);
  
  // データ種類マスタテストデータの投入
  populateTestDataTypes(testSpreadsheetId);
  
  console.log(`テストデータセットアップ完了: ${testSpreadsheetId}`);
}

function populateTestDataTypes(spreadsheetId) {
  // データ種類マスタシートの作成
  const dataTypeSheet = createDataTypeMasterSheet(spreadsheetId);
  
  // テストデータの投入
  testDataTypes.forEach(dataType => {
    addDataTypeMaster(dataType);
  });
  
  // 境界値テストデータの投入
  dataTypeBoundaryTests.forEach(test => {
    try {
      addDataTypeMaster(test.dataType);
      console.log(`境界値テストデータ追加成功: ${test.name}`);
    } catch (error) {
      console.log(`境界値テストデータ追加（期待されるエラー）: ${test.name} - ${error.message}`);
    }
  });
  
  console.log('データ種類マスタテストデータ投入完了');
}

function cleanupTestData() {
  // テストデータの削除
  deleteTestSpreadsheet();
  
  // データ種類マスタテストデータの削除
  cleanupDataTypeTestData();
  
  // テスト用設定のリセット
  resetTestSettings();
  
  console.log('テストデータクリーンアップ完了');
}

function cleanupDataTypeTestData() {
  // テスト用データ種類の削除
  const testDataTypeIds = testDataTypes.map(dt => dt.dataTypeId);
  
  testDataTypeIds.forEach(id => {
    try {
      deleteDataTypeMaster(id);
      console.log(`テストデータ種類削除: ${id}`);
    } catch (error) {
      console.log(`テストデータ種類削除エラー: ${id} - ${error.message}`);
    }
  });
}
```

## 5. テスト実行・管理

### 5.1 テスト実行計画

#### 5.1.1 テスト段階スケジュール
```
Week 1: 単体テスト
├── Day 1-2: バックエンド関数テスト
├── Day 3-4: フロントエンド関数テスト
└── Day 5: テスト結果分析・修正

Week 2: 統合テスト
├── Day 1-2: API連携テスト
├── Day 3-4: データフローテスト
└── Day 5: テスト結果分析・修正

Week 3: システムテスト
├── Day 1-2: 機能テスト
├── Day 3: パフォーマンステスト
├── Day 4: セキュリティテスト
└── Day 5: テスト結果分析・修正

Week 4: 受入テスト
├── Day 1-2: ユーザーシナリオテスト
├── Day 3: ユーザビリティテスト
├── Day 4: アクセシビリティテスト
└── Day 5: 最終結果分析・リリース判定
```

### 5.2 テスト結果管理

#### 5.2.1 テスト結果レポート
```javascript
// テスト結果レポート生成
function generateTestReport() {
  const report = {
    testSuite: 'URL Generation System Test',
    executionDate: new Date().toISOString(),
    environment: 'Test Environment',
    summary: {
      totalTests: 150,
      passed: 145,
      failed: 3,
      skipped: 2,
      passRate: 96.7
    },
    categories: {
      unitTests: { total: 50, passed: 50, failed: 0 },
      integrationTests: { total: 30, passed: 28, failed: 2 },
      systemTests: { total: 40, passed: 38, failed: 1, skipped: 1 },
      acceptanceTests: { total: 30, passed: 29, failed: 0, skipped: 1 }
    },
    failedTests: [
      {
        testId: 'INT-005',
        description: 'Google Sheets API timeout test',
        error: 'Timeout after 10 seconds',
        severity: 'High'
      }
    ]
  };
  
  return report;
}
```

### 5.3 品質メトリクス

#### 5.3.1 テストカバレッジ
- **機能カバレッジ**: 100%（全機能要件のテスト実施）
- **コードカバレッジ**: 80%以上（実行可能コードの測定）
- **パスカバレッジ**: 90%以上（実行パスの網羅）

#### 5.3.2 欠陥密度
- **Critical**: 0件（システム停止レベル）
- **High**: 1件以下（機能不全レベル）
- **Medium**: 5件以下（軽微な機能問題）
- **Low**: 制限なし（UI改善等）

### 5.4 リリース判定基準

#### 5.4.1 Go/No-Go基準
```
Go条件（全て満たす必要あり）:
✅ 全Critical・High欠陥の修正完了
✅ システムテスト通過率95%以上
✅ パフォーマンス要件クリア
✅ セキュリティテスト完全通過
✅ ユーザビリティテスト満足度80%以上
✅ 本番環境での動作確認完了

No-Go条件（いずれか該当でリリース延期）:
❌ Critical欠陥の未修正
❌ システムテスト通過率90%未満
❌ セキュリティ脆弱性の発見
❌ パフォーマンス要件の大幅未達成
```

## 6. 継続的品質改善

### 6.1 テスト自動化
- **CI/CDパイプライン**: コミット時の自動テスト実行
- **回帰テスト**: 定期的な自動実行
- **パフォーマンス監視**: 継続的な性能測定

### 6.2 品質監視
- **エラー率監視**: 本番環境でのエラー発生率追跡
- **ユーザーフィードバック**: 継続的なユーザー満足度調査
- **システムメトリクス**: 使用量・性能データの分析

### 6.3 改善サイクル
```
月次品質レビュー:
1. テスト結果分析
2. 品質課題の特定
3. 改善施策の立案
4. 次月の品質目標設定
```