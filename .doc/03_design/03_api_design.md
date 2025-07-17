# API 設計書

## 1. API 設計概要

### 1.1 アーキテクチャ

```
Frontend (JavaScript) ←→ Google Apps Script (Backend) ←→ Google Sheets (Database)
```

### 1.2 通信方式

- **プロトコル**: `google.script.run` を使用した RPC（Remote Procedure Call）
- **データ形式**: JSON 形式でのパラメータ・レスポンス
- **認証**: Google OAuth 2.0 による自動認証
- **エラーハンドリング**: `.withSuccessHandler()` / `.withFailureHandler()`

### 1.3 共通レスポンス形式

```javascript
// 成功時
{
  success: true,
  data: {...},
  message: "処理が正常に完了しました",
  timestamp: "2024/07/02 10:30:00"
}

// エラー時
{
  success: false,
  error: "エラーの詳細メッセージ",
  errorCode: "ERROR_CODE",
  timestamp: "2024/07/02 10:30:00"
}
```

## 2. 拠点管理番号・URL 生成 API

### 2.1 共通フォーム URL 生成

#### 2.1.1 `generateCommonFormUrl(locationNumber, deviceCategory, generateQrUrl)`

**目的**: 拠点管理番号から共通フォーム URL を生成

**パラメータ**:

```javascript
{
  locationNumber: string,    // 拠点管理番号 例: "OSA_SV_ThinkPad_ABC123_001"
  deviceCategory: string,    // デバイスカテゴリ 例: "Server", "Desktop", "Printer", "Tablet", "Router", "Hub"
  generateQrUrl: boolean     // QRコード用URL生成フラグ（省略可能、デフォルト: false）
}
```

**レスポンス**:

```javascript
{
  success: true,
  url: "https://docs.google.com/forms/d/.../viewform?entry.1372464946=OSA_SV_ThinkPad_ABC123_001",
  baseUrl: "https://docs.google.com/forms/d/.../viewform",
  locationNumber: "OSA_SV_ThinkPad_ABC123_001",
  deviceCategory: "Server",
  isQrUrl: false
}
```

**エラーケース**:

- 拠点管理番号形式不正
- デバイスカテゴリ不正
- 共通フォーム URL 未設定

#### 2.1.2 `generateAndSaveCommonFormUrl(requestData)`

**目的**: URL 生成とマスタデータ保存を一括実行

**パラメータ**:

```javascript
{
  locationNumber: string,
  deviceCategory: string,
  generateQrUrl: boolean,
  deviceInfo: {              // デバイス情報（自動登録用）
    modelName: string,
    manufacturer: string,
    category: string,
    serialNumber: string,
    assetNumber: string      // 資産管理番号（端末マスタの場合は必須）
  }
}
```

**レスポンス**:

```javascript
{
  success: true,
  locationNumber: "OSA_SV_ThinkPad_ABC123_001",
  deviceCategory: "Server",
  generatedUrl: "https://docs.google.com/forms/...",
  baseUrl: "https://docs.google.com/forms/...",
  savedTo: "端末マスタ",
  savedRow: 15,
  savedColumn: 8,
  isQrUrl: false,
  isNewEntry: false
}
```

### 2.2 QRコード生成 API

#### 2.2.1 `generateQRCodeWithImage(locationNumber, deviceCategory)`

**目的**: QRコード用URLを生成し、サーバーサイドでQRコード画像データも生成

**パラメータ**:

```javascript
{
  locationNumber: string,    // 拠点管理番号
  deviceCategory: string     // デバイスカテゴリ
}
```

**レスポンス**:

```javascript
{
  success: true,
  url: "https://script.google.com/.../exec?id=OSA_SV_ThinkPad_ABC123_001",
  baseUrl: "https://script.google.com/.../exec",
  locationNumber: "OSA_SV_ThinkPad_ABC123_001",
  deviceCategory: "Server",
  isQrUrl: true,
  imageData: "data:image/png;base64,iVBORw0KGgoAAAANS...",  // Base64エンコードされた画像データ
  imageProvider: "GoQR.me"  // 使用されたQRコード生成API
}
```

**エラーケース**:

- URL生成失敗
- QRコード画像生成失敗（imageErrorフィールドにエラー内容を含む）
- 全APIの応答タイムアウト

#### 2.2.2 `generateQRCodeImage(url)`

**目的**: URLからQRコード画像をサーバーサイドで生成（内部API）

**使用するQRコード生成API（優先順）**:

1. **GoQR.me API** - 高信頼性、EU GDPR準拠
2. **QRCode Monkey API** - 商用利用可能
3. **QuickChart.io** - 高速レスポンス

**セキュリティ対策**:

- HTTPS通信の強制
- SSL/TLS証明書の検証
- 入力URLの検証とエンコーディング
- 適切なUser-Agentヘッダー

## 3. マスタデータ管理 API

### 3.1 拠点マスタ API

#### 3.1.1 `getLocationMaster()`

**目的**: 拠点マスタデータの取得

**パラメータ**: なし

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      locationId: "osaka",
      locationCode: "OSAKA",
      locationName: "大阪営業所",
      jurisdiction: "関西",
      groupEmail: "test-group@example.com",
      statusChangeNotification: true,
      status: "active",
      createdAt: "2024/07/02 14:30:00",
      updatedAt: "2024/07/02 14:30:00"
    }
  ],
  totalCount: 6,
  jurisdictions: ["関西", "関東", "九州"]
}
```

#### 3.1.2 `addLocation(locationData)`

**目的**: 新規拠点の追加

**パラメータ**:

```javascript
{
  locationId: "osaka",
  locationCode: "OSAKA",
  locationName: "大阪営業所",
  jurisdiction: "関西",
  groupEmail: "test-group@example.com",
  statusChangeNotification: false,
  status: "active"
}
```

**レスポンス**:

```javascript
{
  success: true,
  message: "拠点が正常に追加されました",
  addedLocationId: "osaka",
  savedRow: 4
}
```

#### 3.1.3 `updateLocation(locationId, updateData)`

**目的**: 既存拠点情報の更新

**パラメータ**:

```javascript
{
  locationId: "osaka",
  updateData: {
    locationName: "大阪営業所（更新）",
    locationCode: "OSAKA_UPDATED",
    jurisdiction: "関西",
    statusChangeNotification: true
  }
}
```

### 3.2 機種マスタ API

#### 3.2.1 `getModelMasterData()`

**目的**: 機種マスタデータの取得

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      modelId: "MODEL_001",
      modelName: "ThinkPad-X1-Carbon",
      manufacturer: "Lenovo",
      category: "Laptop",
      createdAt: "2024/07/02 14:30:00",
      updatedAt: "2024/07/02 14:30:00"
    }
  ],
  totalCount: 10
}
```

#### 3.2.2 `addModelMasterData(modelData)`

**目的**: 新規機種の追加

**パラメータ**:

```javascript
{
  modelName: "ThinkPad-X1-Carbon",
  manufacturer: "Lenovo",
  category: "Laptop",
  remarks: "第10世代"
}
```

#### 3.2.3 `getAvailableCategories()`

**目的**: 機種マスタから利用可能なカテゴリ一覧を取得

**パラメータ**: なし

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      categoryId: "Desktop",
      categoryName: "デスクトップPC",
      modelCount: 5
    },
    {
      categoryId: "Laptop",
      categoryName: "ノートPC",
      modelCount: 8
    }
  ],
  totalCount: 8
}
```

#### 3.2.4 `getModelsByCategory(category)`

**目的**: カテゴリ別機種一覧の取得

**パラメータ**:

```javascript
{
  category: "Laptop"; // "Desktop", "Laptop", "Server", "Tablet", "Printer", "Router", "Hub", "Other"
}
```

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      modelName: "ThinkPad-X1-Carbon",
      manufacturer: "Lenovo",
      displayName: "Lenovo - ThinkPad-X1-Carbon"
    }
  ],
  category: "Laptop",
  totalCount: 5
}
```

### 3.3 データタイプマスタ API

#### 3.3.1 `getDataTypeMaster()`

**目的**: データタイプマスタデータの取得

**パラメータ**: なし

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      dataTypeId: "NORMAL",
      dataTypeName: "通常データ",
      description: "日常的な機器管理データ",
      displayOrder: 1,
      status: "active",
      createdAt: "2024/07/02 14:30:00",
      updatedAt: "2024/07/02 14:30:00"
    }
  ],
  totalCount: 3
}
```

#### 3.3.2 `addDataType(dataTypeData)`

**目的**: 新規データタイプの追加

**パラメータ**:

```javascript
{
  dataTypeId: "CUSTOM",
  dataTypeName: "カスタムデータ",
  description: "カスタマイズされたデータタイプ",
  displayOrder: 4,
  status: "active"
}
```

**レスポンス**:

```javascript
{
  success: true,
  message: "データタイプが正常に追加されました",
  addedDataTypeId: "CUSTOM",
  savedRow: 4
}
```

#### 3.3.3 `updateDataType(dataTypeId, updateData)`

**目的**: 既存データタイプ情報の更新

**パラメータ**:

```javascript
{
  dataTypeId: "NORMAL",
  updateData: {
    dataTypeName: "通常データ（更新）",
    displayOrder: 2
  }
}
```

#### 3.3.4 `getActiveDataTypes()`

**目的**: アクティブなデータタイプのみ取得

**パラメータ**: なし

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      dataTypeId: "NORMAL",
      dataTypeName: "通常データ",
      description: "日常的な機器管理データ",
      displayOrder: 1
    }
  ],
  totalCount: 3
}
```

### 3.4 機器マスタ API

#### 3.4.1 `getLocationAllDevices(locationId, dataType, statusFilter)`

**目的**: 拠点別全機器データの統合取得（ステータスフィルタ対応）

**データソース**: 
- マスタシート（端末マスタ、プリンタマスタ、その他マスタ）から取得
- マスタシートは収集シートから最新データを自動的に反映（QUERY/数式）
- リアルタイムで最新のフォーム回答内容を表示

**パラメータ**:

```javascript
{
  locationId: "osaka",           // 拠点ID
  dataType: "NORMAL",           // データタイプID ("NORMAL", "AUDIT", "SUMMARY")
  statusFilter: "1.貸出中"      // ステータスフィルタ（省略可能）
}
```

#### 3.4.1-2 `getJurisdictionDevices(jurisdiction, dataType, statusFilter)`

**目的**: 管轄別全機器データの統合取得（サマリー・監査データ用）

**パラメータ**:

```javascript
{
  jurisdiction: "関西",        // 管轄名
  dataType: "AUDIT",            // データタイプID ("AUDIT", "SUMMARY")
  statusFilter: "1.貸出中"   // ステータスフィルタ（省略可能）
}
```

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      拠点管理番号: "OSAKA_SV_ThinkPad-X1-Carbon_ABC123_001",
      機器種別: "端末",
      機種名: "ThinkPad-X1-Carbon",
      製造番号: "ABC123456",
      利用状況: "利用中",
      拠点名: "大阪",
      管轄: "関西"
    }
  ],
  jurisdiction: "関西",
  includedLocations: ["大阪", "神戸", "姫路"],
  totalCount: 45,
  lastUpdate: "2024/07/02 15:00:00"
}
```

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      拠点管理番号: "OSAKA_SV_ThinkPad-X1-Carbon_ABC123_001",
      機器種別: "端末",
      機種名: "ThinkPad-X1-Carbon",
      製造番号: "ABC123456",
      利用状況: "利用中",
      共通フォームURL: "https://forms.google.com/...",
      更新日時: "2024/07/02 14:30:00"
    },
    {
      拠点管理番号: "OSAKA_Printer_MFP_XYZ123_001",
      機器種別: "プリンタ",
      機種名: "imageCLASS MF644Cdw",
      製造番号: "XYZ123456",
      利用状況: "利用中",
      共通フォームURL: "https://forms.google.com/...",
      更新日時: "2024/07/02 15:00:00"
    }
  ],
  locationId: "osaka",
  totalCount: 25,
  lastUpdate: "2024/07/02 15:00:00"
}
```

#### 3.4.2 `getDevicesByMasterType(masterType, locationId, dataType)`

**目的**: マスタ種別による機器データ取得

**パラメータ**:

```javascript
{
  masterType: "terminal", // "terminal", "printer", "other"
  locationId: "osaka",      // 拠点ID（省略可能）
  dataType: "NORMAL"      // データタイプID
}
```

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      拠点管理番号: "OSAKA_SV_ThinkPad-X1-Carbon_ABC123_001",
      機種名: "ThinkPad-X1-Carbon",
      製造番号: "ABC123456",
      // ... その他の端末固有フィールド
    }
  ],
  masterType: "terminal",
  totalCount: 15
}
```

#### 3.4.3 `getAvailableStatuses(includeNestedStatuses)`

**目的**: システム内で利用可能なステータス一覧を取得（動的ネストステータス対応）

**パラメータ**:

```javascript
{
  includeNestedStatuses: true; // ネストステータスも含めるかどうか（省略可能）
}
```

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      statusId: "1.貸出中",
      statusName: "貸出中",
      statusType: "rental",
      count: 25
    },
    {
      statusId: "2.回収後社内保管",
      statusName: "社内保管",
      statusType: "storage",
      count: 15
    },
    {
      statusId: "3.社内にて保管中",
      statusName: "社内保管中",
      statusType: "internal_storage",
      count: 30,
      hasNestedStatuses: true,
      nestedStatuses: [
        {
          statusId: "3-1.貸し出し可能(初期化済み)",
          statusName: "貸し出し可能",
          count: 15,
          detectedPrefix: "3-1",
          hasColumns: true
        },
        {
          statusId: "3-2.初期化待ち",
          statusName: "初期化待ち",
          count: 8,
          detectedPrefix: "3-2",
          hasColumns: true
        },
        {
          statusId: "3-5.新しいステータス",
          statusName: "新しいステータス",
          count: 7,
          detectedPrefix: "3-5",
          hasColumns: true
        }
      ]
    }
  ],
  totalCount: 3,
  dynamicDetection: {
    detectedNestedPrefixes: ["3-1", "3-2", "3-5"],
    availableColumnPatterns: [
      "3-1-1.初期化完了日", "3-1-2.動作確認日",
      "3-2-1.初期化予定日", "3-2-2.担当者",
      "3-5-1.新項目1", "3-5-2.新項目2"
    ]
  }
}
```

#### 3.4.4 `getAvailableJurisdictions()`

**目的**: 利用可能な管轄一覧を取得

**パラメータ**: なし

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      jurisdictionId: "関西",
      jurisdictionName: "関西エリア",
      locations: [
        { locationId: "osaka", locationName: "大阪営業所" },
        { locationId: "kobe", locationName: "神戸営業所" },
        { locationId: "himeji", locationName: "姫路営業所" }
      ],
      totalLocations: 3
    },
    {
      jurisdictionId: "関東",
      jurisdictionName: "関東エリア",
      locations: [
        { locationId: "tokyo", locationName: "東京営業所" },
        { locationId: "yokohama", locationName: "横浜営業所" }
      ],
      totalLocations: 2
    }
  ],
  totalCount: 3
}
```

#### 3.4.5 `getDataTypeDisplayConfig(dataTypeId, statusFilter, nestedStatusFilter, allAvailableColumns)`

**目的**: データタイプとステータスに応じた表示設定を取得（動的 2 段階ネストフィルタリング対応）

**パラメータ**:

```javascript
{
  dataTypeId: "NORMAL",
  statusFilter: "3.社内にて保管中",    // 省略可能
  nestedStatusFilter: "3-5.新しいステータス", // 省略可能（動的検出された社内ステータス）
  allAvailableColumns: [             // 利用可能な全列一覧（省略可能）
    "拠点管理番号", "機器種別", "機種名", "製造番号",
    "3-0.社内ステータス",
    "3-1-1.初期化完了日", "3-1-2.動作確認日",
    "3-2-1.初期化予定日", "3-2-2.担当者",
    "3-5-1.新項目1", "3-5-2.新項目2", "3-5-3.新項目3"
  ]
}
```

**レスポンス**（動的 2 段階ネストフィルタリング適用時）:

```javascript
{
  success: true,
  data: {
    dataTypeId: "NORMAL",
    statusDriven: true,
    prefixBasedFiltering: true,
    nestedStatusFiltering: true,
    dynamicNestedFiltering: true,
    currentStatus: "3.社内にて保管中",
    nestedStatus: "3-5.新しいステータス",
    statusPrefix: "3",
    nestedStatusPrefix: "3-5",
    columns: [
      "拠点管理番号",
      "機器種別",
      "機種名",
      "製造番号",
      "3-0.社内ステータス",
      "3-5-1.新項目1",
      "3-5-2.新項目2",
      "3-5-3.新項目3"
    ],
    headers: [
      { key: "拠点管理番号", label: "管理番号", width: "180px" },
      { key: "機器種別", label: "種別", width: "70px" },
      { key: "機種名", label: "機種", width: "130px" },
      { key: "製造番号", label: "製造番号", width: "110px" },
      { key: "3-0.社内ステータス", label: "社内ステータス", width: "120px" },
      { key: "3-5-1.新項目1", label: "新項目1", width: "120px" },
      { key: "3-5-2.新項目2", label: "新項目2", width: "120px" },
      { key: "3-5-3.新項目3", label: "新項目3", width: "120px" }
    ],
    filteringLogic: {
      description: "動的2段階ネストフィルタリング：社内ステータス（3-5）に基づいて3-5-*の列のみ表示",
      level: "dynamic_nested",
      baseColumns: ["拠点管理番号", "機器種別", "機種名", "製造番号"],
      statusColumn: "3-0.社内ステータス",
      filteredColumns: ["3-5-1.新項目1", "3-5-2.新項目2", "3-5-3.新項目3"],
      dynamicDetection: {
        detectedPrefixes: ["3-1", "3-2", "3-5"],
        selectedPrefix: "3-5",
        availableColumns: [
          "3-1-1.初期化完了日", "3-1-2.動作確認日",
          "3-2-1.初期化予定日", "3-2-2.担当者",
          "3-5-1.新項目1", "3-5-2.新項目2", "3-5-3.新項目3"
        ]
      }
    }
  }
}
```

## 4. システム設定 API

### 4.1 共通フォーム設定 API

#### 4.1.1 `getSystemSettings()`

**目的**: システム設定全体の取得

**レスポンス**:

```javascript
{
  success: true,
  data: {
    terminalCommonFormUrl: "https://docs.google.com/forms/d/.../viewform",
    printerCommonFormUrl: "https://docs.google.com/forms/d/.../viewform",
    qrRedirectUrl: "https://script.google.com/macros/s/.../exec",
    errorNotificationEmail: "admin@example.com,team@example.com",
    alertNotificationEmail: "alert@example.com",
    debugMode: false,
    editableSettings: {
      formUrls: false,      // debugMode依存
      logNotifications: false  // debugMode依存（ログ通知設定）
    },
    statusChangeNotificationEnabled: true  // グローバル通知有効/無効設定
  }
}
```

#### 4.1.2 `saveSystemSettings(settings)`

**目的**: システム設定の保存（debugMode制御付き）

**パラメータ**:

```javascript
{
  // URL設定（debugMode=true時のみ変更可能）
  terminalCommonFormUrl: "https://docs.google.com/forms/d/.../viewform",
  printerCommonFormUrl: "https://docs.google.com/forms/d/.../viewform",
  qrRedirectUrl: "https://script.google.com/macros/s/.../exec",
  
  // ログ通知設定（debugMode=true時のみ変更可能）
  errorNotificationEmail: "admin@example.com,team@example.com",
  alertNotificationEmail: "alert@example.com",
  
  // ステータス変更通知グローバル設定（常に変更可能）
  statusChangeNotificationEnabled: true
}
```

**レスポンス（成功時）**:

```javascript
{
  success: true,
  message: "システム設定が正常に保存されました",
  savedSettings: {
    terminalCommonFormUrl: "...",
    printerCommonFormUrl: "...",
    qrRedirectUrl: "...",
    errorNotificationEmail: "...",
    alertNotificationEmail: "..."
  }
}
```

**レスポンス（debugMode無効時）**:

```javascript
{
  success: false,
  error: "設定の変更にはデバッグモードを有効にしてください",
  errorCode: "DEBUG_MODE_REQUIRED",
  details: {
    currentDebugMode: false,
    affectedSettings: ["formUrls", "notifications"]
  }
}
```

#### 4.1.3 `validateGoogleFormUrl(formUrl)`

**目的**: Google Forms の URL 検証

**パラメータ**:

```javascript
{
  formUrl: "https://docs.google.com/forms/d/1FAIpQLSe.../viewform";
}
```

**レスポンス**:

```javascript
{
  success: true,
  valid: true,
  title: "代替機申請フォーム",
  formId: "1FAIpQLSe...",
  accessible: true
}
```

### 4.2 通知設定 API

#### 4.2.1 `validateEmailAddresses(emailString)`

**目的**: メールアドレスの形式検証（複数アドレス対応）

**パラメータ**:

```javascript
{
  emailString: "admin@example.com,team@example.com,invalid-email"
}
```

**レスポンス**:

```javascript
{
  success: true,
  allValid: false,
  validEmails: ["admin@example.com", "team@example.com"],
  invalidEmails: ["invalid-email"],
  totalCount: 3,
  validCount: 2
}
```

#### 4.2.2 `sendTestNotification(notificationType, emailAddresses)`

**目的**: 通知メールのテスト送信

**パラメータ**:

```javascript
{
  notificationType: "error",  // "error" または "alert"
  emailAddresses: "admin@example.com,team@example.com"
}
```

**レスポンス**:

```javascript
{
  success: true,
  message: "テストメールを送信しました",
  sentTo: ["admin@example.com", "team@example.com"],
  sentAt: "2024/07/02 10:30:00"
}
```

#### 4.2.3 `toggleDebugMode(enabled)`

**目的**: デバッグモードの切り替え（管理者のみ）

**パラメータ**:

```javascript
{
  enabled: true  // true または false
}
```

**レスポンス**:

```javascript
{
  success: true,
  message: "デバッグモードを有効にしました",
  debugMode: true,
  editableSettings: {
    formUrls: true,
    logNotifications: true
  }
}
```

#### 4.2.4 `toggleGlobalStatusChangeNotification(enabled)`

**目的**: グローバルステータス変更通知の有効/無効切り替え

**パラメータ**:

```javascript
{
  enabled: true        // true（有効）または false（無効）
}
```

**レスポンス**:

```javascript
{
  success: true,
  message: "グローバル通知設定を更新しました",
  enabled: true
}
```

#### 4.2.5 `sendStatusChangeNotification(locationId, changeData)`

**目的**: ステータス変更通知の送信（フォーム送信時に自動実行）

**実行タイミング**: 
- Google Formsの回答がステータス収集シートに保存される際に自動的に実行
- onFormSubmitトリガーから呼び出される
- スプレッドシートの直接編集では実行されない

**パラメータ**:

```javascript
{
  locationId: "osaka",
  changeData: {
    deviceId: "OSA_SV_ThinkPad_ABC123_001",
    oldStatus: "2.商談や金額の問題で返却不可",  // 前回の値（収集シートから取得）
    newStatus: "1.返却可能",                   // フォームで送信された新しい値
    changedBy: "田中太郎",                     // フォーム回答者
    changedAt: "2024/07/15 14:30:00",         // フォーム送信タイムスタンプ
    formResponse: {                           // フォーム回答の詳細
      responseId: "2KqWr...",
      formId: "1FAIpQL..."
    }
  }
}
```

**レスポンス**:

```javascript
{
  success: true,
  message: "ステータス変更通知を送信しました",
  sentTo: "osaka@example.com",
  sentAt: "2024/07/15 14:30:15",
  notificationEnabled: true,
  globalNotificationEnabled: true
}
```

**注意**: 
- 拠点マスタで通知がOFFの場合、通知は送信されません
- PropertiesServiceでグローバル通知が無効の場合も送信されません

#### 4.2.6 `getLocationNotificationSettings()`

**目的**: 全拠点の通知設定一覧を取得

**パラメータ**: なし

**レスポンス**:

```javascript
{
  success: true,
  data: [
    {
      locationId: "osaka",
      locationName: "大阪営業所",
      groupEmail: "osaka@example.com",
      notificationEnabled: true
    },
    {
      locationId: "kobe",
      locationName: "神戸営業所",
      groupEmail: "kobe@example.com",
      notificationEnabled: false
    }
  ],
  globalNotificationEnabled: true
}
```

## 5. ユーティリティ API

### 5.1 ログ・パフォーマンス API

#### 5.1.1 `getSystemLogs()`

**目的**: システムログの取得（管理者用）

**レスポンス**:

```javascript
{
  success: true,
  logs: [
    {
      timestamp: "2024/07/02 10:30:00",
      level: "INFO",
      message: "URL生成完了",
      data: {...},
      user: "testuser@example.com"
    }
  ],
  totalCount: 100
}
```

#### 5.1.2 `getPerformanceMetrics()`

**目的**: パフォーマンス指標の取得

**レスポンス**:

```javascript
{
  success: true,
  metrics: {
    averageResponseTime: 1.2,  // 秒
    totalRequests: 1500,
    errorRate: 0.02,           // 2%
    lastUpdate: "2024/07/02 10:30:00"
  }
}
```

### 5.2 データ検証 API

#### 5.2.1 `validateLocationNumber(locationNumber)`

**目的**: 拠点管理番号の形式検証

**パラメータ**:

```javascript
{
  locationNumber: "OSA_SV_ThinkPad_ABC123_001";
}
```

**レスポンス**:

```javascript
{
  success: true,
  valid: true,
  components: {
    location: "OSA",
    category: "SV",
    model: "ThinkPad",
    serialNumber: "ABC123",
    sequence: "001"
  },
  exists: false  // システム内での重複チェック
}
```

## 6. エラー処理

### 6.1 エラーコード体系

| コード           | 分類           | 説明                     |
| ---------------- | -------------- | ------------------------ |
| AUTH_ERROR       | 認証           | 認証・認可エラー         |
| VALIDATION_ERROR | バリデーション | 入力値検証エラー         |
| DATA_ERROR       | データ         | データ不整合・制約違反   |
| SYSTEM_ERROR     | システム       | システム内部エラー       |
| EXTERNAL_ERROR   | 外部連携       | 外部 API・サービスエラー |
| TIMEOUT_ERROR    | タイムアウト   | 処理時間超過             |

### 6.2 エラーレスポンス例

#### 6.2.1 バリデーションエラー

```javascript
{
  success: false,
  error: "入力値に不正があります",
  errorCode: "VALIDATION_ERROR",
  details: {
    field: "locationNumber",
    message: "拠点管理番号の形式が正しくありません",
    expected: "拠点_カテゴリ_モデル_製造番号_連番"
  },
  timestamp: "2024/07/02 10:30:00"
}
```

#### 6.2.2 データエラー

```javascript
{
  success: false,
  error: "指定されたデータが見つかりません",
  errorCode: "DATA_ERROR",
  details: {
    resource: "拠点マスタ",
    locationId: "XXX",
    message: "拠点ID 'XXX' は存在しません"
  },
  timestamp: "2024/07/02 10:30:00"
}
```

#### 6.2.3 システムエラー

```javascript
{
  success: false,
  error: "システムエラーが発生しました",
  errorCode: "SYSTEM_ERROR",
  details: {
    message: "Google Sheets APIの呼び出しに失敗しました",
    retryable: true,
    supportContact: "システム管理者にお問い合わせください"
  },
  timestamp: "2024/07/02 10:30:00"
}
```

## 7. パフォーマンス仕様

### 7.1 レスポンス時間目標

| API 分類         | 目標時間 | 最大時間 |
| ---------------- | -------- | -------- |
| データ取得（小） | 1 秒     | 3 秒     |
| データ取得（大） | 3 秒     | 5 秒     |
| データ更新       | 2 秒     | 4 秒     |
| URL 生成         | 1 秒     | 2 秒     |
| 設定変更         | 2 秒     | 3 秒     |

### 7.2 並行処理制限

- **同時リクエスト**: 最大 10 リクエスト/秒
- **実行時間制限**: 6 分（Google Apps Script 制限）
- **メモリ制限**: 100MB（Google Apps Script 制限）

### 7.3 キャッシュ戦略

- **マスタデータ**: 5 分間キャッシュ
- **設定データ**: 10 分間キャッシュ
- **動的データ**: キャッシュなし（リアルタイム取得）

## 8. セキュリティ仕様

### 8.1 アクセス制御

- **認証**: Google OAuth 2.0 必須
- **認可**: 組織ドメインのみアクセス許可
- **セッション**: Google 管理のセッション

### 8.2 データ保護

- **入力検証**: 全パラメータのサニタイゼーション
- **出力エスケープ**: XSS 攻撃対策
- **SQL インジェクション**: パラメータ化クエリ（該当なし）

### 8.3 ログ・監査

- **アクセスログ**: 全 API 呼び出しの記録
- **エラーログ**: エラー発生時の詳細記録
- **操作ログ**: データ変更操作の履歴

## 9. API 使用例

### 9.1 フロントエンド実装例

#### 9.1.1 基本的な API 呼び出し

```javascript
// URL生成API呼び出し
google.script.run
  .withSuccessHandler(function (result) {
    if (result.success) {
      console.log("URL生成成功:", result.url);
      displayGeneratedUrl(result);
    } else {
      console.error("URL生成失敗:", result.error);
      showErrorMessage(result.error);
    }
  })
  .withFailureHandler(function (error) {
    console.error("通信エラー:", error);
    showErrorMessage("通信エラーが発生しました");
  })
  .generateCommonFormUrl(locationNumber, deviceCategory);
```

#### 9.1.2 プロミス化実装例

```javascript
function callGASFunction(functionName, ...args) {
  return new Promise((resolve, reject) => {
    google.script.run
      .withSuccessHandler(resolve)
      .withFailureHandler(reject)
      [functionName](...args);
  });
}

// 使用例
async function generateUrl() {
  try {
    const result = await callGASFunction(
      "generateCommonFormUrl",
      locationNumber,
      deviceCategory
    );

    if (result.success) {
      displayResult(result);
    } else {
      showError(result.error);
    }
  } catch (error) {
    showError("通信エラー: " + error.message);
  }
}
```

### 9.2 エラーハンドリング実装例

```javascript
function handleApiError(error, context = "") {
  const errorMessage = error.error || error.message || "不明なエラー";
  const errorCode = error.errorCode || "UNKNOWN_ERROR";

  // エラーログ記録
  console.error(`API Error [${context}]:`, {
    code: errorCode,
    message: errorMessage,
    details: error.details,
    timestamp: "2024/07/02 10:30:00",
  });

  // ユーザー向けエラー表示
  let userMessage;
  switch (errorCode) {
    case "VALIDATION_ERROR":
      userMessage = "入力内容を確認してください: " + errorMessage;
      break;
    case "AUTH_ERROR":
      userMessage = "アクセス権限がありません。管理者にお問い合わせください。";
      break;
    case "TIMEOUT_ERROR":
      userMessage =
        "処理に時間がかかっています。しばらく待ってから再試行してください。";
      break;
    default:
      userMessage =
        "システムエラーが発生しました。管理者にお問い合わせください。";
  }

  showErrorNotification(userMessage);
}
```
