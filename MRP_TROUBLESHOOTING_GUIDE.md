# MRP 系統故障排除指南

## 連接測試失敗問題診斷

### 問題症狀
- 連接測試結果顯示 `null`
- 初始化結果顯示 `null`
- 無法與 Google Apps Script 建立連接

### 診斷步驟

#### 1. 檢查 Google Apps Script 部署
```javascript
// 在 Google Apps Script 編輯器中執行以下測試函數
function testBasicConnection() {
  try {
    return {
      success: true,
      message: '基本連接測試成功',
      timestamp: new Date().toISOString(),
      spreadsheetId: SPREADSHEET_ID
    };
  } catch (error) {
    return {
      success: false,
      message: '基本連接測試失敗: ' + error.message
    };
  }
}
```

#### 2. 檢查 Google Sheets 存取權限
```javascript
// 測試 Google Sheets 存取
function testSheetsAccess() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    return {
      success: true,
      message: 'Google Sheets 存取成功',
      sheetCount: sheets.length,
      sheetNames: sheets.map(sheet => sheet.getName())
    };
  } catch (error) {
    return {
      success: false,
      message: 'Google Sheets 存取失敗: ' + error.message
    };
  }
}
```

#### 3. 檢查 SPREADSHEET_ID 是否正確
- 確認 `SPREADSHEET_ID = '1ieSa4W6_crSYoIobMiEACW2wG-PYb1PLTRFElhS5FoQ'` 是否為正確的 Google Sheets ID
- 在 Google Sheets URL 中確認 ID：`https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit`

#### 4. 檢查權限設定
- 確保 Google Apps Script 專案有權限存取 Google Sheets
- 在 Google Apps Script 編輯器中，點擊「執行」→「執行函數」→「授權」

#### 5. 檢查部署設定
- 確保 Web App 已正確部署
- 檢查執行身分是否設定為「自己」
- 確保「誰可以存取」設定為「任何人」

### 解決方案

#### 方案 1：重新部署 Web App
1. 在 Google Apps Script 編輯器中
2. 點擊「部署」→「新增部署」
3. 選擇「網頁應用程式」
4. 設定執行身分為「自己」
5. 設定存取權限為「任何人」
6. 點擊「部署」

#### 方案 2：檢查並修正 Google Sheets ID
1. 開啟您的 Google Sheets
2. 複製 URL 中的 ID
3. 更新 `mrp-system.gs` 中的 `SPREADSHEET_ID`
4. 重新部署

#### 方案 3：建立新的 Google Sheets
1. 建立新的 Google Sheets
2. 更新 `SPREADSHEET_ID`
3. 執行初始化函數建立必要的工作表

### 測試步驟

#### 步驟 1：基本連接測試
```javascript
// 在瀏覽器控制台中執行
async function testConnection() {
  try {
    const result = await callGoogleAppsScript('testConnection');
    console.log('連接測試結果:', result);
    return result;
  } catch (error) {
    console.error('連接測試錯誤:', error);
    return null;
  }
}
```

#### 步驟 2：Google Sheets 存取測試
```javascript
// 在 Google Apps Script 編輯器中執行
function testSheetsConnection() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    console.log('可存取的工作表:', sheets.map(sheet => sheet.getName()));
    return { success: true, sheets: sheets.map(sheet => sheet.getName()) };
  } catch (error) {
    console.error('Google Sheets 存取錯誤:', error);
    return { success: false, error: error.message };
  }
}
```

#### 步驟 3：初始化測試
```javascript
// 在 Google Apps Script 編輯器中執行
function testInitialization() {
  try {
    const result = initializeGoogleSheets();
    console.log('初始化結果:', result);
    return result;
  } catch (error) {
    console.error('初始化錯誤:', error);
    return { success: false, error: error.message };
  }
}
```

### 常見錯誤及解決方法

#### 錯誤 1：`ScriptError: We're sorry, a server error occurred while reading from storage. Error code PERMISSION_DENIED.`
**解決方法：**
- 檢查 Google Sheets 的共用設定
- 確保 Google Apps Script 有權限存取該 Google Sheets

#### 錯誤 2：`ScriptError: We're sorry, a server error occurred while reading from storage. Error code NOT_FOUND.`
**解決方法：**
- 檢查 `SPREADSHEET_ID` 是否正確
- 確認 Google Sheets 是否存在且可存取

#### 錯誤 3：`ScriptError: We're sorry, a server error occurred while reading from storage. Error code UNAUTHENTICATED.`
**解決方法：**
- 重新授權 Google Apps Script
- 檢查 Google 帳戶登入狀態

### 進階診斷

#### 檢查網路連接
```javascript
// 在瀏覽器控制台中執行
function checkNetworkConnection() {
  const startTime = Date.now();
  return fetch('https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ action: 'testConnection' })
  })
  .then(response => response.json())
  .then(data => {
    const endTime = Date.now();
    console.log('網路延遲:', endTime - startTime, 'ms');
    return data;
  })
  .catch(error => {
    console.error('網路連接錯誤:', error);
    return null;
  });
}
```

#### 檢查部署 ID
1. 在 Google Apps Script 編輯器中
2. 點擊「部署」→「管理部署」
3. 複製 Web App URL 中的部署 ID
4. 確認 HTML 檔案中的部署 ID 是否正確

### 聯絡支援

如果以上步驟都無法解決問題，請提供以下資訊：
1. 錯誤訊息的完整內容
2. 瀏覽器控制台的錯誤日誌
3. Google Apps Script 執行日誌
4. 使用的瀏覽器版本
5. 作業系統版本

---

## 其他常見問題

### 1. 資料無法正確載入
**症狀：** 從 Google Sheets 載入的資料不正確或為空

**解決方法：**
- 檢查 Google Sheets 中的資料格式
- 確認工作表名稱是否正確
- 執行初始化函數重新建立資料結構

### 2. 模擬結果不正確
**症狀：** MRP 計算結果與預期不符

**解決方法：**
- 檢查輸入資料的正確性
- 確認需求資料格式
- 檢查 BOM 結構設定

### 3. 頁面載入緩慢
**症狀：** 頁面載入時間過長

**解決方法：**
- 檢查網路連接
- 確認 Google Apps Script 部署設定
- 檢查瀏覽器快取設定

### 4. 功能按鈕無回應
**症狀：** 點擊按鈕後沒有反應

**解決方法：**
- 檢查瀏覽器控制台是否有錯誤訊息
- 確認 JavaScript 是否正確載入
- 檢查 Google Apps Script 部署狀態 