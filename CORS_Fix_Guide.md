# CORS 修復指南

## 問題說明
當您從本地文件（origin 'null'）訪問 Google Apps Script 時，會遇到 CORS 阻止錯誤。

## 已實施的修復

### 1. Google Apps Script 修復
已在 `gas-script.js` 中添加：

```javascript
/**
 * Handle OPTIONS requests for CORS
 */
function doOptions(e) {
  console.log('=== doOptions called ===');
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, DELETE')
    .setHeader('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization, Cache-Control')
    .setHeader('Access-Control-Allow-Credentials', 'true')
    .setHeader('Access-Control-Max-Age', '86400');
}
```

### 2. 響應格式修復
將 MIME 類型改為 `TEXT`：

```javascript
return ContentService
  .createTextOutput(jsonResponse)
  .setMimeType(ContentService.MimeType.TEXT)
  .setHeader('Access-Control-Allow-Origin', '*')
  // ... 其他標頭
```

### 3. 請求數據解析增強
支持多種請求格式：

```javascript
// 從 URL 參數讀取（URL-encoded）
if (e.parameter && e.parameter.data) {
  data = JSON.parse(e.parameter.data);
}
// 從 POST 正文讀取（JSON）
else if (e.postData && e.postData.contents) {
  data = JSON.parse(e.postData.contents);
}
```

### 4. 前端請求修復
使用 URL-encoded 格式，對本地文件更友好：

```javascript
const formData = new URLSearchParams({
  'data': JSON.stringify(requestData)
});

const response = await fetch(SCRIPT_URL, {
  method: 'POST',
  headers: {
    'Content-Type': 'application/x-www-form-urlencoded',
  },
  body: formData
});
```

## 部署步驟

### 1. 更新 Google Apps Script
1. 打開 Google Apps Script 控制台
2. 找到您的項目
3. 將修復後的代碼替換到 `Code.gs` 文件中
4. 點擊 "Deploy" → "New deployment"
5. 選擇 "Web app" 類型
6. 將 "Execute as" 設為 "Me"
7. 將 "Who has access" 設為 "Anyone"
8. 點擊 "Deploy"
9. 複製新的 Web app URL

### 2. 更新 HTML 文件
將新的 Google Apps Script URL 替換到 `index.html` 中的 `SCRIPT_URL` 變量。

### 3. 測試修復
1. 打開 `index.html`
2. 打開瀏覽器開發者工具（F12）
3. 查看控制台消息
4. 嘗試提交數據

## 如果仍然遇到問題

### 方案 1: 使用本地 HTTP 服務器
```bash
# 使用 Python 3
python -m http.server 8000

# 使用 Node.js
npx http-server

# 使用 PHP
php -S localhost:8000
```

然後訪問 `http://localhost:8000/index.html`

### 方案 2: 使用 Chrome 無安全模式
```bash
# Windows
chrome.exe --user-data-dir="C:/Chrome dev Session" --disable-web-security

# macOS
open -n -a /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --args --user-data-dir="/tmp/chrome_dev_test" --disable-web-security

# Linux
google-chrome --user-data-dir="/tmp/chrome_dev_test" --disable-web-security
```

### 方案 3: 檢查 Google Apps Script 設置
1. 確保腳本已部署為 "Anyone" 可訪問
2. 確保腳本的執行權限設置正確
3. 檢查 Google Sheets 的共享設置

## 調試技巧

### 1. 檢查網絡請求
- 打開開發者工具的 Network 標籤
- 查看請求是否發送
- 檢查響應狀態碼

### 2. 檢查控制台日誌
- 查看 JavaScript 錯誤
- 查看 API 調用日誌
- 檢查響應數據

### 3. 檢查 Google Apps Script 日誌
1. 打開 Google Apps Script 控制台
2. 點擊 "Executions" 查看執行日誌
3. 查看詳細的錯誤信息

## 常見錯誤及解決方案

### 錯誤 1: "No 'Access-Control-Allow-Origin' header"
- 確保已添加 `doOptions` 函數
- 確保所有響應都包含 CORS 標頭

### 錯誤 2: "405 Method Not Allowed"
- 確保 Google Apps Script 支持 OPTIONS 請求
- 檢查部署設置

### 錯誤 3: "Failed to fetch"
- 檢查網絡連接
- 確保 Google Apps Script URL 正確
- 嘗試使用本地 HTTP 服務器

### 錯誤 4: "Invalid JSON response"
- 檢查 Google Apps Script 是否正確返回 JSON
- 查看響應的原始內容
- 確保沒有 HTML 錯誤頁面

## 支持
如果您仍然遇到問題，請提供：
1. 瀏覽器開發者工具的錯誤信息
2. Google Apps Script 的執行日誌
3. 使用的瀏覽器版本
4. 操作系統信息 