# Google Apps Script 部署指南 - 解決 CORS 問題

## 🚨 重要提醒
如果你遇到 CORS 錯誤，請**完全按照**以下步驟重新部署 Google Apps Script。

## 步驟 1: 更新 Google Apps Script 代碼

1. 前往 [Google Apps Script](https://script.google.com/)
2. 打開你的專案
3. 將整個 `gas-script.js` 文件內容複製並貼到 `Code.gs` 文件中（**完全替換**原有代碼）
4. 確保 `SPREADSHEET_ID` 設定正確

## 步驟 2: 重新部署 Web App

### 🔴 這是最關鍵的步驟 - 必須重新部署！

1. 點擊右上角的 **"部署"** 按鈕
2. 選擇 **"新增部署"**
3. 設定如下：
   - **類型**: 選擇 "Web 應用程式"
   - **執行身分**: 選擇 "我"
   - **存取權限**: **必須選擇 "任何人"**（這是解決 CORS 的關鍵）
   - **描述**: 寫入 "CORS Fix Version"

4. 點擊 **"部署"**
5. 授權必要的權限
6. **複製新的 Web App URL**

## 步驟 3: 測試 Google Apps Script

1. 在 Google Apps Script 中點擊 **"執行"** 按鈕
2. 在下方的 **"執行"** 選項中選擇 `testScript` 函數
3. 點擊執行並檢查是否成功

## 步驟 4: 更新前端 API URL

將新的 Web App URL 更新到你的前端代碼中：

```javascript
// 在 index.html 中找到這行，並更新 URL
const GAS_WEB_APP_URL = 'https://script.google.com/macros/s/YOUR_NEW_DEPLOYMENT_ID/exec';
```

## 步驟 5: 清除瀏覽器緩存

1. 按 `Ctrl + Shift + Delete` (Windows) 或 `Cmd + Shift + Delete` (Mac)
2. 選擇清除 **"快取的圖片和檔案"**
3. 重新載入你的 GitHub Pages 網站

## 步驟 6: 驗證修復

1. 打開瀏覽器開發者工具 (F12)
2. 切換到 **"Console"** 頁籤
3. 重新載入網站
4. 檢查是否還有 CORS 錯誤

## 常見問題解決

### 問題 1: 仍然有 CORS 錯誤
**解決方案**: 確保 Google Apps Script 的存取權限設定為 **"任何人"**

### 問題 2: 401 或 403 錯誤
**解決方案**: 
1. 檢查 Google Apps Script 的執行權限
2. 確保已授權 Google Sheets API 權限

### 問題 3: 網站無法載入序列號
**解決方案**:
1. 檢查 Google Sheets 中是否有資料
2. 確認 `SPREADSHEET_ID` 設定正確
3. 執行 `testScript` 函數驗證連接

## 除錯步驟

如果問題仍然存在，請按照以下步驟除錯：

1. **檢查 Google Apps Script 日誌**:
   - 在 Google Apps Script 中點擊 **"執行"** → **"檢視執行記錄"**
   - 檢查是否有錯誤訊息

2. **檢查瀏覽器 Console**:
   - 打開 F12 開發者工具
   - 查看 Console 中的詳細錯誤信息

3. **測試 API 端點**:
   - 在瀏覽器中直接訪問你的 Web App URL
   - 應該看到 JSON 響應，包含 `success: true`

## 重要提醒

- **每次修改 Google Apps Script 代碼後，都必須重新部署**
- **存取權限必須設定為 "任何人"**
- **使用新的部署 URL，不要使用舊的**
- **清除瀏覽器緩存很重要**

## 支援

如果依然有問題，請提供以下資訊：

1. 瀏覽器 Console 中的完整錯誤訊息
2. Google Apps Script 執行記錄的截圖
3. 你的 Web App URL（去掉敏感部分）

---

**⚠️ 記住：CORS 問題通常是由於 Google Apps Script 的部署設定造成的，重新部署並正確設定存取權限通常可以解決問題！** 