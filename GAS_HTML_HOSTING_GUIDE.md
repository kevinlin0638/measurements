# Google Apps Script HTML Hosting 部署指南

## 🎉 完美解決方案！

你的想法太棒了！使用 Google Apps Script 直接 host HTML 文件，這樣就完全避免了 CORS 問題，因為前端和後端都在同一個域名下。

## 🚀 優點

✅ **完全避免 CORS 問題** - 前端和後端同域名  
✅ **簡單部署** - 只需要一個 Google Apps Script 專案  
✅ **統一管理** - 所有代碼都在一個地方  
✅ **免費 hosting** - Google 提供免費的 Web App 服務  
✅ **自動 HTTPS** - Google 自動提供 SSL 憑證  

## 📋 部署步驟

### 步驟 1: 創建 Google Apps Script 專案

1. 前往 [Google Apps Script](https://script.google.com/)
2. 點擊 **"新專案"**
3. 將專案命名為 `Measurement-BOM-System`

### 步驟 2: 設定 Google Apps Script 代碼

1. 刪除預設的 `Code.gs` 內容
2. 複製 `gas-script-with-html.js` 的全部內容到 `Code.gs`
3. 修改 `SPREADSHEET_ID` 為你的 Google Sheets ID

### 步驟 3: 創建 HTML 文件

1. 點擊 **"+"** 按鈕 → 選擇 **"HTML 文件"**
2. 命名為 `index`（重要：必須是 index）
3. 刪除預設內容
4. 複製 `gas-html-page.html` 的全部內容到 `index.html`

### 步驟 4: 部署 Web App

1. 點擊 **"部署"** → **"新增部署"**
2. 設定：
   - **類型**: 選擇 "Web 應用程式"
   - **執行身分**: 選擇 "我"
   - **存取權限**: 選擇 "任何人"
   - **描述**: 輸入 "Measurement BOM System"

3. 點擊 **"部署"**
4. 授權必要的權限
5. **複製 Web App URL**

### 步驟 5: 測試系統

1. 打開 Web App URL
2. 你應該看到完整的系統界面
3. 測試各項功能：
   - 測量數據輸入
   - BOM 數據輸入
   - 查詢和更新功能

## 🛠️ 專案結構

```
Google Apps Script 專案
├── Code.gs                 # 後端邏輯 (gas-script-with-html.js)
├── index.html              # 前端界面 (gas-html-page.html)
└── Google Sheets          # 數據存儲
    ├── bom 工作表
    └── measurements 工作表
```

## 🔧 主要差異

### 與之前版本的不同：

1. **API 調用方式**：
   ```javascript
   // 之前 (外部 API)
   fetch(GAS_WEB_APP_URL, { method: 'POST', body: data })
   
   // 現在 (內部調用)
   google.script.run.handleApiCall(data)
   ```

2. **無需 CORS 設定**：
   - 前端和後端都在 `script.google.com` 域名下
   - 瀏覽器不會阻止同域名請求

3. **部署更簡單**：
   - 只需要一個 Google Apps Script 專案
   - 無需 GitHub Pages 或其他 hosting 服務

## 🎯 功能完整性

✅ **數據輸入頁籤**：
- 測量數據輸入（支持多筆測量）
- BOM 數據輸入（包含下拉選單）

✅ **查詢更新頁籤**：
- BOM 樹狀結構顯示
- 點擊部件顯示序列號
- 點擊序列號顯示測量數據
- 編輯/刪除測量數據

✅ **系統功能**：
- 響應式設計（支援手機和桌面）
- 載入提示
- 錯誤處理
- 成功/失敗訊息

## 🔍 除錯指南

### 如果遇到問題：

1. **檢查 Google Apps Script 日誌**：
   - 在 Google Apps Script 中點擊 **"執行"** → **"檢視執行記錄"**

2. **測試後端函數**：
   - 執行 `testScript` 函數
   - 檢查是否能成功寫入 Google Sheets

3. **檢查 HTML 文件**：
   - 確保 HTML 文件命名為 `index`
   - 確保所有代碼都正確複製

## 💡 進階功能

### 可以進一步擴展的功能：

1. **數據匯出**：
   - 匯出 CSV 或 Excel 格式
   - 產生報告

2. **用戶管理**：
   - 登入系統
   - 權限控制

3. **數據視覺化**：
   - 圖表顯示
   - 統計分析

4. **自動化**：
   - 定期備份
   - 自動通知

## 🎉 部署完成！

完成部署後，你將擁有一個：

- **完全功能的 Web 應用程式**
- **無 CORS 問題**
- **專業的用戶界面**
- **完整的 CRUD 功能**
- **響應式設計**

---

## 🚀 立即開始

1. 複製 `gas-script-with-html.js` 內容到 Google Apps Script 的 `Code.gs`
2. 創建 `index.html` 文件並複製 `gas-html-page.html` 內容
3. 部署為 Web App
4. 開始使用你的專業數據管理系統！

**這個解決方案完全避免了 CORS 問題，是最優雅的解決方案！** 🎯 