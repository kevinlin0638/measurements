# CORS 問題修復說明

## 🚨 問題描述

您遇到的 "405 Method Not Allowed" 錯誤是因為 CORS (跨域資源共享) 預檢請求問題。瀏覽器在發送 POST 請求前會先發送一個 OPTIONS 請求，但原始的 Google Apps Script 不支援這個方法。

## ✅ 解決方案

我已經修改了 Google Apps Script 代碼來解決這個問題。現在您需要重新部署 Google Apps Script。

## 🔧 重新部署步驟

### 步驟 1：更新 Google Apps Script 代碼

1. **開啟您的 Google Apps Script 專案**
   - 前往 [https://script.google.com/](https://script.google.com/)
   - 打開您之前創建的專案

2. **替換代碼**
   - 選擇並刪除所有現有代碼
   - 複製最新的 `gas-script.js` 文件內容
   - 貼入到 Google Apps Script 編輯器中

3. **保存專案**
   - 按 `Ctrl+S` 保存

### 步驟 2：重新部署

1. **管理部署**
   - 點擊「部署」→「管理部署作業」

2. **編輯現有部署**
   - 在現有部署旁邊點擊「編輯」圖標（鉛筆圖標）

3. **更新版本**
   - 在「版本」下拉選單中選擇「新版本」
   - 在「說明」欄位中輸入：「修復 CORS 問題」

4. **重新部署**
   - 點擊「部署」按鈕
   - 等待部署完成

5. **確認 URL**
   - 確認 Web App URL 沒有改變
   - 應該仍然是：`https://script.google.com/macros/s/AKfycbxFMPexBLSndCWUm9-7snOLEzJrNa-HuEBimA0rK8rbR026MG12o20RdMZdmQ8PtVfd/exec`

## 🧪 測試修復

### 測試 1：直接訪問 URL
1. 在瀏覽器中打開您的 Web App URL
2. 您應該看到類似這樣的 JSON 回應：
   ```json
   {
     "success": true,
     "message": "Google Apps Script 運行正常",
     "timestamp": "2024-01-XX..."
   }
   ```

### 測試 2：使用 HTML 表單
1. 開啟 `index.html` 文件
2. 嘗試提交一筆測試資料
3. 應該會看到「資料提交成功！」訊息

## 🔍 修復細節

### 修改內容：
1. **添加 CORS 標頭**
   - 在所有回應中添加適當的 CORS 標頭
   - 允許跨域請求

2. **簡化請求處理**
   - 使用 `text/plain` Content-Type 來避免複雜的預檢請求
   - 保持 JSON 格式的資料傳輸

3. **改進錯誤處理**
   - 更好的錯誤訊息顯示
   - 統一的回應格式

## 🐛 如果問題持續存在

### 檢查清單：
- [ ] 已經重新部署 Google Apps Script
- [ ] 確認 URL 正確
- [ ] 瀏覽器快取已清除
- [ ] 網路連接正常

### 額外故障排除：

1. **清除瀏覽器快取**
   - 按 `Ctrl+Shift+Delete`
   - 清除瀏覽器快取和 cookie

2. **檢查開發者工具**
   - 按 `F12` 開啟開發者工具
   - 查看 Console 和 Network 頁籤
   - 確認請求狀態碼是 200 而不是 405

3. **嘗試不同瀏覽器**
   - 使用 Chrome、Firefox 或 Edge 測試
   - 有時不同瀏覽器的 CORS 處理方式不同

4. **檢查 Google Apps Script 權限**
   - 確認腳本有權限存取 Google Sheets
   - 重新授權如果需要

## 🎯 成功標誌

當修復成功時，您應該看到：
- ✅ 不再出現 405 錯誤
- ✅ 表單提交成功訊息
- ✅ 資料正確出現在 Google Sheets 中
- ✅ 瀏覽器開發者工具顯示 200 狀態碼

## 📞 仍需協助？

如果按照以上步驟仍然無法解決問題，請檢查：
1. Google Apps Script 的執行記錄
2. 瀏覽器開發者工具的詳細錯誤訊息
3. 確認 Google Sheets 的存取權限

---

**重要提醒：** 每次修改 Google Apps Script 代碼後都需要重新部署，否則更改不會生效。 