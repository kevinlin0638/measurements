# MRP 系統 Google Apps Script 部署指南

## 📋 系統概述

這是一個完整的 MRP (Material Requirements Planning) 物料需求規劃系統，使用 Google Apps Script 開發，可以託管在 Google 雲端並與 Google Sheets 整合。

## 🚀 部署步驟

### 步驟 1: 建立 Google Apps Script 專案

1. 前往 [Google Apps Script](https://script.google.com/)
2. 點擊「新增專案」
3. 將專案命名為「MRP 物料需求規劃系統」

### 步驟 2: 建立程式碼檔案

#### 2.1 建立主要程式碼檔案
1. 在 Google Apps Script 編輯器中，將預設的 `Code.gs` 重新命名為 `mrp-system.gs`
2. 將 `mrp-system.gs` 的內容複製到這個檔案中
3. 確保更新 `SPREADSHEET_ID` 為您的 Google Sheets ID

#### 2.2 建立 HTML 模板檔案
1. 點擊左側的「+」按鈕，選擇「HTML」
2. 將檔案命名為 `MRP_HTML`
3. 將 `MRP_HTML.html` 的內容複製到這個檔案中

### 步驟 3: 設定 Google Sheets

#### 3.1 建立 Google Sheets
1. 前往 [Google Sheets](https://sheets.google.com/)
2. 建立新的試算表
3. 複製試算表的 ID（從 URL 中取得）

#### 3.2 更新程式碼中的 Sheets ID
在 `mrp-system.gs` 中更新以下常數：
```javascript
const SPREADSHEET_ID = '您的試算表ID';
```

#### 3.3 建立必要的工作表
在 Google Sheets 中建立以下工作表：
- `mrp_data` - MRP 資料工作表
- `mrp_results` - MRP 結果工作表
- `demand` - 需求資料工作表
- `bom` - BOM 結構工作表

### 步驟 4: 設定權限

#### 4.1 啟用必要的 API
1. 在 Google Apps Script 編輯器中，點擊「服務」
2. 啟用以下服務：
   - Google Sheets API
   - Google Drive API

#### 4.2 設定執行權限
1. 點擊「部署」→「新增部署」
2. 選擇「網頁應用程式」
3. 設定執行身分為「自己」
4. 設定存取權限為「任何人」

### 步驟 5: 部署應用程式

#### 5.1 建立網頁應用程式
1. 點擊「部署」→「新增部署」
2. 選擇「網頁應用程式」
3. 設定以下參數：
   - **執行身分**: 自己
   - **存取權限**: 任何人
   - **版本**: 新版本
4. 點擊「部署」

#### 5.2 取得應用程式 URL
部署完成後，您會得到一個網頁應用程式的 URL，類似：
```
https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec
```

## 🔧 系統功能

### 主要功能
1. **成品需求管理** - 設定各期間的成品需求量
2. **庫存管理** - 期初庫存和動態庫存追蹤
3. **BOM 結構管理** - 完整的組件結構管理
4. **前置時間管理** - 支援不同組件的前置時間設定
5. **供應商管理** - 為每個組件指定供應商
6. **期間推進模擬** - 可以逐期間推進模擬
7. **Google Sheets 整合** - 資料可以儲存和載入

### API 端點
系統提供以下 API 端點：
- `initialize` - 初始化系統
- `setBasicParameters` - 設定基本參數
- `setDemand` - 設定需求資料
- `addComponent` - 新增組件
- `runSimulation` - 執行 MRP 模擬
- `getCurrentPeriod` - 獲取當前期間結果
- `nextPeriod` - 推進到下一個期間
- `resetPeriod` - 重置期間
- `getTableData` - 獲取表格資料
- `exportToSheets` - 匯出到 Google Sheets
- `loadFromSheets` - 從 Google Sheets 載入
- `saveToSheets` - 儲存到 Google Sheets

## 📊 資料結構

### Google Sheets 工作表結構

#### demand 工作表
| 期間 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 |
|------|---|---|---|---|---|---|---|---|---|-----|
| 需求 | 0 | 0 | 0 | 0 | 4 | 0 | 0 | 0 | 0 | 4 |

#### bom 工作表
| 組件名稱 | 需求量 | 供應商 | 前置時間 |
|----------|--------|--------|----------|
| Component A | 2 | Supplier A | 2 |
| Component B | 1 | Supplier B | 1 |

#### mrp_results 工作表
系統會自動產生完整的 MRP 結果表格，包含：
- 成品需求
- 期初庫存
- 庫存缺口
- 生產開始
- 組件訂單明細

## 🛠️ 故障排除

### 常見問題

#### 1. 權限錯誤
**問題**: 出現權限相關錯誤
**解決方案**:
- 確保已啟用必要的 Google API
- 檢查 Google Sheets 的共用權限
- 確認 Apps Script 的執行身分設定

#### 2. 工作表不存在
**問題**: 找不到指定的工作表
**解決方案**:
- 確保在 Google Sheets 中建立了所有必要的工作表
- 檢查工作表名稱是否正確
- 執行 `createRequiredSheets()` 函數自動建立工作表

#### 3. 資料載入失敗
**問題**: 無法從 Google Sheets 載入資料
**解決方案**:
- 檢查資料格式是否正確
- 確保資料欄位不為空
- 檢查 Google Sheets 的存取權限

### 除錯技巧

#### 1. 查看執行記錄
1. 在 Google Apps Script 編輯器中點擊「執行記錄」
2. 查看詳細的錯誤訊息

#### 2. 測試 API 端點
可以使用以下方式測試 API：
```javascript
// 在 Google Apps Script 編輯器中測試
function testAPI() {
  const result = initializeMRPSystem();
  console.log(result);
}
```

#### 3. 檢查 Google Sheets 權限
確保 Google Sheets 的共用設定允許 Apps Script 存取。

## 📈 使用範例

### 基本使用流程

1. **開啟應用程式**
   - 使用部署後的網頁應用程式 URL
   - 系統會顯示三個主要頁面

2. **設定基本參數**
   - 在「輸入資料」頁面設定模擬期間數
   - 設定成品名稱和期初庫存

3. **設定需求資料**
   - 輸入各期間的成品需求
   - 使用逗號分隔格式

4. **設定 BOM 結構**
   - 新增組件並設定相關參數
   - 包括需求量、供應商、前置時間

5. **執行模擬**
   - 點擊「執行 MRP 模擬」
   - 系統會計算並顯示結果

6. **查看結果**
   - 在「模擬結果」頁面查看完整表格
   - 在「MRP 模擬」頁面逐步推進期間

### 範例資料
根據您提供的範例，系統會產生以下結果：

| 項目 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 |
|------|---|---|---|---|---|---|---|---|---|-----|
| 成品需求 | 0 | 0 | 0 | 0 | 4 | 0 | 0 | 0 | 0 | 4 |
| 期初庫存 | 1 | 1 | 1 | 1 | 1 | 0 | 0 | 0 | 0 | 0 |
| 庫存缺口 | 0 | 0 | Null | Null | 3 | Null | Null | Null | Null | 4 |
| 生產開始 | 0 | 0 | 0 | 3 | 0 | 0 | 0 | 0 | 4 | 0 |

## 🔄 更新和維護

### 更新程式碼
1. 在 Google Apps Script 編輯器中修改程式碼
2. 點擊「部署」→「管理部署」
3. 選擇現有部署並點擊「編輯」
4. 選擇「新版本」並重新部署

### 備份資料
- 定期備份 Google Sheets 中的資料
- 使用 Google Drive 的版本控制功能
- 考慮匯出重要資料到本地檔案

## 📞 支援

如有任何問題或需要協助，請：
1. 檢查本指南的故障排除章節
2. 查看 Google Apps Script 的官方文件
3. 檢查 Google Sheets API 的權限設定

---

**版本**: 1.0  
**更新日期**: 2024年  
**開發者**: MRP System Team 