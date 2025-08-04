# MRP 系統 - 直接 Google Sheets 操作指南

## 📋 系統概述

這個 MRP 系統直接與您的 Google Sheets 進行即時操作，無需匯入匯出功能。系統會自動讀取您 Google Sheets 中的資料，並將計算結果直接寫入到指定的工作表中。

## 🔗 Google Sheets 連結

**您的 Google Sheets**: https://docs.google.com/spreadsheets/d/1ieSa4W6_crSYoIobMiEACW2wG-PYb1PLTRFElhS5FoQ/edit?gid=655606014#gid=655606014

## 🚀 部署步驟

### 步驟 1: 建立 Google Apps Script 專案

1. 前往 [Google Apps Script](https://script.google.com/)
2. 點擊「新增專案」
3. 將專案命名為「MRP 物料需求規劃系統」

### 步驟 2: 建立程式碼檔案

#### 2.1 建立主要程式碼檔案
1. 在 Google Apps Script 編輯器中，將預設的 `Code.gs` 重新命名為 `mrp-system.gs`
2. 將 `mrp-system.gs` 的內容複製到這個檔案中
3. **重要**: 系統已設定為您的 Google Sheets ID，無需修改

#### 2.2 建立 HTML 模板檔案
1. 點擊左側的「+」按鈕，選擇「HTML」
2. 將檔案命名為 `MRP_HTML`
3. 將 `MRP_HTML.html` 的內容複製到這個檔案中

### 步驟 3: 設定權限

#### 3.1 啟用必要的 API
1. 在 Google Apps Script 編輯器中，點擊「服務」
2. 啟用以下服務：
   - Google Sheets API
   - Google Drive API

#### 3.2 設定執行權限
1. 點擊「部署」→「新增部署」
2. 選擇「網頁應用程式」
3. 設定執行身分為「自己」
4. 設定存取權限為「任何人」

### 步驟 4: 部署應用程式

1. 點擊「部署」→「新增部署」
2. 選擇「網頁應用程式」
3. 設定以下參數：
   - **執行身分**: 自己
   - **存取權限**: 任何人
   - **版本**: 新版本
4. 點擊「部署」

## 📊 資料格式說明

### 系統會自動讀取以下格式的資料：

#### 1. 需求資料格式
系統會嘗試從以下工作表名稱讀取需求資料：
- `Demand`
- `demand`
- `需求`
- `MRP_Data`

**支援的格式**:
```
期間    1    2    3    4    5    6    7    8    9    10
需求    0    0    0    0    4    0    0    0    0    4
```

或者簡單的數字行：
```
0    0    0    0    4    0    0    0    0    4
```

#### 2. BOM 資料格式
系統會嘗試從以下工作表名稱讀取 BOM 資料：
- `BOM`
- `bom`
- `組件`
- `Components`

**支援的格式**:
```
組件名稱    需求量    供應商    前置時間
Component A    2    Supplier A    2
Component B    1    Supplier B    1
```

#### 3. 基本設定格式
系統會嘗試從 `Settings` 工作表讀取基本設定：
```
設定項目    數值
模擬期間數    10
成品名稱    Product A
期初庫存    1
```

## 🎯 使用流程

### 1. 準備 Google Sheets 資料
在您的 Google Sheets 中準備好以下資料：
- 需求資料（各期間的成品需求）
- BOM 資料（組件結構）
- 基本設定（可選）

### 2. 開啟 MRP 系統
使用部署後的網頁應用程式 URL 開啟系統

### 3. 讀取 Google Sheets 資料
1. 在「輸入資料」頁面
2. 點擊「從 Google Sheets 讀取資料」
3. 系統會自動讀取並更新表單

### 4. 執行 MRP 模擬
1. 點擊「執行 MRP 模擬」
2. 系統會計算 MRP 結果

### 5. 查看結果
1. 在「模擬結果」頁面查看完整表格
2. 在「MRP 模擬」頁面逐步推進期間

### 6. 寫入結果到 Google Sheets
1. 點擊「將結果寫入 Google Sheets」
2. 系統會在您的 Google Sheets 中建立 `MRP_Results` 工作表
3. 結果會包含完整的 MRP 表格和組件訂單明細

## 📈 預期結果

根據您提供的範例，系統會產生以下結果並寫入 Google Sheets：

| 項目 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 |
|------|---|---|---|---|---|---|---|---|---|-----|
| 成品需求 | 0 | 0 | 0 | 0 | 4 | 0 | 0 | 0 | 0 | 4 |
| 期初庫存 | 1 | 1 | 1 | 1 | 1 | 0 | 0 | 0 | 0 | 0 |
| 庫存缺口 | 0 | 0 | Null | Null | 3 | Null | Null | Null | Null | 4 |
| 生產開始 | 0 | 0 | 0 | 3 | 0 | 0 | 0 | 0 | 4 | 0 |

## 🔧 系統特色

### ✅ 直接操作 Google Sheets
- **即時讀取**: 直接從您的 Google Sheets 讀取資料
- **即時寫入**: 計算結果直接寫入 Google Sheets
- **無需匯入匯出**: 系統直接操作，無需手動匯入匯出

### ✅ 智能資料識別
- **自動識別格式**: 系統會自動識別不同格式的資料
- **多工作表支援**: 支援多個工作表名稱
- **容錯處理**: 即使資料格式不完全符合，也能嘗試讀取

### ✅ 完整的 MRP 功能
- **成品需求管理**: 支援各期間需求設定
- **庫存管理**: 期初庫存和動態庫存追蹤
- **BOM 結構管理**: 完整的組件結構管理
- **前置時間管理**: 支援不同組件的前置時間
- **供應商管理**: 為每個組件指定供應商
- **期間推進模擬**: 可以逐期間推進模擬

## 🛠️ 故障排除

### 常見問題

#### 1. 無法讀取 Google Sheets 資料
**解決方案**:
- 確保 Google Sheets 的共用權限設定正確
- 檢查資料格式是否符合系統要求
- 確認工作表名稱是否正確

#### 2. 無法寫入結果到 Google Sheets
**解決方案**:
- 確保已執行 MRP 模擬
- 檢查 Google Sheets 的編輯權限
- 確認 Apps Script 的執行權限

#### 3. 資料格式錯誤
**解決方案**:
- 檢查需求資料是否為數字格式
- 確認 BOM 資料包含所有必要欄位
- 確保資料沒有空行或格式錯誤

### 除錯技巧

#### 1. 查看執行記錄
1. 在 Google Apps Script 編輯器中點擊「執行記錄」
2. 查看詳細的錯誤訊息

#### 2. 測試資料讀取
可以使用以下方式測試資料讀取：
```javascript
// 在 Google Apps Script 編輯器中測試
function testDataReading() {
  const result = loadMRPFromSheets();
  console.log(result);
}
```

## 📞 支援

如有任何問題或需要協助，請：
1. 檢查本指南的故障排除章節
2. 查看 Google Apps Script 的執行記錄
3. 確認 Google Sheets 的權限設定

---

**版本**: 1.0  
**更新日期**: 2024年  
**開發者**: MRP System Team 