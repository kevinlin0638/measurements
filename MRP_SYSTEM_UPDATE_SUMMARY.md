# MRP 系統修改總結

## 修改概述

根據您的需求，我已經對 MRP 物料需求規劃系統進行了以下主要修改：

### 1. 資料結構簡化

#### 基本設定合併
- **原結構**: 模擬期間數、成品名稱、期初庫存分別設定
- **新結構**: 產品名稱 (PART NUMBER) 和期初庫存合併為一個設定區塊

#### BOM 結構簡化
- **原結構**: 組件名稱、需求量、供應商、前置時間
- **新結構**: 組件名稱、需求量、供應商 (移除前置時間)

### 2. Google Sheets 資料結構更新

#### 新增工作表
1. **Products** - 產品庫存管理
   - 產品名稱
   - 期初庫存
   - 各期間需求 (期間1-10)

2. **BOM** - 簡化版 BOM 結構
   - 產品名稱
   - 組件名稱
   - 需求量
   - 供應商

3. **Component_Inventory** - 組件庫存追蹤
   - 組件名稱
   - 供應商
   - 各期間庫存

4. **Settings** - 系統設定
   - 模擬期間數
   - 當前期間
   - 預設產品

5. **MRP_Results** - MRP 結果輸出

#### 範例資料
```javascript
// Products 工作表範例
['產品名稱', '期初庫存', '需求期間1', '需求期間2', ..., '需求期間10']
['Car Model A', 1, 0, 0, 0, 0, 4, 0, 0, 0, 0, 4]
['Car Model B', 2, 0, 0, 0, 0, 2, 0, 0, 0, 0, 2]

// BOM 工作表範例
['產品名稱', '組件名稱', '需求量', '供應商']
['Car Model A', 'Engine', 1, 'Supplier A']
['Car Model A', 'Body', 1, 'Supplier B']
['Car Model A', 'Transmission', 1, 'Supplier C']
```

### 3. 輸出表格格式

#### 第一個表格：成品 MRP 表
```
Time Period    1    2    3    4    5    6    7    8    9    10
Demand for Finished Goods    0    0    0    0    4    0    0    0    0    4
On-Hand Inventory    1    1    1    1    1    0    0    0    0    0
Inventory Gap    0    0    Null    Null    3    Null    Null    Null    Null    4
Production Start    0    0    0    3    0    0    0    0    4    0
```

#### 第二個表格：組件庫存表
```
Time Period    Production Start    Engine LT    Engine Qty (X1)    Body LT    Body Qty (X1)    ...    Supplier Name
1    0    Null    0    Null    0    ...    A
2    0    Null    3    Null    0    ...    A
3    0    Null    0    Null    0    ...    A
...
```

### 4. 主要功能修改

#### MRPSystem 類別更新
- 新增 `products` 物件管理多產品庫存
- 新增 `currentProduct` 追蹤當前期間分析的產品
- 簡化組件結構，移除前置時間
- 更新 MRP 計算邏輯以支援組件庫存追蹤

#### Google Sheets 整合更新
- `initializeGoogleSheets()` - 建立新的工作表結構
- `loadMRPFromSheets()` - 讀取產品和 BOM 資料
- `writeMRPResultsToSheet()` - 產生兩個表格格式
- `exportMRPToSheets()` - 匯出到 Google Sheets

#### HTML 介面更新
- 合併基本設定為產品設定
- 簡化 BOM 輸入表單
- 更新 JavaScript 函數以支援新資料結構

### 5. 使用流程

1. **初始化**: 執行 `initializeGoogleSheets()` 建立工作表
2. **輸入產品**: 選擇產品名稱和期初庫存
3. **設定需求**: 輸入各期間需求
4. **新增組件**: 為產品新增 BOM 組件
5. **執行模擬**: 產生 MRP 計算結果
6. **查看結果**: 顯示兩個表格格式的結果

### 6. 技術改進

- 支援多產品管理
- 簡化使用者介面
- 標準化輸出格式
- 改善資料讀取和寫入邏輯
- 增強錯誤處理和驗證

這些修改使系統更符合您的需求，提供了更簡潔的介面和更清晰的輸出格式。 