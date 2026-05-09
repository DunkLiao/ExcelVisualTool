# Excel 視覺化桌面工具 MVP 落地計畫

## Summary

將 `idea.md` 收斂成第一版可開發的 Windows 桌面工具：使用者選取本機 `.xlsx`，工具解析工作表與欄位，提供資料預覽、欄位映射、互動式圖表與 PNG 匯出。第一版定位為本機處理、免開發環境安裝、portable 優先。

## Key Changes

- 技術棧採用 `Tauri + React + TypeScript + ECharts`。
- Excel 解析第一版採用前端 SheetJS 類套件處理 `.xlsx`，避免 Rust parser 與前後端資料橋接過早複雜化。
- 不依賴 Microsoft Excel、Python、pandas、VBA、COM automation 或雲端上傳。
- 發佈形式以 Windows portable 為主，installer 作為第二種包裝，用於補 WebView2 Runtime、捷徑與後續更新。

## MVP 功能

- 開啟本機 `.xlsx` 檔案。
- 顯示工作表清單並切換工作表。
- 預覽前 200-500 列資料。
- 自動推斷欄位型態：文字、數值、日期、空值比例。
- 圖表設定支援：
  - X 軸欄位
  - Y 軸欄位
  - 分類欄位
  - 圖表類型
- 第一版圖表支援：
  - 長條圖
  - 折線圖
  - 圓餅圖
  - 散佈圖
  - 堆疊長條圖
- 匯出目前圖表為 PNG。
- 儲存最近開啟檔案與最近一次圖表設定到本機。

## UI / Data Flow

- 主畫面直接是工具介面，不做 landing page。
- 版面分成四個區域：
  - 左側：檔案與工作表選擇。
  - 中央上方：資料預覽表。
  - 右側：圖表設定面板。
  - 中央下方或主區：ECharts 圖表預覽。
- 資料流：
  1. 使用者選取 `.xlsx`。
  2. 前端讀取檔案 binary。
  3. SheetJS 解析 workbook、sheet、headers、rows。
  4. 型態推斷產出欄位 metadata。
  5. 使用者設定欄位映射。
  6. 將資料轉成 ECharts option。
  7. 匯出 PNG 或保存設定。

## Test Plan

- 使用 3 份測試 Excel：
  - 一般銷售資料：日期、品項、地區、金額。
  - 缺值與混合型態資料：驗證欄位推斷與錯誤提示。
  - 多工作表檔案：驗證 sheet 切換與設定重置。
- 驗證情境：
  - 可開啟 `.xlsx`。
  - 可切換工作表。
  - 預覽表不因大量資料卡死。
  - 每種圖表類型可正常生成。
  - 不合法欄位組合會顯示明確提示。
  - PNG 匯出成功。
  - 關閉重開後可看到最近檔案與設定。
- 打包驗證：
  - Windows portable build 可在無 Node.js、無 Python 的環境執行。
  - 缺 WebView2 的情境列為 installer 驗證項，不阻塞 portable MVP。

## Assumptions

- 第一版只支援 `.xlsx`，不支援 `.xls`、`.csv`、巨集或高度複雜格式保留。
- 第一版資料只在本機處理，不做登入、雲端同步、多人協作。
- 單檔資料量以中小型報表為目標；大型資料效能最佳化延後處理。
- 圖表設定先做通用欄位映射，不做自然語言問答或 AI 自動分析。
