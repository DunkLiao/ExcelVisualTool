# Excel 視覺化桌面工具建議

## 目標

建立一個可在使用者 Windows 電腦執行的 Excel 視覺化工具，讓使用者選取本機 Excel 檔案後，自動解析資料、產生圖表，並提供互動式檢視與匯出功能。

第一版應優先追求「使用者不需要安裝開發環境」，也就是不要求使用者自行安裝 Node.js、Python、pandas 或前端套件。

## 免安裝可行性

此工具可以做成接近免安裝的桌面應用，例如：

- 提供單一 `.exe`。
- 提供 portable 資料夾，使用者解壓後直接執行。
- 所有前端資源與必要執行邏輯都包在應用程式內。

但要注意，嚴格意義上的「完全免安裝」仍有前提：

- Tauri 在 Windows 上依賴 Microsoft WebView2 Runtime。
- Windows 10/11 多數環境通常已內建 WebView2，但不能保證所有使用者電腦都有。
- 若目標環境沒有 WebView2，使用者仍可能需要安裝 WebView2 Runtime，或由安裝流程自動補齊。

因此建議對外描述為：

> 可提供 Windows portable 版本，使用者不需安裝 Node.js、Python 或其他開發工具；若使用者電腦缺少 WebView2 Runtime，需由安裝包或首次啟動流程補齊。

## 推薦技術棧

第一版建議採用：

- Tauri
- React
- TypeScript
- ECharts
- Rust 或前端 Excel parser

不建議第一版使用 Python/pandas 作為主要資料處理核心，除非確定有很強的資料清理或分析需求。

## 為什麼第一版避免 Python/pandas

Python/pandas 很適合資料處理，但對免安裝桌面工具會增加部署複雜度。

如果使用 Python/pandas，必須處理：

- Python runtime 打包。
- pandas、openpyxl 等套件打包。
- sidecar 執行檔管理。
- 不同 Windows 環境下的相容性測試。
- 應用程式體積變大。

如果沒有把 Python runtime 和依賴一起打包，使用者就必須自行安裝 Python 與相關套件，這不符合免安裝目標。

因此第一版建議先用 Rust 或前端套件直接解析 Excel，降低部署風險。

## Excel 處理方式

第一版建議支援 `.xlsx`，並以讀取檔案內容為主，不依賴使用者電腦安裝 Microsoft Excel。

可行方式：

- 前端使用 SheetJS 之類的套件解析 Excel。
- 或後端使用 Rust Excel parser 解析檔案。

如果只是讀取 Excel 檔案並產生圖表，不需要安裝 Microsoft Excel。

只有在以下需求出現時，才需要考慮依賴本機 Excel：

- 操作使用者已安裝的 Excel 應用程式。
- 使用 COM automation。
- 執行 VBA。
- 保留高度複雜的 Excel 格式、巨集或互動行為。

這些需求會讓部署、權限與相容性都變複雜，不建議放入 MVP。

## 架構與資料流

建議資料流如下：

1. 使用者在桌面工具中選取 Excel 檔案。
2. 應用程式讀取 `.xlsx`。
3. 解析工作表、欄位名稱與資料列。
4. 前端提供欄位選擇與圖表類型設定。
5. 使用 ECharts 產生互動式圖表。
6. 使用者可匯出圖表圖片或儲存分析設定。

資料預設只在本機處理，不上傳伺服器。這對企業、銀行、內部報表等情境比較容易被接受。

## 前端 UI 與圖表設計

第一版 UI 建議保持工具導向，不做行銷式首頁。

主要畫面可包含：

- 檔案選擇區。
- 工作表選擇。
- 資料預覽表格。
- 欄位映射設定。
- 圖表類型切換。
- 圖表預覽。
- 匯出按鈕。

圖表第一版建議支援：

- 長條圖。
- 折線圖。
- 圓餅圖。
- 散佈圖。
- 堆疊長條圖。

## 打包與部署方式

建議提供兩種發佈形式：

- Portable 版本：解壓後直接執行，最符合免安裝期待。
- Installer 版本：適合需要建立捷徑、自動處理 WebView2 Runtime 或後續自動更新的情境。

若第一版不使用 Python/pandas，打包會相對單純：

- 前端建置成靜態檔。
- Tauri 封裝桌面殼與本機能力。
- Excel 解析邏輯隨應用程式一起打包。

若未來需要 Python/pandas，應將它設計成可替換的資料處理模組，並把 Python runtime 與依賴完整封裝成 sidecar，不要求使用者另行安裝。

## MVP 功能建議

第一版建議聚焦以下功能：

- 開啟本機 `.xlsx` 檔案。
- 顯示工作表清單。
- 預覽資料前數百列。
- 自動辨識欄位型態。
- 選擇 X 軸、Y 軸與分類欄位。
- 產生常見互動式圖表。
- 匯出圖表為 PNG。
- 儲存最近開啟檔案與圖表設定。

## 不建議第一版採用的技術選擇

以下項目不建議放入第一版：

- Python/pandas 作為必要執行環境。
- 依賴使用者本機 Microsoft Excel。
- 使用 VBA 或 COM automation。
- Electron，除非團隊更熟悉 Electron 且不介意應用程式體積較大。
- 一開始就做雲端上傳與多人協作。
- 一開始就支援所有複雜 Excel 格式與巨集。

## 結論

若目標是讓使用者在自己的 Windows 電腦上盡量免安裝使用，建議第一版採用 Tauri + React + TypeScript + ECharts，並用 Rust 或前端 Excel parser 處理 `.xlsx`。

Python/pandas 可以保留為未來進階資料處理選項，但不建議作為 MVP 的必要依賴。這樣能降低打包難度、減少使用者環境問題，也更符合 portable 桌面工具的定位。
