# Excel Visual Tool

Excel Visual Tool 是一個 Windows 桌面小工具，可以在本機開啟 `.xlsx`、`.xls` 與 `.csv` 試算表檔案，預覽資料，選擇欄位，快速產生圖表，並把圖表匯出成 PNG 圖片。

這個工具適合想把 Excel 表格資料快速視覺化，但不想上傳檔案到雲端、也不想額外安裝 Python 或寫程式的使用者。

## 可以做什麼

- 開啟本機 `.xlsx`、`.xls` 與 `.csv` 檔案。
- 匯入 UTF-8、Big5/CP950 與 UTF-16 等常見繁體中文 CSV。
- 切換 Excel 裡的不同工作表。
- 預覽資料表內容，最多先顯示前 200 筆資料。
- 自動判斷欄位是文字、數字、日期或混合資料。
- 選擇 X 軸、Y 軸與分類欄位。
- 產生長條圖、折線圖、圓餅圖、散佈圖、堆疊長條圖與時間序列圖。
- 將目前圖表匯出為 PNG 圖片，並儲存在程式旁的 `exports` 資料夾。
- 記住最近開啟的檔案名稱與上次使用的圖表設定。
- 在本機記錄錯誤 log，方便排查問題。
- 只讀取檔案內容，不寫回或轉存原始檔，因此不會破壞原始 Excel 檔案中的複雜格式。

## 目前不支援

- 不支援 Excel 巨集、VBA 或複雜 Excel 物件。
- 工具內的資料預覽只顯示解析出的資料值，不會完整重建公式、樞紐分析表、合併儲存格、條件格式或原始 Excel 樣式。
- 不支援登入、雲端同步、多人協作、AI 自動分析或自然語言問答。
- 不適合非常大型的 Excel 檔案。

## 使用方式

如果你只是要使用程式，請執行：

```text
portableapps\excel-visual-tool.exe
```

開啟後會看到三個主要區域：

- 左側：開啟 Excel 檔案、選擇工作表。
- 中間：上方預覽資料，下方預覽圖表。
- 右側：設定圖表類型、X 軸、Y 軸與分類欄位。

## 基本操作流程

1. 點選 `Open spreadsheet`。
2. 選擇一個 `.xlsx`、`.xls` 或 `.csv` 試算表檔案。
3. 如果 Excel 有多個工作表，從左側選擇要分析的工作表。
4. 在右側選擇圖表類型。
5. 選擇 X 軸欄位與 Y 軸欄位。
6. 如有需要，選擇 Category 分類欄位。
7. 圖表會自動出現在中間下方。
8. 點選 `Export PNG` 匯出圖表圖片。
9. 匯出成功後，右側會顯示 PNG 檔案儲存路徑。

## PNG 匯出位置

`Export PNG` 會把目前圖表輸出到：

```text
portableapps\exports\
```

檔名會使用原始檔名與圖表類型，例如：

```text
portableapps\exports\csv-sales-sample-pie.png
```

如果同名檔案已經存在，程式會自動加上流水號，例如 `csv-sales-sample-pie-1.png`，不會覆蓋既有圖片。

如果匯出失敗，右側設定面板會顯示錯誤訊息，詳細錯誤也會寫入 `portableapps\logs\app.log`。

## 圖表欄位怎麼選

不同圖表需要的欄位略有不同：

- 長條圖、折線圖：X 軸通常選類別或日期，Y 軸選數字欄位。
- 圓餅圖：X 軸選分類名稱，Y 軸選數字欄位。
- 散佈圖：X 軸和 Y 軸都要選數字欄位。
- 堆疊長條圖：X 軸選主要分類，Y 軸選數字欄位，Category 選要堆疊的分類。
- 時間序列圖：X 軸要選日期欄位，Y 軸選數字欄位。

如果欄位不符合圖表需求，右側會顯示提示訊息，例如「Y axis field must be numeric.」代表 Y 軸必須選數字欄位。

## CSV 編碼支援

CSV 匯入會自動嘗試常見繁體中文編碼：

- UTF-8
- UTF-8 BOM
- Big5 / CP950
- UTF-16LE
- UTF-16BE

這代表從新版 Excel、舊版 Excel、Windows 繁體中文系統或部分企業系統匯出的 CSV，多數情況可以直接開啟並正確顯示中文。

如果 CSV 仍出現亂碼，建議先用 Excel、記事本或文字編輯器另存為 UTF-8 CSV，再重新匯入。

## 測試資料

專案內有一些範例試算表檔，可以用來測試功能：

```text
test-data\sales-sample.xlsx
test-data\chinese-sales-sample.xlsx
test-data\chinese-sales-sample.xls
test-data\mixed-missing-sample.xlsx
test-data\multi-sheet-sample.xlsx
test-data\csv-sales-sample.csv
test-data\big5-sales-sample.csv
```

建議第一次使用時可以先開啟 `test-data\chinese-sales-sample.xlsx`，確認中文欄位、中文分類與圖表匯出是否正常。測試 `.csv` 時可使用 `test-data\csv-sales-sample.csv` 與 `test-data\big5-sales-sample.csv`，確認 UTF-8 與 Big5 繁體中文 CSV 不會出現亂碼。

## 系統需求

- Windows 10 或 Windows 11。
- Microsoft Edge WebView2 Runtime。

多數 Windows 10/11 電腦已經內建 WebView2 Runtime。如果程式無法開啟，可能需要安裝 Microsoft Edge WebView2 Runtime。

使用 portable 版本不需要安裝：

- Node.js
- npm
- Python
- pandas
- Microsoft Excel
- VBA 或 COM automation

## 錯誤記錄

程式執行時如果發生錯誤，會寫入：

```text
portableapps\logs\app.log
```

這個檔案主要用來排查問題。一般使用者不需要主動打開它；如果程式開不起來、Excel 檔解析失敗或畫面發生異常，可以把這個 log 檔提供給維護者查看。

## 建置 portable exe

如果你是開發者，或需要重新產生 portable exe，請在專案根目錄執行：

```bat
build.bat
```

成功後會產生：

```text
portableapps\excel-visual-tool.exe
```

`portableapps` 是輸出資料夾，已加入 `.gitignore`，不應提交到 Git。執行程式後產生的 `portableapps\exports` 與 `portableapps\logs` 也是本機輸出資料，不應提交到 Git。

## 開發者指令

安裝依賴：

```bat
npm install
```

啟動前端開發伺服器：

```bat
npm run dev
```

啟動 Tauri 桌面開發模式：

```bat
npm run tauri dev
```

執行基本檢查：

```bat
test.bat
```

`test.bat` 會檢查 npm 依賴、建置前端，並在 `src-tauri` 執行 `cargo check`。

其他 Windows 快捷指令：

```bat
test.bat dev      :: 啟動 Vite dev server
test.bat tauri    :: 啟動 Tauri 桌面應用
test.bat build    :: 呼叫 build.bat 建置 portable exe
test.bat help     :: 顯示說明
```

## 專案資料夾

```text
src\                  React 前端介面
src-tauri\            Tauri / Rust 桌面程式
test-data\            測試用 Excel 檔案
portableapps\         portable exe 輸出資料夾，不提交 Git
dist\                 前端建置輸出，不提交 Git
node_modules\         npm 依賴，不提交 Git
src-tauri\target\     Rust 建置輸出，不提交 Git
```

## 技術簡介

本工具使用 Tauri、React、TypeScript、Vite、ECharts 與 SheetJS 製作。試算表檔案會在本機解析，不會上傳到雲端，也不會寫回原始檔。
