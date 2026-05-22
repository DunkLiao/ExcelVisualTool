# Excel 視覺化工具

Excel 視覺化工具是一個 Windows 桌面小工具，可在本機開啟 `.xlsx`、`.xls` 與 `.csv` 試算表，預覽資料、選擇欄位、產生圖表，並將目前圖表匯出成 PNG 圖片。

這個專案使用 Tauri、React、TypeScript、Vite、ECharts 與 SheetJS 製作。試算表檔案會在本機解析，不會上傳到雲端，也不會寫回原始檔。

## 功能

- 開啟本機 `.xlsx`、`.xls` 與 `.csv` 檔案。
- 自動嘗試 UTF-8、Big5、UTF-16LE 與 UTF-16BE 等常見 CSV 編碼。
- 切換活頁簿中的不同工作表。
- 預覽資料表內容，最多先顯示前 200 筆資料列。
- 使用 SQLite SQL 語法篩選目前工作表資料，再用查詢結果產生預覽與圖表。
- 自動判斷欄位型別：文字、數字、日期、混合、空白。
- 選擇 X 軸、Y 軸與分類欄位。
- 支援長條圖、折線圖、圓餅圖、散佈圖、堆疊長條圖、時間序列圖、面積圖、水平長條圖、雷達圖、矩形式樹狀圖、漏斗圖與儀表圖。
- 將目前圖表匯出為 PNG 圖片，儲存在執行檔旁的 `exports` 資料夾。
- 記住最近開啟的檔案名稱與上次使用的圖表設定。
- 將前端錯誤與匯出狀態寫入本機 log，方便排查問題。

## 使用方式

如果只是要使用 portable 版本，請執行：

```text
portableapps\excel-visual-tool.exe
```

開啟後會看到三個主要區域：

- 左側：開啟試算表、選擇工作表。
- 中間：上方輸入 SQL 篩選語法，中段顯示資料預覽，下方顯示圖表預覽。
- 右側：設定圖表類型、X 軸、Y 軸、分類欄位，並匯出 PNG。

## 基本操作流程

1. 點選 `開啟試算表`。
2. 選擇 `.xlsx`、`.xls` 或 `.csv` 檔案。
3. 如果活頁簿有多個工作表，從左側選擇要視覺化的工作表。
4. 如需篩選資料，在中間上方輸入 SQL 後點選 `執行 SQL`。
5. 在右側選擇圖表類型。
6. 選擇 X 軸欄位與 Y 軸欄位。
7. 視圖表需要選擇分類欄位。
8. 圖表會自動顯示在中間下方。
9. 點選 `匯出 PNG`。
10. 匯出成功後，右側會顯示 PNG 檔案儲存路徑。

## SQL 篩選

載入檔案後，程式會把目前選取的工作表放進本機記憶體中的 SQLite 資料表。資料表名稱固定為 `data`，預設查詢是：

```sql
SELECT * FROM data
```

可用 SQLite 的 `SELECT` 語法篩選、排序或限制資料筆數。查詢結果會成為資料預覽、欄位清單與圖表的來源。

例如用 `LIKE` 篩選文字：

```sql
SELECT *
FROM data
WHERE "產品" LIKE '%A%'
```

篩選數值並排序：

```sql
SELECT *
FROM data
WHERE "金額" > 1000
ORDER BY "日期"
```

只取部分欄位與前 20 筆：

```sql
SELECT "日期", "產品", "金額"
FROM data
WHERE "地區" = '北區'
LIMIT 20
```

欄位名稱如果包含中文、空白或特殊字元，請用 SQLite 雙引號包起來，例如 `"客戶名稱"`。SQL 語法錯誤時，畫面會顯示錯誤訊息，並保留上一個成功查詢結果。

## 圖表欄位建議

不同圖表需要的欄位略有不同：

- 長條圖、折線圖、面積圖、水平長條圖：X 軸通常選類別或日期，Y 軸選數字欄位。
- 圓餅圖：X 軸選分類名稱，Y 軸選數字欄位。
- 散佈圖：X 軸與 Y 軸都要選數字欄位。
- 堆疊長條圖：X 軸選主要分類，Y 軸選數字欄位，分類欄位選要堆疊的分組。
- 時間序列圖：X 軸選日期欄位，Y 軸選數字欄位。
- 雷達圖：X 軸選指標名稱，Y 軸選數字欄位，可用分類欄位分成多組資料。
- 矩形式樹狀圖、漏斗圖：X 軸選分類名稱，Y 軸選數字欄位。
- 儀表圖：Y 軸選數字欄位，程式會彙總目前資料作為主要指標值。

如果欄位不符合圖表需求，右側會顯示繁體中文提示，例如「Y 軸欄位必須是數值。」或「請選擇分類欄位。」。

## PNG 匯出

PNG 會輸出到執行檔旁的 `exports` 資料夾。例如 portable 版本的輸出位置是：

```text
portableapps\exports\
```

檔名會使用原始檔名與繁體中文圖表類型，例如：

```text
portableapps\exports\sales-sample-圓餅圖.png
```

如果同名檔案已存在，程式會自動加上流水號，例如 `sales-sample-圓餅圖-1.png`，不會覆蓋既有圖片。

## CSV 編碼支援

CSV 匯入會自動嘗試：

- UTF-8
- UTF-8 BOM
- Big5 / CP950
- UTF-16LE
- UTF-16BE

從新版 Excel、舊版 Excel、Windows 繁體中文系統或企業系統匯出的 CSV，多數情況可以直接開啟並正確顯示中文。如果仍出現亂碼，建議先用 Excel、記事本或文字編輯器另存為 UTF-8 CSV，再重新匯入。

## 目前限制

- 不支援 Excel 巨集、VBA 或複雜 Excel 物件。
- 資料預覽只顯示解析出的資料值，不會重建公式、樞紐分析表、合併儲存格、條件格式或原始 Excel 樣式。
- SQL 查詢只針對目前選取的工作表，資料表名稱固定為 `data`；目前不支援跨工作表 `JOIN`。
- SQL 面板僅支援會回傳結果的 `SELECT` 或 `WITH` 查詢，不支援修改資料的 `INSERT`、`UPDATE`、`DELETE`。
- 不支援登入、雲端同步、多人協作、AI 自動分析或自然語言問答。
- 非常大型的活頁簿可能載入較慢或不適合使用。

## 測試資料

專案內提供測試用檔案：

```text
test-data\sales-sample.xlsx
test-data\chinese-sales-sample.xlsx
test-data\chinese-sales-sample.xls
test-data\mixed-missing-sample.xlsx
test-data\multi-sheet-sample.xlsx
test-data\csv-sales-sample.csv
test-data\big5-sales-sample.csv
```

第一次使用可先開啟 `test-data\chinese-sales-sample.xlsx`，確認中文欄位、中文分類與圖表匯出是否正常。測試 CSV 編碼時可使用 `test-data\csv-sales-sample.csv` 與 `test-data\big5-sales-sample.csv`。

## 系統需求

使用 portable exe：

- Windows 10 或 Windows 11。
- Microsoft Edge WebView2 Runtime。

多數 Windows 10/11 電腦已內建 WebView2 Runtime。如果程式無法開啟，可能需要安裝 Microsoft Edge WebView2 Runtime。

使用 portable 版本不需要安裝：

- Node.js
- npm
- Python
- pandas
- Microsoft Excel
- VBA 或 COM automation

## 錯誤記錄

程式執行時如果發生錯誤，會寫入執行檔旁的 log 檔：

```text
portableapps\logs\app.log
```

一般使用者不需要主動打開它；如果程式開不起來、檔案解析失敗或畫面發生異常，可以把這個 log 檔提供給維護者查看。

## 開發

安裝依賴：

```bat
npm install
```

啟動 Vite 前端開發伺服器：

```bat
npm run dev
```

啟動 Tauri 桌面開發模式：

```bat
npm run tauri dev
```

建置前端：

```bat
npm run build
```

執行基本檢查：

```bat
test.bat
```

`test.bat` 會在需要時安裝 npm 依賴，接著執行前端 build，最後在 `src-tauri` 執行 `cargo check`。

Windows 快捷指令：

```bat
test.bat dev      :: 啟動 Vite dev server
test.bat tauri    :: 啟動 Tauri 桌面應用
test.bat build    :: 呼叫 build.bat 建置 portable exe
test.bat help     :: 顯示說明
```

## 建置 portable exe

在專案根目錄執行：

```bat
build.bat
```

成功後會產生：

```text
portableapps\excel-visual-tool.exe
```

`portableapps` 是本機輸出資料夾，不應提交到 Git。執行程式後產生的 `portableapps\exports` 與 `portableapps\logs` 也屬於本機輸出資料。

## 專案結構

```text
src\                  React / TypeScript 前端介面
src-tauri\            Tauri / Rust 桌面程式
test-data\            測試用試算表檔案
rfp\                  規劃與參考資料
dist\                 前端建置輸出，不提交 Git
node_modules\         npm 依賴，不提交 Git
src-tauri\target\     Rust 建置輸出，不提交 Git
portableapps\         portable exe 與本機輸出，不提交 Git
```
