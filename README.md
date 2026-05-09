# Excel Visual Tool

Excel Visual Tool 是一個以 Windows 桌面使用情境為主的 Excel 視覺化工具原型。專案採用 Tauri + React + TypeScript + ECharts，目標是讓使用者在本機開啟 `.xlsx` 檔案、預覽資料、設定欄位映射，並產生互動式圖表。

目前專案仍在 MVP 初期：桌面應用外殼、React 版面與 ECharts 範例圖表已建立；Excel 檔案解析、欄位型態推斷、圖表匯出與本機狀態保存仍在待開發階段。

## 技術棧

- **Tauri 2**：提供 Windows 桌面應用外殼與打包能力。
- **React 18 + TypeScript**：建立前端工具介面。
- **Vite**：前端開發伺服器與建置工具。
- **ECharts**：互動式圖表渲染。
- **Rust**：Tauri 後端入口與原生能力擴充基礎。

第一版規劃避免依賴 Microsoft Excel、Python、pandas、VBA、COM automation 或雲端上傳，降低部署複雜度並維持本機處理資料的定位。

## 專案結構

```text
.
├── src/                  # React 前端程式
│   ├── main.tsx          # 主要 UI 與 ECharts 範例設定
│   └── styles.css        # 全域樣式與響應式版面
├── src-tauri/            # Tauri / Rust 桌面應用
│   ├── src/lib.rs        # Tauri Builder 與 plugin 初始化
│   ├── src/main.rs       # 桌面程式入口
│   ├── tauri.conf.json   # 視窗、建置與 bundle 設定
│   └── icons/            # Windows app icon
├── dist/                 # 前端建置輸出
├── test.bat              # Windows 開發與驗證捷徑
├── TODO.md               # MVP 待辦事項
└── idea.md               # 產品與架構規劃筆記
```

## 目前功能

- 顯示三欄式工具介面：左側檔案/工作表區、中間資料與圖表預覽、右側圖表設定。
- 內建 ECharts 長條圖範例，用於驗證圖表渲染流程。
- Tauri 視窗設定完成，預設尺寸為 `1200x780`，最小尺寸為 `960x640`。
- 可使用 Vite 在瀏覽器中開發，也可透過 Tauri 啟動桌面應用。

## 開發環境需求

- Node.js 與 npm
- Rust toolchain
- Windows 開發 Tauri 應用所需環境
- Microsoft WebView2 Runtime（多數 Windows 10/11 已內建；最終發佈仍需驗證）

## 安裝與執行

安裝依賴：

```bash
npm install
```

啟動前端開發伺服器：

```bash
npm run dev
```

啟動 Tauri 桌面應用：

```bash
npm run tauri dev
```

執行前端 TypeScript 檢查與建置：

```bash
npm run build
```

預覽前端 production build：

```bash
npm run preview
```

## Windows 捷徑指令

專案提供 `test.bat` 作為 Windows 常用流程入口：

```bat
test.bat         :: 安裝缺少的 npm 依賴、建置前端、執行 cargo check
test.bat dev     :: 啟動 Vite dev server
test.bat tauri   :: 啟動 Tauri 桌面應用
test.bat build   :: 建置 release portable exe（不產生 installer bundle）
```

`test.bat build` 預期輸出：

```text
src-tauri\target\release\excel-visual-tool.exe
```

## 建置與驗證

建議提交或交付前至少執行：

```bat
test.bat
```

此流程會：

1. 檢查 `node_modules`，缺少時執行 `npm install`。
2. 執行 `npm run build`。
3. 進入 `src-tauri/` 執行 `cargo check`。

目前尚未設定 Vitest、Playwright 或 Rust integration tests。新增功能時，應優先補上貼近功能邏輯的測試。

## MVP 規劃

短期目標集中在本機 `.xlsx` 視覺化流程：

- 開啟本機 `.xlsx` 檔案。
- 解析 workbook、工作表、headers 與 rows。
- 預覽前 200-500 列資料。
- 推斷欄位型態，例如文字、數值與日期。
- 設定 X 軸、Y 軸、分類欄位與圖表類型。
- 支援長條圖、折線圖、圓餅圖、散佈圖與堆疊長條圖。
- 匯出目前圖表為 PNG。
- 保存最近檔案與圖表設定到本機。

## 發佈方向

第一版建議以 Windows portable 版本為主要發佈形式，讓使用者不需要安裝 Node.js、Python 或前端套件即可執行。Installer 可作為後續選項，用於建立捷徑、自動處理 WebView2 Runtime 或加入更新流程。

## 開發注意事項

- 不要提交 `node_modules/`、`dist/`、`src-tauri/target/` 或 `.env`。
- 目前資料處理設計以本機為主，不應新增雲端上傳流程，除非產品需求明確改變。
- 若新增 Tauri plugin、檔案系統權限或安全設定，需同步更新 `src-tauri/tauri.conf.json` 並在變更說明中標註。
