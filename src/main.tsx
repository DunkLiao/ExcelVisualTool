import React, { ChangeEvent, useMemo, useRef, useState } from "react";
import ReactDOM from "react-dom/client";
import ReactECharts from "echarts-for-react";
import * as XLSX from "xlsx";
import "./styles.css";

const PREVIEW_ROW_LIMIT = 200;

type CellValue = string | number | boolean | Date | null;

type FieldType = "text" | "number" | "date" | "mixed" | "empty";

type ParsedSheet = {
  name: string;
  headers: string[];
  rows: Record<string, CellValue>[];
};

type FieldMetadata = {
  name: string;
  type: FieldType;
  emptyRatio: number;
};

type WorkbookState = {
  fileName: string;
  sheets: ParsedSheet[];
  activeSheetName: string;
};

const sampleChartOption = {
  tooltip: {
    trigger: "axis"
  },
  grid: {
    left: 40,
    right: 24,
    top: 36,
    bottom: 36
  },
  xAxis: {
    type: "category",
    data: ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
  },
  yAxis: {
    type: "value"
  },
  series: [
    {
      name: "Sales",
      type: "bar",
      data: [120, 180, 150, 240, 210, 280],
      itemStyle: {
        color: "#2563eb"
      }
    }
  ]
};

function formatCellValue(value: CellValue) {
  if (value === null || value === "") {
    return "";
  }

  if (value instanceof Date) {
    return value.toLocaleDateString();
  }

  return String(value);
}

function isEmptyCell(value: CellValue) {
  return value === null || value === "";
}

function isDateLike(value: CellValue) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return true;
  }

  if (typeof value !== "string") {
    return false;
  }

  const trimmed = value.trim();
  if (!trimmed) {
    return false;
  }

  const parsed = Date.parse(trimmed);
  return !Number.isNaN(parsed) && /\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(trimmed);
}

function inferFieldType(values: CellValue[]): FieldType {
  const presentValues = values.filter((value) => !isEmptyCell(value));

  if (presentValues.length === 0) {
    return "empty";
  }

  const numberCount = presentValues.filter(
    (value) =>
      typeof value === "number" ||
      (typeof value === "string" && value.trim() !== "" && !Number.isNaN(Number(value)))
  ).length;
  const dateCount = presentValues.filter(isDateLike).length;

  if (numberCount === presentValues.length) {
    return "number";
  }

  if (dateCount === presentValues.length) {
    return "date";
  }

  const textCount = presentValues.filter((value) => typeof value === "string").length;
  return textCount === presentValues.length ? "text" : "mixed";
}

function buildFieldMetadata(sheet: ParsedSheet): FieldMetadata[] {
  return sheet.headers.map((header) => {
    const values = sheet.rows.map((row) => row[header] ?? null);
    const emptyCount = values.filter(isEmptyCell).length;

    return {
      name: header,
      type: inferFieldType(values),
      emptyRatio: values.length === 0 ? 1 : emptyCount / values.length
    };
  });
}

function normalizeHeaders(headerRow: CellValue[], columnCount: number) {
  const usedNames = new Map<string, number>();

  return Array.from({ length: columnCount }, (_, index) => {
    const rawHeader = headerRow[index];
    const baseName =
      rawHeader === null || rawHeader === ""
        ? `Column ${index + 1}`
        : String(rawHeader).trim() || `Column ${index + 1}`;
    const existingCount = usedNames.get(baseName) ?? 0;
    usedNames.set(baseName, existingCount + 1);

    return existingCount === 0 ? baseName : `${baseName} ${existingCount + 1}`;
  });
}

function parseWorksheet(name: string, worksheet: XLSX.WorkSheet): ParsedSheet {
  const rawRows = XLSX.utils.sheet_to_json<CellValue[]>(worksheet, {
    header: 1,
    defval: null,
    raw: true
  });
  const columnCount = rawRows.reduce((count, row) => Math.max(count, row.length), 0);

  if (rawRows.length === 0 || columnCount === 0) {
    return {
      name,
      headers: [],
      rows: []
    };
  }

  const headers = normalizeHeaders(rawRows[0] ?? [], columnCount);
  const rows = rawRows.slice(1).map((rawRow) =>
    headers.reduce<Record<string, CellValue>>((row, header, index) => {
      row[header] = rawRow[index] ?? null;
      return row;
    }, {})
  );

  return {
    name,
    headers,
    rows
  };
}

function parseWorkbook(fileName: string, buffer: ArrayBuffer): WorkbookState {
  const workbook = XLSX.read(buffer, {
    cellDates: true,
    type: "array"
  });

  const sheets = workbook.SheetNames.map((sheetName) =>
    parseWorksheet(sheetName, workbook.Sheets[sheetName])
  );

  if (sheets.length === 0) {
    throw new Error("No worksheets found in this workbook.");
  }

  return {
    fileName,
    sheets,
    activeSheetName: sheets[0].name
  };
}

function App() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [workbookState, setWorkbookState] = useState<WorkbookState | null>(null);
  const [errorMessage, setErrorMessage] = useState("");
  const [xField, setXField] = useState("");
  const [yField, setYField] = useState("");

  const activeSheet = useMemo(() => {
    return (
      workbookState?.sheets.find((sheet) => sheet.name === workbookState.activeSheetName) ?? null
    );
  }, [workbookState]);

  const previewRows = activeSheet?.rows.slice(0, PREVIEW_ROW_LIMIT) ?? [];
  const fieldMetadata = useMemo(() => {
    return activeSheet ? buildFieldMetadata(activeSheet) : [];
  }, [activeSheet]);

  function resetChartFields() {
    setXField("");
    setYField("");
  }

  async function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setErrorMessage("Only .xlsx files are supported in this version.");
      setWorkbookState(null);
      resetChartFields();
      event.target.value = "";
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const parsedWorkbook = parseWorkbook(file.name, buffer);
      setWorkbookState(parsedWorkbook);
      setErrorMessage("");
      resetChartFields();
    } catch (error) {
      setWorkbookState(null);
      resetChartFields();
      setErrorMessage(error instanceof Error ? error.message : "Unable to parse this workbook.");
    } finally {
      event.target.value = "";
    }
  }

  function handleSheetChange(sheetName: string) {
    setWorkbookState((currentState) =>
      currentState
        ? {
            ...currentState,
            activeSheetName: sheetName
          }
        : currentState
    );
    resetChartFields();
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div>
          <h1>Excel Visual Tool</h1>
          <p>{workbookState?.fileName ?? "Local workbook visualization workspace."}</p>
        </div>
        <input
          ref={fileInputRef}
          className="file-input"
          type="file"
          accept=".xlsx"
          onChange={handleFileChange}
        />
        <button
          type="button"
          className="primary-button"
          onClick={() => fileInputRef.current?.click()}
        >
          Open .xlsx
        </button>
        {errorMessage ? <div className="error-message">{errorMessage}</div> : null}
        <section className="panel">
          <h2>Worksheets</h2>
          {workbookState ? (
            <div className="worksheet-list">
              {workbookState.sheets.map((sheet) => (
                <button
                  key={sheet.name}
                  type="button"
                  className={
                    sheet.name === workbookState.activeSheetName
                      ? "worksheet-button active"
                      : "worksheet-button"
                  }
                  onClick={() => handleSheetChange(sheet.name)}
                >
                  <span>{sheet.name}</span>
                  <small>{sheet.rows.length} rows</small>
                </button>
              ))}
            </div>
          ) : (
            <div className="empty-state">No workbook selected</div>
          )}
        </section>
      </aside>

      <section className="workspace">
        <div className="preview-area">
          <div className="section-header">
            <h2>Data Preview</h2>
            <span>
              {activeSheet
                ? `${Math.min(activeSheet.rows.length, PREVIEW_ROW_LIMIT)} of ${
                    activeSheet.rows.length
                  } rows`
                : `First ${PREVIEW_ROW_LIMIT} rows`}
            </span>
          </div>
          {activeSheet && activeSheet.headers.length > 0 ? (
            <div className="table-scroll">
              <table className="data-table">
                <thead>
                  <tr>
                    {activeSheet.headers.map((header) => (
                      <th key={header}>{header}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {previewRows.map((row, rowIndex) => (
                    <tr key={`${activeSheet.name}-${rowIndex}`}>
                      {activeSheet.headers.map((header) => (
                        <td key={header}>{formatCellValue(row[header] ?? null)}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {activeSheet.rows.length === 0 ? (
                <div className="table-empty-note">This worksheet has headers but no data rows.</div>
              ) : null}
            </div>
          ) : (
            <div className="table-placeholder">
              <span>
                {activeSheet
                  ? "This worksheet is empty."
                  : "Select an Excel file to preview worksheet rows."}
              </span>
            </div>
          )}
        </div>

        <div className="chart-area">
          <div className="section-header">
            <h2>Chart Preview</h2>
            <span>ECharts ready</span>
          </div>
          <ReactECharts option={sampleChartOption} className="chart" />
        </div>
      </section>

      <aside className="settings">
        <h2>Chart Settings</h2>
        <label>
          Chart type
          <select defaultValue="bar">
            <option value="bar">Bar</option>
            <option value="line">Line</option>
            <option value="pie">Pie</option>
            <option value="scatter">Scatter</option>
          </select>
        </label>
        <label>
          X axis
          <select
            disabled={fieldMetadata.length === 0}
            value={xField}
            onChange={(event) => setXField(event.target.value)}
          >
            <option value="">Select field</option>
            {fieldMetadata.map((field) => (
              <option key={field.name} value={field.name}>
                {field.name}
              </option>
            ))}
          </select>
        </label>
        <label>
          Y axis
          <select
            disabled={fieldMetadata.length === 0}
            value={yField}
            onChange={(event) => setYField(event.target.value)}
          >
            <option value="">Select field</option>
            {fieldMetadata.map((field) => (
              <option key={field.name} value={field.name}>
                {field.name}
              </option>
            ))}
          </select>
        </label>
        <section className="metadata-panel">
          <h2>Fields</h2>
          {fieldMetadata.length > 0 ? (
            <div className="field-list">
              {fieldMetadata.map((field) => (
                <div className="field-row" key={field.name}>
                  <span>{field.name}</span>
                  <small>
                    {field.type} · {Math.round(field.emptyRatio * 100)}% empty
                  </small>
                </div>
              ))}
            </div>
          ) : (
            <div className="empty-state compact">No fields available</div>
          )}
        </section>
      </aside>
    </main>
  );
}

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
