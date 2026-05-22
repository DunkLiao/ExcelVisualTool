import React, { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import ReactDOM from "react-dom/client";
import { invoke } from "@tauri-apps/api/core";
import ReactECharts from "echarts-for-react";
import * as XLSX from "xlsx";
import "./styles.css";

const PREVIEW_ROW_LIMIT = 200;
const SETTINGS_STORAGE_KEY = "excel-visual-tool.chart-settings.v1";
const RECENT_FILE_STORAGE_KEY = "excel-visual-tool.recent-file.v1";
const DAY_IN_MS = 24 * 60 * 60 * 1000;

type CellValue = string | number | boolean | Date | null;

type FieldType = "text" | "number" | "date" | "mixed" | "empty";
type ChartType = "bar" | "line" | "pie" | "scatter" | "stackedBar" | "timeSeries";

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

type ChartSettings = {
  chartType: ChartType;
  xField: string;
  yField: string;
  categoryField: string;
};

type ChartValidation = {
  valid: boolean;
  message: string;
  warnings: string[];
};

type ChartOption = Record<string, unknown>;
type TimeAxisGranularity = "year" | "month" | "day";
type TooltipParam = {
  axisValue?: string | number;
  marker?: string;
  seriesName?: string;
  value?: unknown;
};

const DEFAULT_CHART_SETTINGS: ChartSettings = {
  chartType: "bar",
  xField: "",
  yField: "",
  categoryField: ""
};

function describeError(error: unknown) {
  if (error instanceof Error) {
    return error.stack || error.message;
  }

  return String(error);
}

function writeAppLog(level: "info" | "warn" | "error", message: string) {
  void invoke("write_app_log", { level, message }).catch(() => {
    // Logging must never interrupt the UI.
  });
}

function installFrontendLogging() {
  window.addEventListener("error", (event) => {
    writeAppLog("error", `Frontend error: ${event.message} at ${event.filename}:${event.lineno}`);
  });

  window.addEventListener("unhandledrejection", (event) => {
    writeAppLog("error", `Unhandled promise rejection: ${describeError(event.reason)}`);
  });
}

const chartTypeLabels: Record<ChartType, string> = {
  bar: "Bar",
  line: "Line",
  pie: "Pie",
  scatter: "Scatter",
  stackedBar: "Stacked bar",
  timeSeries: "Time series"
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

function formatDimensionValue(value: CellValue) {
  const formattedValue = formatCellValue(value);
  return formattedValue === "" ? "(blank)" : formattedValue;
}

function isEmptyCell(value: CellValue) {
  return value === null || value === "";
}

function coerceNumber(value: CellValue) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string" && value.trim() !== "") {
    const parsedValue = Number(value);
    return Number.isFinite(parsedValue) ? parsedValue : null;
  }

  return null;
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

  return parseDateTime(trimmed) !== null;
}

function parseDateTime(value: string) {
  const dateOnlyMatch = value.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (dateOnlyMatch) {
    const [, year, month, day] = dateOnlyMatch;
    const parsedDate = new Date(Number(year), Number(month) - 1, Number(day));
    return Number.isNaN(parsedDate.getTime()) ? null : parsedDate.getTime();
  }

  if (!/\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(value)) {
    return null;
  }

  const parsed = Date.parse(value);
  return Number.isNaN(parsed) ? null : parsed;
}

function coerceDateTime(value: CellValue) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.getTime();
  }

  if (typeof value !== "string") {
    return null;
  }

  const trimmed = value.trim();
  if (!trimmed || !/\d{1,4}[-/]\d{1,2}[-/]\d{1,4}/.test(trimmed)) {
    return null;
  }

  return parseDateTime(trimmed);
}

function formatDateLabel(value: string | number | Date, granularity: TimeAxisGranularity = "day") {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "";
  }

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  if (granularity === "year") {
    return String(year);
  }

  if (granularity === "month") {
    return `${year}-${month}`;
  }

  return `${year}-${month}-${day}`;
}

function getTimeAxisGranularity(timestamps: number[]): TimeAxisGranularity {
  if (timestamps.length < 2) {
    return "day";
  }

  const minTime = Math.min(...timestamps);
  const maxTime = Math.max(...timestamps);
  const rangeInDays = (maxTime - minTime) / DAY_IN_MS;

  if (rangeInDays > 730) {
    return "year";
  }

  if (rangeInDays > 90) {
    return "month";
  }

  return "day";
}

function formatTimeSeriesTooltip(params: TooltipParam | TooltipParam[]) {
  const pointParams = Array.isArray(params) ? params : [params];
  const firstPoint = pointParams[0];
  const firstValue = Array.isArray(firstPoint?.value) ? firstPoint.value : [];
  const rawDate = firstValue[0] ?? firstPoint?.axisValue ?? "";
  const lines = pointParams.map((point) => {
    const pointValue = Array.isArray(point.value) ? point.value[1] : "";
    const marker = point.marker ?? "";
    const seriesName = point.seriesName ?? "";
    return `${marker}${seriesName}<span style="float:right;margin-left:16px;font-weight:600">${pointValue}</span>`;
  });

  return [formatDateLabel(rawDate as string | number), ...lines].join("<br/>");
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

function loadChartSettings(): ChartSettings {
  try {
    const storedSettings = window.localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!storedSettings) {
      return DEFAULT_CHART_SETTINGS;
    }

    const parsedSettings = JSON.parse(storedSettings) as Partial<ChartSettings>;
    const chartType =
      parsedSettings.chartType && parsedSettings.chartType in chartTypeLabels
        ? parsedSettings.chartType
        : DEFAULT_CHART_SETTINGS.chartType;

    return {
      chartType,
      xField: parsedSettings.xField ?? "",
      yField: parsedSettings.yField ?? "",
      categoryField: parsedSettings.categoryField ?? ""
    };
  } catch {
    return DEFAULT_CHART_SETTINGS;
  }
}

function fieldExists(metadata: FieldMetadata[], fieldName: string) {
  return metadata.some((field) => field.name === fieldName);
}

function getFieldMetadata(metadata: FieldMetadata[], fieldName: string) {
  return metadata.find((field) => field.name === fieldName) ?? null;
}

function validateDimensionField(
  metadata: FieldMetadata[],
  fieldName: string,
  label: string
): string | null {
  if (!fieldName) {
    return `Select a ${label} field.`;
  }

  const field = getFieldMetadata(metadata, fieldName);
  if (!field) {
    return `${label} field is not available in this worksheet.`;
  }

  if (field.type === "empty") {
    return `${label} field cannot be empty.`;
  }

  return null;
}

function validateNumericField(
  metadata: FieldMetadata[],
  fieldName: string,
  label: string
): string | null {
  const dimensionError = validateDimensionField(metadata, fieldName, label);
  if (dimensionError) {
    return dimensionError;
  }

  const field = getFieldMetadata(metadata, fieldName);
  if (field?.type !== "number") {
    return `${label} field must be numeric.`;
  }

  return null;
}

function validateDateField(
  metadata: FieldMetadata[],
  fieldName: string,
  label: string
): string | null {
  const dimensionError = validateDimensionField(metadata, fieldName, label);
  if (dimensionError) {
    return dimensionError;
  }

  const field = getFieldMetadata(metadata, fieldName);
  if (field?.type !== "date") {
    return `${label} field must be date/time.`;
  }

  return null;
}

function buildChartValidation(
  sheet: ParsedSheet | null,
  metadata: FieldMetadata[],
  settings: ChartSettings
): ChartValidation {
  if (!sheet) {
    return {
      valid: false,
      message: "Open an .xlsx workbook to build a chart.",
      warnings: []
    };
  }

  if (sheet.headers.length === 0 || sheet.rows.length === 0) {
    return {
      valid: false,
      message: "This worksheet has no rows available for charting.",
      warnings: []
    };
  }

  const warnings: string[] = [];
  const yLabel = settings.chartType === "scatter" ? "Y axis" : "value";
  const xError =
    settings.chartType === "scatter"
      ? validateNumericField(metadata, settings.xField, "X axis")
      : settings.chartType === "timeSeries"
        ? validateDateField(metadata, settings.xField, "X axis")
      : validateDimensionField(metadata, settings.xField, "X axis");
  const yError = validateNumericField(metadata, settings.yField, yLabel);

  if (xError || yError) {
    return {
      valid: false,
      message: xError ?? yError ?? "Select valid chart fields.",
      warnings
    };
  }

  if (settings.chartType === "stackedBar") {
    const categoryError = validateDimensionField(metadata, settings.categoryField, "category");
    if (categoryError) {
      return {
        valid: false,
        message: categoryError,
        warnings
      };
    }
  }

  [settings.xField, settings.yField, settings.categoryField]
    .filter(Boolean)
    .forEach((fieldName) => {
      const field = getFieldMetadata(metadata, fieldName);
      if (field && field.emptyRatio > 0) {
        warnings.push(`${field.name} has ${Math.round(field.emptyRatio * 100)}% empty cells.`);
      }
    });

  return {
    valid: true,
    message: "",
    warnings
  };
}

function createBaseChartOption(settings: ChartSettings): ChartOption {
  return {
    color: ["#2563eb", "#0891b2", "#16a34a", "#f59e0b", "#dc2626", "#7c3aed", "#475569"],
    tooltip: {
      trigger: settings.chartType === "pie" ? "item" : "axis"
    },
    legend: {
      type: "scroll",
      top: 8
    },
    grid: {
      left: 52,
      right: 28,
      top: 56,
      bottom: 42,
      containLabel: true
    }
  };
}

function aggregateRows(sheet: ParsedSheet, settings: ChartSettings) {
  const xValues: string[] = [];
  const seriesValues = new Map<string, Map<string, number>>();
  const seriesField = settings.categoryField;

  sheet.rows.forEach((row) => {
    const rawXValue = row[settings.xField] ?? null;
    const numericValue = coerceNumber(row[settings.yField] ?? null);
    if (numericValue === null || isEmptyCell(rawXValue)) {
      return;
    }

    const xValue = formatDimensionValue(rawXValue);
    const seriesName = seriesField
      ? formatDimensionValue(row[seriesField] ?? null)
      : settings.yField;

    if (!xValues.includes(xValue)) {
      xValues.push(xValue);
    }

    if (!seriesValues.has(seriesName)) {
      seriesValues.set(seriesName, new Map<string, number>());
    }

    const currentSeries = seriesValues.get(seriesName);
    currentSeries?.set(xValue, (currentSeries.get(xValue) ?? 0) + numericValue);
  });

  return {
    xValues,
    seriesValues
  };
}

function aggregateTimeSeriesRows(sheet: ParsedSheet, settings: ChartSettings) {
  const seriesValues = new Map<string, Map<number, number>>();
  const seriesField = settings.categoryField;

  sheet.rows.forEach((row) => {
    const xValue = coerceDateTime(row[settings.xField] ?? null);
    const yValue = coerceNumber(row[settings.yField] ?? null);
    if (xValue === null || yValue === null) {
      return;
    }

    const seriesName = seriesField
      ? formatDimensionValue(row[seriesField] ?? null)
      : settings.yField;
    if (!seriesValues.has(seriesName)) {
      seriesValues.set(seriesName, new Map<number, number>());
    }

    const currentSeries = seriesValues.get(seriesName);
    currentSeries?.set(xValue, (currentSeries.get(xValue) ?? 0) + yValue);
  });

  return seriesValues;
}

function buildChartOption(
  sheet: ParsedSheet | null,
  settings: ChartSettings,
  validation: ChartValidation
): ChartOption {
  if (!sheet || !validation.valid) {
    return {
      title: {
        text: validation.message,
        left: "center",
        top: "middle",
        textStyle: {
          color: "#657084",
          fontSize: 14,
          fontWeight: 500
        }
      }
    };
  }

  const baseOption = createBaseChartOption(settings);

  if (settings.chartType === "pie") {
    const { xValues, seriesValues } = aggregateRows(sheet, settings);
    const values = seriesValues.get(settings.yField) ?? new Map<string, number>();
    return {
      ...baseOption,
      tooltip: {
        trigger: "item"
      },
      series: [
        {
          name: settings.yField,
          type: "pie",
          radius: ["35%", "68%"],
          data: xValues.map((xValue) => ({
            name: xValue,
            value: values.get(xValue) ?? 0
          }))
        }
      ]
    };
  }

  if (settings.chartType === "scatter") {
    const seriesField = settings.categoryField;
    const groupedPoints = new Map<string, number[][]>();

    sheet.rows.forEach((row) => {
      const xValue = coerceNumber(row[settings.xField] ?? null);
      const yValue = coerceNumber(row[settings.yField] ?? null);
      if (xValue === null || yValue === null) {
        return;
      }

    const seriesName = seriesField
      ? formatDimensionValue(row[seriesField] ?? null)
      : settings.yField;
      groupedPoints.set(seriesName, [...(groupedPoints.get(seriesName) ?? []), [xValue, yValue]]);
    });

    return {
      ...baseOption,
      xAxis: {
        type: "value",
        name: settings.xField
      },
      yAxis: {
        type: "value",
        name: settings.yField
      },
      series: Array.from(groupedPoints.entries()).map(([seriesName, data]) => ({
        name: seriesName,
        type: "scatter",
        data,
        symbolSize: 8
      }))
    };
  }

  if (settings.chartType === "timeSeries") {
    const seriesValues = aggregateTimeSeriesRows(sheet, settings);
    const timestamps = Array.from(seriesValues.values()).flatMap((values) =>
      Array.from(values.keys())
    );
    const axisGranularity = getTimeAxisGranularity(timestamps);

    return {
      ...baseOption,
      tooltip: {
        trigger: "axis",
        formatter: formatTimeSeriesTooltip
      },
      grid: {
        left: 52,
        right: 56,
        top: 56,
        bottom: 72,
        containLabel: true
      },
      xAxis: {
        type: "time",
        name: settings.xField,
        nameLocation: "middle",
        nameGap: 48,
        axisLabel: {
          hideOverlap: true,
          formatter: (value: number) => formatDateLabel(value, axisGranularity)
        }
      },
      yAxis: {
        type: "value",
        name: settings.yField
      },
      series: Array.from(seriesValues.entries()).map(([seriesName, values]) => ({
        name: seriesName,
        type: "line",
        smooth: true,
        data: Array.from(values.entries())
          .sort(([leftDate], [rightDate]) => leftDate - rightDate)
          .map(([xValue, yValue]) => [xValue, yValue])
      }))
    };
  }

  const { xValues, seriesValues } = aggregateRows(sheet, settings);
  const isStacked = settings.chartType === "stackedBar";
  const seriesType = settings.chartType === "line" ? "line" : "bar";

  return {
    ...baseOption,
    xAxis: {
      type: "category",
      data: xValues
    },
    yAxis: {
      type: "value",
      name: settings.yField
    },
    series: Array.from(seriesValues.entries()).map(([seriesName, values]) => ({
      name: seriesName,
      type: seriesType,
      stack: isStacked ? "total" : undefined,
      smooth: settings.chartType === "line",
      data: xValues.map((xValue) => values.get(xValue) ?? 0)
    }))
  };
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
  const chartRef = useRef<ReactECharts>(null);
  const [workbookState, setWorkbookState] = useState<WorkbookState | null>(null);
  const [errorMessage, setErrorMessage] = useState("");
  const [chartSettings, setChartSettings] = useState<ChartSettings>(loadChartSettings);
  const [recentFileName, setRecentFileName] = useState(
    () => window.localStorage.getItem(RECENT_FILE_STORAGE_KEY) ?? ""
  );

  const activeSheet = useMemo(() => {
    return (
      workbookState?.sheets.find((sheet) => sheet.name === workbookState.activeSheetName) ?? null
    );
  }, [workbookState]);

  const previewRows = activeSheet?.rows.slice(0, PREVIEW_ROW_LIMIT) ?? [];
  const fieldMetadata = useMemo(() => {
    return activeSheet ? buildFieldMetadata(activeSheet) : [];
  }, [activeSheet]);
  const chartValidation = useMemo(() => {
    return buildChartValidation(activeSheet, fieldMetadata, chartSettings);
  }, [activeSheet, fieldMetadata, chartSettings]);
  const chartOption = useMemo(() => {
    return buildChartOption(activeSheet, chartSettings, chartValidation);
  }, [activeSheet, chartSettings, chartValidation]);
  const hasRestoredMissingFields = useMemo(() => {
    if (!activeSheet) {
      return false;
    }

    return [chartSettings.xField, chartSettings.yField, chartSettings.categoryField]
      .filter(Boolean)
      .some((fieldName) => !fieldExists(fieldMetadata, fieldName));
  }, [activeSheet, chartSettings, fieldMetadata]);

  useEffect(() => {
    window.localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(chartSettings));
  }, [chartSettings]);

  function resetChartFields() {
    setChartSettings((currentSettings) => ({
      ...currentSettings,
      xField: "",
      yField: "",
      categoryField: ""
    }));
  }

  function updateChartSettings(nextSettings: Partial<ChartSettings>) {
    setChartSettings((currentSettings) => ({
      ...currentSettings,
      ...nextSettings
    }));
  }

  function handleExportPng() {
    if (!chartValidation.valid) {
      return;
    }

    const chartInstance = chartRef.current?.getEchartsInstance();
    if (!chartInstance) {
      return;
    }

    const imageUrl = chartInstance.getDataURL({
      type: "png",
      pixelRatio: 2,
      backgroundColor: "#ffffff"
    });
    const downloadLink = document.createElement("a");
    const fileBaseName = workbookState?.fileName.replace(/\.xlsx$/i, "") || "chart";
    downloadLink.href = imageUrl;
    downloadLink.download = `${fileBaseName}-${chartSettings.chartType}.png`;
    downloadLink.click();
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
      setRecentFileName(file.name);
      window.localStorage.setItem(RECENT_FILE_STORAGE_KEY, file.name);
    } catch (error) {
      setWorkbookState(null);
      resetChartFields();
      writeAppLog("error", `Unable to parse workbook ${file.name}: ${describeError(error)}`);
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
        {!workbookState && recentFileName ? (
          <div className="info-message">Recent file: {recentFileName}</div>
        ) : null}
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
            <span>
              {chartValidation.valid
                ? chartTypeLabels[chartSettings.chartType]
                : "Waiting for valid settings"}
            </span>
          </div>
          <ReactECharts ref={chartRef} option={chartOption} className="chart" notMerge />
          {chartValidation.message ? (
            <div className="chart-message error">{chartValidation.message}</div>
          ) : null}
          {chartValidation.warnings.length > 0 ? (
            <div className="chart-message warning">{chartValidation.warnings.join(" ")}</div>
          ) : null}
        </div>
      </section>

      <aside className="settings">
        <h2>Chart Settings</h2>
        <label>
          Chart type
          <select
            value={chartSettings.chartType}
            onChange={(event) =>
              updateChartSettings({
                chartType: event.target.value as ChartType
              })
            }
          >
            <option value="bar">Bar</option>
            <option value="line">Line</option>
            <option value="pie">Pie</option>
            <option value="scatter">Scatter</option>
            <option value="stackedBar">Stacked bar</option>
            <option value="timeSeries">Time series</option>
          </select>
        </label>
        <label>
          X axis
          <select
            disabled={fieldMetadata.length === 0}
            value={fieldExists(fieldMetadata, chartSettings.xField) ? chartSettings.xField : ""}
            onChange={(event) => updateChartSettings({ xField: event.target.value })}
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
            value={fieldExists(fieldMetadata, chartSettings.yField) ? chartSettings.yField : ""}
            onChange={(event) => updateChartSettings({ yField: event.target.value })}
          >
            <option value="">Select field</option>
            {fieldMetadata.map((field) => (
              <option key={field.name} value={field.name}>
                {field.name} {field.type === "number" ? "" : `(${field.type})`}
              </option>
            ))}
          </select>
        </label>
        <label>
          Category
          <select
            disabled={fieldMetadata.length === 0}
            value={
              fieldExists(fieldMetadata, chartSettings.categoryField)
                ? chartSettings.categoryField
                : ""
            }
            onChange={(event) => updateChartSettings({ categoryField: event.target.value })}
          >
            <option value="">
              {chartSettings.chartType === "stackedBar" ? "Select field" : "None"}
            </option>
            {fieldMetadata.map((field) => (
              <option key={field.name} value={field.name}>
                {field.name}
              </option>
            ))}
          </select>
        </label>
        {hasRestoredMissingFields ? (
          <div className="info-message">
            Saved settings include fields that are not in this worksheet.
          </div>
        ) : null}
        {chartValidation.message ? (
          <div className="error-message">{chartValidation.message}</div>
        ) : null}
        {chartValidation.warnings.length > 0 ? (
          <div className="warning-message">{chartValidation.warnings.join(" ")}</div>
        ) : null}
        <button
          type="button"
          className="secondary-button"
          disabled={!chartValidation.valid}
          onClick={handleExportPng}
        >
          Export PNG
        </button>
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

installFrontendLogging();

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
