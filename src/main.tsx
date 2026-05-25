import React, { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import ReactDOM from "react-dom/client";
import { invoke } from "@tauri-apps/api/core";
import ReactECharts from "echarts-for-react";
import initSqlJs from "sql.js";
import type { Database, QueryExecResult, SqlJsStatic, SqlValue } from "sql.js";
import sqlWasmUrl from "sql.js/dist/sql-wasm.wasm?url";
import * as XLSX from "xlsx";
import "./styles.css";

const PREVIEW_ROW_LIMIT = 5;
const SETTINGS_STORAGE_KEY = "excel-visual-tool.chart-settings.v1";
const RECENT_FILE_STORAGE_KEY = "excel-visual-tool.recent-file.v1";
const DAY_IN_MS = 24 * 60 * 60 * 1000;
const SUPPORTED_FILE_EXTENSIONS = [".xlsx", ".xls", ".csv"];
const SUPPORTED_FILE_LABEL = ".xlsx、.xls 或 .csv";
const CSV_DECODER_CANDIDATES = ["utf-8", "big5", "utf-16le", "utf-16be"];
const DEFAULT_SQL_QUERY = "SELECT * FROM data";
const SQL_TABLE_NAME = "data";

type CellValue = string | number | boolean | Date | null;

type FieldType = "text" | "number" | "date" | "mixed" | "empty";
type ChartType =
  | "bar"
  | "line"
  | "pie"
  | "scatter"
  | "stackedBar"
  | "timeSeries"
  | "area"
  | "horizontalBar"
  | "radar"
  | "treemap"
  | "funnel"
  | "gauge";

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
type SaveChartPngResult = string;
type SaveQueryXlsxResult = string;
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
  bar: "長條圖",
  line: "折線圖",
  pie: "圓餅圖",
  scatter: "散佈圖",
  stackedBar: "堆疊長條圖",
  timeSeries: "時間序列圖",
  area: "面積圖",
  horizontalBar: "水平長條圖",
  radar: "雷達圖",
  treemap: "矩形式樹狀圖",
  funnel: "漏斗圖",
  gauge: "儀表圖"
};

const fieldTypeLabels: Record<FieldType, string> = {
  text: "文字",
  number: "數字",
  date: "日期",
  mixed: "混合",
  empty: "空白"
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
  return formattedValue === "" ? "(空白)" : formattedValue;
}

function formatTsvCell(value: CellValue | string) {
  return formatCellValue(value).replace(/[\t\r\n]+/g, " ");
}

function buildSheetTsv(sheet: ParsedSheet) {
  const headerRow = sheet.headers.map(formatTsvCell).join("\t");
  const dataRows = sheet.rows.map((row) =>
    sheet.headers.map((header) => formatTsvCell(row[header] ?? null)).join("\t")
  );

  return [headerRow, ...dataRows].join("\n");
}

function sanitizeWorksheetName(name: string) {
  const sanitizedName = name.replace(/[:\\/?*[\]]/g, " ").trim();
  return (sanitizedName || "查詢結果").slice(0, 31);
}

function buildSheetWorkbookBytes(sheet: ParsedSheet) {
  const rows = [
    sheet.headers,
    ...sheet.rows.map((row) => sheet.headers.map((header) => row[header] ?? null))
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, sanitizeWorksheetName(sheet.name));
  const workbookBuffer = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array"
  }) as ArrayBuffer;

  return Array.from(new Uint8Array(workbookBuffer));
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

function shouldUseItemTooltip(chartType: ChartType) {
  return ["pie", "treemap", "funnel", "gauge"].includes(chartType);
}

function shouldUseCategorySeries(chartType: ChartType) {
  return ["bar", "line", "stackedBar", "timeSeries", "area", "horizontalBar", "radar"].includes(
    chartType
  );
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
    return `請選擇${label}欄位。`;
  }

  const field = getFieldMetadata(metadata, fieldName);
  if (!field) {
    return `${label}欄位不存在於目前工作表。`;
  }

  if (field.type === "empty") {
    return `${label}欄位不能為空白。`;
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
    return `${label}欄位必須是數值。`;
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
    return `${label}欄位必須是日期或時間。`;
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
      message: "請開啟支援的試算表檔案以建立圖表。",
      warnings: []
    };
  }

  if (sheet.headers.length === 0 || sheet.rows.length === 0) {
    return {
      valid: false,
      message: "此工作表沒有可用於繪製圖表的資料列。",
      warnings: []
    };
  }

  const warnings: string[] = [];
  const yLabel = settings.chartType === "scatter" ? "Y 軸" : "數值";
  const xError =
    settings.chartType === "scatter"
      ? validateNumericField(metadata, settings.xField, "X 軸")
      : settings.chartType === "timeSeries"
        ? validateDateField(metadata, settings.xField, "X 軸")
      : validateDimensionField(metadata, settings.xField, "X 軸");
  const yError = validateNumericField(metadata, settings.yField, yLabel);

  if (xError || yError) {
    return {
      valid: false,
      message: xError ?? yError ?? "請選擇有效的圖表欄位。",
      warnings
    };
  }

  if (settings.chartType === "stackedBar") {
    const categoryError = validateDimensionField(metadata, settings.categoryField, "分類");
    if (categoryError) {
      return {
        valid: false,
        message: categoryError,
        warnings
      };
    }
  }

  if (settings.chartType === "radar" && settings.categoryField) {
    const categoryError = validateDimensionField(metadata, settings.categoryField, "分類");
    if (categoryError) {
      return {
        valid: false,
        message: categoryError,
        warnings
      };
    }
  }

  const warningFields = [
    settings.xField,
    settings.yField,
    shouldUseCategorySeries(settings.chartType) || settings.chartType === "scatter"
      ? settings.categoryField
      : ""
  ];

  warningFields
    .filter(Boolean)
    .forEach((fieldName) => {
      const field = getFieldMetadata(metadata, fieldName);
      if (field && field.emptyRatio > 0) {
        warnings.push(`${field.name} 有 ${Math.round(field.emptyRatio * 100)}% 的空白儲存格。`);
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
      trigger: shouldUseItemTooltip(settings.chartType) ? "item" : "axis"
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
  const seriesField = shouldUseCategorySeries(settings.chartType) ? settings.categoryField : "";

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

function sumSeriesValues(values: Map<string, number>) {
  return Array.from(values.values()).reduce((sum, value) => sum + value, 0);
}

function getMaxSeriesValue(seriesValues: Map<string, Map<string, number>>) {
  return Math.max(
    0,
    ...Array.from(seriesValues.values()).flatMap((values) => Array.from(values.values()))
  );
}

function aggregateTimeSeriesRows(sheet: ParsedSheet, settings: ChartSettings) {
  const seriesValues = new Map<string, Map<number, number>>();
  const seriesField = shouldUseCategorySeries(settings.chartType) ? settings.categoryField : "";

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

  if (settings.chartType === "treemap") {
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
          type: "treemap",
          roam: false,
          breadcrumb: {
            show: false
          },
          data: xValues.map((xValue) => ({
            name: xValue,
            value: values.get(xValue) ?? 0
          }))
        }
      ]
    };
  }

  if (settings.chartType === "funnel") {
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
          type: "funnel",
          left: "10%",
          top: 64,
          bottom: 24,
          width: "80%",
          sort: "descending",
          data: xValues.map((xValue) => ({
            name: xValue,
            value: values.get(xValue) ?? 0
          }))
        }
      ]
    };
  }

  if (settings.chartType === "gauge") {
    const { seriesValues } = aggregateRows(sheet, settings);
    const value = sumSeriesValues(seriesValues.get(settings.yField) ?? new Map<string, number>());
    const maxValue = value <= 100 ? 100 : Math.ceil((value * 1.2) / 10) * 10;
    return {
      ...baseOption,
      tooltip: {
        trigger: "item"
      },
      series: [
        {
          name: settings.yField,
          type: "gauge",
          min: 0,
          max: maxValue,
          progress: {
            show: true
          },
          detail: {
            valueAnimation: true,
            formatter: "{value}"
          },
          data: [
            {
              name: settings.yField,
              value
            }
          ]
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
  if (settings.chartType === "radar") {
    const maxValue = getMaxSeriesValue(seriesValues);
    return {
      ...baseOption,
      tooltip: {
        trigger: "item"
      },
      radar: {
        indicator: xValues.map((xValue) => ({
          name: xValue,
          max: maxValue <= 0 ? 100 : Math.ceil(maxValue * 1.2)
        }))
      },
      series: [
        {
          name: settings.yField,
          type: "radar",
          data: Array.from(seriesValues.entries()).map(([seriesName, values]) => ({
            name: seriesName,
            value: xValues.map((xValue) => values.get(xValue) ?? 0)
          }))
        }
      ]
    };
  }

  if (settings.chartType === "horizontalBar") {
    return {
      ...baseOption,
      grid: {
        left: 96,
        right: 28,
        top: 56,
        bottom: 42,
        containLabel: true
      },
      xAxis: {
        type: "value",
        name: settings.yField
      },
      yAxis: {
        type: "category",
        data: xValues
      },
      series: Array.from(seriesValues.entries()).map(([seriesName, values]) => ({
        name: seriesName,
        type: "bar",
        data: xValues.map((xValue) => values.get(xValue) ?? 0)
      }))
    };
  }

  const isStacked = settings.chartType === "stackedBar";
  const seriesType = settings.chartType === "line" || settings.chartType === "area" ? "line" : "bar";

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
      smooth: settings.chartType === "line" || settings.chartType === "area",
      areaStyle: settings.chartType === "area" ? {} : undefined,
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
        ? `欄 ${index + 1}`
        : String(rawHeader).trim() || `欄 ${index + 1}`;
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

function isCsvFile(fileName: string) {
  return fileName.toLowerCase().endsWith(".csv");
}

function countMatches(value: string, pattern: RegExp) {
  return value.match(pattern)?.length ?? 0;
}

function scoreDecodedCsv(value: string) {
  const replacementCount = countMatches(value, /\uFFFD/g);
  const controlCount = countMatches(value, /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g);
  const traditionalChineseCount = countMatches(value, /[\u4E00-\u9FFF]/g);
  const separatorCount = countMatches(value, /[,;\t]/g);
  const lineBreakCount = countMatches(value, /\r\n|\n|\r/g);

  return (
    traditionalChineseCount * 3 +
    separatorCount +
    lineBreakCount -
    replacementCount * 20 -
    controlCount * 10
  );
}

function decodeCsvBuffer(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer);

  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder("utf-8").decode(bytes.subarray(3));
  }

  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder("utf-16le").decode(bytes.subarray(2));
  }

  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder("utf-16be").decode(bytes.subarray(2));
  }

  const decodedCandidates = CSV_DECODER_CANDIDATES.map((encoding) => {
    try {
      const decodedValue = new TextDecoder(encoding).decode(bytes);
      return {
        encoding,
        value: decodedValue,
        score: scoreDecodedCsv(decodedValue)
      };
    } catch {
      return null;
    }
  }).filter((candidate): candidate is { encoding: string; value: string; score: number } =>
    Boolean(candidate)
  );

  return (
    decodedCandidates.sort((left, right) => right.score - left.score)[0]?.value ??
    new TextDecoder("utf-8").decode(bytes)
  );
}

function parseWorkbook(fileName: string, content: ArrayBuffer | string): WorkbookState {
  const workbook = XLSX.read(content, {
    cellDates: true,
    type: typeof content === "string" ? "string" : "array"
  });

  const sheets = workbook.SheetNames.map((sheetName) =>
    parseWorksheet(sheetName, workbook.Sheets[sheetName])
  );

  if (sheets.length === 0) {
    throw new Error("此活頁簿中找不到工作表。");
  }

  return {
    fileName,
    sheets,
    activeSheetName: sheets[0].name
  };
}

function isSupportedSpreadsheetFile(fileName: string) {
  const normalizedFileName = fileName.toLowerCase();
  return SUPPORTED_FILE_EXTENSIONS.some((extension) => normalizedFileName.endsWith(extension));
}

function stripSupportedSpreadsheetExtension(fileName: string) {
  return fileName.replace(/\.(xlsx|xls|csv)$/i, "");
}

function quoteSqlIdentifier(identifier: string) {
  return `"${identifier.replace(/"/g, '""')}"`;
}

function toSqlValue(value: CellValue): SqlValue {
  if (value === null || value === "") {
    return null;
  }

  if (value instanceof Date) {
    return formatDateLabel(value);
  }

  if (typeof value === "boolean") {
    return value ? 1 : 0;
  }

  return value;
}

function fromSqlValue(value: SqlValue): CellValue {
  if (value instanceof Uint8Array) {
    return new TextDecoder("utf-8").decode(value);
  }

  return value;
}

function validateSqlQuery(sqlText: string) {
  const normalizedSql = sqlText.trim();
  if (!normalizedSql) {
    throw new Error("請輸入 SQL 查詢。");
  }

  if (!/^(select|with)\b/i.test(normalizedSql)) {
    throw new Error("目前僅支援 SELECT 或 WITH 查詢。");
  }

  return normalizedSql;
}

function createSqlTableFromSheet(
  db: Database,
  tableName: string,
  sheet: ParsedSheet
) {
  if (sheet.headers.length === 0) {
    db.run(`CREATE TABLE ${quoteSqlIdentifier(tableName)} ("_empty" TEXT)`);
    return;
  }

  const quotedTableName = quoteSqlIdentifier(tableName);
  const quotedHeaders = sheet.headers.map(quoteSqlIdentifier);
  db.run(`CREATE TABLE ${quotedTableName} (${quotedHeaders.join(", ")})`);

  if (sheet.rows.length === 0) {
    return;
  }

  const placeholders = sheet.headers.map(() => "?").join(", ");
  const insertStatement = db.prepare(
    `INSERT INTO ${quotedTableName} (${quotedHeaders.join(", ")}) VALUES (${placeholders})`
  );

  try {
    sheet.rows.forEach((row) => {
      insertStatement.run(sheet.headers.map((header) => toSqlValue(row[header] ?? null)));
    });
  } finally {
    insertStatement.free();
  }
}

function executeSqlQuery(
  sqlApi: SqlJsStatic,
  workbook: WorkbookState,
  activeSheet: ParsedSheet,
  sqlText: string
): ParsedSheet {
  const normalizedSql = validateSqlQuery(sqlText);
  const db = new sqlApi.Database();

  try {
    workbook.sheets.forEach((sheet) => {
      createSqlTableFromSheet(db, sheet.name, sheet);
    });

    if (!workbook.sheets.some((sheet) => sheet.name.toLowerCase() === SQL_TABLE_NAME)) {
      createSqlTableFromSheet(db, SQL_TABLE_NAME, activeSheet);
    }

    const results: QueryExecResult[] = db.exec(normalizedSql);
    const result = results[results.length - 1];
    if (!result) {
      throw new Error("SQL 查詢必須回傳結果集。");
    }

    const resultHeaders = normalizeHeaders(result.columns, result.columns.length);

    return {
      name: "SQL 查詢結果",
      headers: resultHeaders,
      rows: result.values.map((values) =>
        resultHeaders.reduce<Record<string, CellValue>>((row, column, index) => {
          row[column] = fromSqlValue(values[index] ?? null);
          return row;
        }, {})
      )
    };
  } finally {
    db.close();
  }
}

function App() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const chartRef = useRef<ReactECharts>(null);
  const [sqlApi, setSqlApi] = useState<SqlJsStatic | null>(null);
  const [sqlText, setSqlText] = useState(DEFAULT_SQL_QUERY);
  const [sqlErrorMessage, setSqlErrorMessage] = useState("");
  const [querySheet, setQuerySheet] = useState<ParsedSheet | null>(null);
  const [workbookState, setWorkbookState] = useState<WorkbookState | null>(null);
  const [errorMessage, setErrorMessage] = useState("");
  const [exportMessage, setExportMessage] = useState("");
  const [exportErrorMessage, setExportErrorMessage] = useState("");
  const [copyMessage, setCopyMessage] = useState("");
  const [copyErrorMessage, setCopyErrorMessage] = useState("");
  const [isExporting, setIsExporting] = useState(false);
  const [queryExportMessage, setQueryExportMessage] = useState("");
  const [queryExportErrorMessage, setQueryExportErrorMessage] = useState("");
  const [isExportingQuery, setIsExportingQuery] = useState(false);
  const [chartSettings, setChartSettings] = useState<ChartSettings>(loadChartSettings);
  const [recentFileName, setRecentFileName] = useState(
    () => window.localStorage.getItem(RECENT_FILE_STORAGE_KEY) ?? ""
  );

  const activeSheet = useMemo(() => {
    return (
      workbookState?.sheets.find((sheet) => sheet.name === workbookState.activeSheetName) ?? null
    );
  }, [workbookState]);

  const displaySheet = querySheet ?? activeSheet;
  const previewRows = displaySheet?.rows.slice(0, PREVIEW_ROW_LIMIT) ?? [];
  const canCopyDisplaySheet = Boolean(displaySheet && displaySheet.headers.length > 0);
  const fieldMetadata = useMemo(() => {
    return displaySheet ? buildFieldMetadata(displaySheet) : [];
  }, [displaySheet]);
  const chartValidation = useMemo(() => {
    return buildChartValidation(displaySheet, fieldMetadata, chartSettings);
  }, [displaySheet, fieldMetadata, chartSettings]);
  const chartOption = useMemo(() => {
    return buildChartOption(displaySheet, chartSettings, chartValidation);
  }, [displaySheet, chartSettings, chartValidation]);
  const hasRestoredMissingFields = useMemo(() => {
    if (!displaySheet) {
      return false;
    }

    return [chartSettings.xField, chartSettings.yField, chartSettings.categoryField]
      .filter(Boolean)
      .some((fieldName) => !fieldExists(fieldMetadata, fieldName));
  }, [displaySheet, chartSettings, fieldMetadata]);

  useEffect(() => {
    let isMounted = true;

    initSqlJs({
      locateFile: () => sqlWasmUrl
    })
      .then((loadedSqlApi) => {
        if (isMounted) {
          setSqlApi(loadedSqlApi);
        }
      })
      .catch((error) => {
        writeAppLog("error", `Unable to load SQL engine: ${describeError(error)}`);
        if (isMounted) {
          setSqlErrorMessage("無法載入 SQL 查詢引擎。");
        }
      });

    return () => {
      isMounted = false;
    };
  }, []);

  useEffect(() => {
    window.localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(chartSettings));
  }, [chartSettings]);

  useEffect(() => {
    setSqlText(DEFAULT_SQL_QUERY);
    setSqlErrorMessage("");
    setQuerySheet(activeSheet);
  }, [activeSheet]);

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

  function clearCopyMessages() {
    setCopyMessage("");
    setCopyErrorMessage("");
  }

  function clearQueryExportMessages() {
    setQueryExportMessage("");
    setQueryExportErrorMessage("");
  }

  function handleSqlTextChange(nextSqlText: string) {
    setSqlText(nextSqlText);
    setSqlErrorMessage("");
    setQuerySheet(activeSheet);
    clearCopyMessages();
    clearQueryExportMessages();
  }

  function handleSqlExecute() {
    if (!workbookState || !activeSheet || !sqlApi) {
      return;
    }

    try {
      const nextQuerySheet = executeSqlQuery(sqlApi, workbookState, activeSheet, sqlText);
      setQuerySheet(nextQuerySheet);
      setSqlErrorMessage("");
      clearCopyMessages();
      clearQueryExportMessages();
      resetChartFields();
    } catch (error) {
      writeAppLog("warn", `SQL query failed: ${describeError(error)}`);
      setSqlErrorMessage(error instanceof Error ? error.message : "SQL 查詢失敗。");
    }
  }

  function handleSqlReset() {
    setSqlText(DEFAULT_SQL_QUERY);
    setSqlErrorMessage("");
    setQuerySheet(activeSheet);
    clearCopyMessages();
    clearQueryExportMessages();
    resetChartFields();
  }

  async function handleCopyQueryResult() {
    if (!displaySheet || displaySheet.headers.length === 0) {
      return;
    }

    setCopyMessage("");
    setCopyErrorMessage("");

    try {
      if (!navigator.clipboard?.writeText) {
        throw new Error("Clipboard API is unavailable.");
      }

      await navigator.clipboard.writeText(buildSheetTsv(displaySheet));
      setCopyMessage(`已複製 ${displaySheet.rows.length} 筆資料列`);
    } catch (error) {
      writeAppLog("warn", `Unable to copy query result: ${describeError(error)}`);
      setCopyErrorMessage("無法複製資料。請確認剪貼簿權限。");
    }
  }

  async function handleExportQueryXlsx() {
    if (!displaySheet || displaySheet.headers.length === 0) {
      return;
    }

    setIsExportingQuery(true);
    setQueryExportMessage("");
    setQueryExportErrorMessage("");

    try {
      const workbookBytes = buildSheetWorkbookBytes(displaySheet);
      const fileBaseName = workbookState
        ? stripSupportedSpreadsheetExtension(workbookState.fileName)
        : "查詢結果";
      const savedPath = await invoke<SaveQueryXlsxResult>("save_query_xlsx", {
        fileName: `${fileBaseName}-${displaySheet.name}.xlsx`,
        workbookBytes
      });

      setQueryExportMessage(`已匯出 ${displaySheet.rows.length} 筆資料列至 ${savedPath}`);
    } catch (error) {
      writeAppLog("error", `Unable to export query XLSX: ${describeError(error)}`);
      setQueryExportErrorMessage("無法匯出 Excel。");
    } finally {
      setIsExportingQuery(false);
    }
  }

  async function handleExportPng() {
    if (!chartValidation.valid) {
      return;
    }

    const chartInstance = chartRef.current?.getEchartsInstance();
    if (!chartInstance) {
      setExportErrorMessage("圖表尚未準備完成。");
      return;
    }

    setIsExporting(true);
    setExportMessage("");
    setExportErrorMessage("");

    try {
      const imageUrl = chartInstance.getDataURL({
        type: "png",
        pixelRatio: 2,
        backgroundColor: "#ffffff"
      });
      const imageResponse = await fetch(imageUrl);
      const imageBuffer = await imageResponse.arrayBuffer();
      const imageBytes = Array.from(new Uint8Array(imageBuffer));
      const fileBaseName = workbookState
        ? stripSupportedSpreadsheetExtension(workbookState.fileName)
        : "圖表";
      const savedPath = await invoke<SaveChartPngResult>("save_chart_png", {
        fileName: `${fileBaseName}-${chartTypeLabels[chartSettings.chartType]}.png`,
        imageBytes
      });

      setExportMessage(`已儲存至 ${savedPath}`);
    } catch (error) {
      writeAppLog("error", `Unable to export PNG: ${describeError(error)}`);
      setExportErrorMessage("無法匯出 PNG。");
    } finally {
      setIsExporting(false);
    }
  }

  async function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];

    if (!file) {
      return;
    }

    if (!isSupportedSpreadsheetFile(file.name)) {
      setErrorMessage(`僅支援 ${SUPPORTED_FILE_LABEL} 檔案。`);
      setWorkbookState(null);
      resetChartFields();
      event.target.value = "";
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const content = isCsvFile(file.name) ? decodeCsvBuffer(buffer) : buffer;
      const parsedWorkbook = parseWorkbook(file.name, content);
      setWorkbookState(parsedWorkbook);
      setErrorMessage("");
      setSqlText(DEFAULT_SQL_QUERY);
      setSqlErrorMessage("");
      setQuerySheet(parsedWorkbook.sheets[0] ?? null);
      clearCopyMessages();
      clearQueryExportMessages();
      setExportMessage("");
      setExportErrorMessage("");
      setRecentFileName(file.name);
      window.localStorage.setItem(RECENT_FILE_STORAGE_KEY, file.name);
    } catch (error) {
      setWorkbookState(null);
      resetChartFields();
      writeAppLog("error", `Unable to parse workbook ${file.name}: ${describeError(error)}`);
      setErrorMessage("無法解析此活頁簿。");
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
    clearCopyMessages();
    clearQueryExportMessages();
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div>
          <h1>Excel 視覺化工具</h1>
          <p>{workbookState?.fileName ?? "本機活頁簿視覺化工作區。"}</p>
        </div>
        <input
          ref={fileInputRef}
          className="file-input"
          type="file"
          accept={SUPPORTED_FILE_EXTENSIONS.join(",")}
          onChange={handleFileChange}
        />
        <button
          type="button"
          className="primary-button"
          onClick={() => fileInputRef.current?.click()}
        >
          開啟試算表
        </button>
        {errorMessage ? <div className="error-message">{errorMessage}</div> : null}
        {!workbookState && recentFileName ? (
          <div className="info-message">最近檔案：{recentFileName}</div>
        ) : null}
        <section className="panel">
          <h2>工作表</h2>
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
                  <small>{sheet.rows.length} 筆資料列</small>
                </button>
              ))}
            </div>
          ) : (
            <div className="empty-state">尚未選擇活頁簿</div>
          )}
        </section>
      </aside>

      <section className="workspace">
        <div className="sql-area">
          <div className="section-header">
            <h2>SQL 查詢</h2>
            <span>
              {displaySheet
                ? `查詢結果 ${displaySheet.rows.length} 筆`
                : "等待資料"}
            </span>
          </div>
          <div className="sql-editor">
            <textarea
              value={sqlText}
              disabled={!activeSheet}
              spellCheck={false}
              onChange={(event) => handleSqlTextChange(event.target.value)}
            />
            <div className="sql-actions">
              <button
                type="button"
                className="primary-button"
                disabled={!workbookState || !activeSheet || !sqlApi}
                onClick={handleSqlExecute}
              >
                執行 SQL
              </button>
              <button
                type="button"
                className="secondary-button"
                disabled={!activeSheet}
                onClick={handleSqlReset}
              >
                重設
              </button>
            </div>
            {workbookState ? (
              <div className="sql-help">
                可用工作表名稱做 JOIN；中文、空白或特殊字元請加雙引號，例如{" "}
                {quoteSqlIdentifier(workbookState.sheets[0]?.name ?? "工作表1")}。data 代表目前工作表。
              </div>
            ) : null}
            {sqlErrorMessage ? <div className="error-message">{sqlErrorMessage}</div> : null}
          </div>
        </div>

        <div className="preview-area">
          <div className="section-header">
            <h2>資料預覽</h2>
            <div className="section-header-actions">
              <button
                type="button"
                className="secondary-button compact-button"
                disabled={!canCopyDisplaySheet}
                onClick={handleCopyQueryResult}
              >
                複製資料
              </button>
              <button
                type="button"
                className="secondary-button compact-button"
                disabled={!canCopyDisplaySheet || isExportingQuery}
                onClick={handleExportQueryXlsx}
              >
                {isExportingQuery ? "匯出中..." : "匯出 Excel"}
              </button>
              <span>
                {displaySheet
                  ? `顯示 ${Math.min(displaySheet.rows.length, PREVIEW_ROW_LIMIT)} / ${
                      displaySheet.rows.length
                    } 筆資料列`
                  : `前 ${PREVIEW_ROW_LIMIT} 筆資料列`}
              </span>
            </div>
          </div>
          {copyMessage ? <div className="info-message preview-message">{copyMessage}</div> : null}
          {copyErrorMessage ? (
            <div className="error-message preview-message">{copyErrorMessage}</div>
          ) : null}
          {queryExportMessage ? (
            <div className="info-message preview-message">{queryExportMessage}</div>
          ) : null}
          {queryExportErrorMessage ? (
            <div className="error-message preview-message">{queryExportErrorMessage}</div>
          ) : null}
          {displaySheet && displaySheet.headers.length > 0 ? (
            <div className="table-scroll">
              <table className="data-table">
                <thead>
                  <tr>
                    {displaySheet.headers.map((header) => (
                      <th key={header}>{header}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {previewRows.map((row, rowIndex) => (
                    <tr key={`${displaySheet.name}-${rowIndex}`}>
                      {displaySheet.headers.map((header) => (
                        <td key={header}>{formatCellValue(row[header] ?? null)}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
              {displaySheet.rows.length === 0 ? (
                <div className="table-empty-note">目前查詢結果沒有資料列。</div>
              ) : null}
            </div>
          ) : (
            <div className="table-placeholder">
              <span>
                {displaySheet
                  ? "目前查詢結果沒有可預覽的欄位。"
                  : "請選擇 Excel 檔案以預覽工作表資料列。"}
              </span>
            </div>
          )}
        </div>

        <div className="chart-area">
          <div className="section-header">
            <h2>圖表預覽</h2>
            <span>
              {chartValidation.valid
                ? chartTypeLabels[chartSettings.chartType]
                : "等待有效設定"}
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
        <h2>圖表設定</h2>
        <label>
          圖表類型
          <select
            value={chartSettings.chartType}
            onChange={(event) =>
              updateChartSettings({
                chartType: event.target.value as ChartType
              })
            }
          >
            <option value="bar">長條圖</option>
            <option value="line">折線圖</option>
            <option value="pie">圓餅圖</option>
            <option value="scatter">散佈圖</option>
            <option value="stackedBar">堆疊長條圖</option>
            <option value="timeSeries">時間序列圖</option>
            <option value="area">面積圖</option>
            <option value="horizontalBar">水平長條圖</option>
            <option value="radar">雷達圖</option>
            <option value="treemap">矩形式樹狀圖</option>
            <option value="funnel">漏斗圖</option>
            <option value="gauge">儀表圖</option>
          </select>
        </label>
        <label>
          X 軸
          <select
            disabled={fieldMetadata.length === 0}
            value={fieldExists(fieldMetadata, chartSettings.xField) ? chartSettings.xField : ""}
            onChange={(event) => updateChartSettings({ xField: event.target.value })}
          >
            <option value="">選擇欄位</option>
            {fieldMetadata.map((field) => (
              <option key={field.name} value={field.name}>
                {field.name}
              </option>
            ))}
          </select>
        </label>
        <label>
          Y 軸
          <select
            disabled={fieldMetadata.length === 0}
            value={fieldExists(fieldMetadata, chartSettings.yField) ? chartSettings.yField : ""}
            onChange={(event) => updateChartSettings({ yField: event.target.value })}
          >
            <option value="">選擇欄位</option>
            {fieldMetadata.map((field) => (
              <option key={field.name} value={field.name}>
                {field.name} {field.type === "number" ? "" : `(${fieldTypeLabels[field.type]})`}
              </option>
            ))}
          </select>
        </label>
        <label>
          分類
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
              {chartSettings.chartType === "stackedBar" ? "選擇欄位" : "無"}
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
            已儲存的設定包含目前工作表中不存在的欄位。
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
          disabled={!chartValidation.valid || isExporting}
          onClick={handleExportPng}
        >
          {isExporting ? "匯出中..." : "匯出 PNG"}
        </button>
        {exportMessage ? <div className="info-message">{exportMessage}</div> : null}
        {exportErrorMessage ? <div className="error-message">{exportErrorMessage}</div> : null}
        <section className="metadata-panel">
          <h2>欄位</h2>
          {fieldMetadata.length > 0 ? (
            <div className="field-list">
              {fieldMetadata.map((field) => (
                <div className="field-row" key={field.name}>
                  <span>{field.name}</span>
                  <small>
                    {fieldTypeLabels[field.type]} · {Math.round(field.emptyRatio * 100)}% 空白
                  </small>
                </div>
              ))}
            </div>
          ) : (
            <div className="empty-state compact">沒有可用欄位</div>
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
