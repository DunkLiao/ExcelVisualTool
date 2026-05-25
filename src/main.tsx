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
  | "gauge"
  | "boxplot"
  | "heatmap"
  | "candlestick"
  | "sankey"
  | "sunburst"
  | "graph"
  | "themeRiver"
  | "calendarHeatmap";

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
  openField: string;
  highField: string;
  lowField: string;
  closeField: string;
};

type ChartValidation = {
  valid: boolean;
  message: string;
  warnings: string[];
};

type ChartHelpContent = {
  purpose: string;
  useCase: string;
  inputs: string;
  notes: string;
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
  categoryField: "",
  openField: "",
  highField: "",
  lowField: "",
  closeField: ""
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
  gauge: "儀表圖",
  boxplot: "箱型圖",
  heatmap: "熱力圖",
  candlestick: "K 線圖",
  sankey: "桑基圖",
  sunburst: "旭日圖",
  graph: "關係圖",
  themeRiver: "主題河流圖",
  calendarHeatmap: "日曆熱力圖"
};

const chartTypeOrder = Object.keys(chartTypeLabels) as ChartType[];

const chartHelpContents: Record<ChartType, ChartHelpContent> = {
  bar: {
    purpose: "比較不同分類的數值大小。",
    useCase: "適合比較產品、部門、地區或月份的銷售額與數量。",
    inputs: "X 軸選分類或日期欄位，數值選數字欄位，分類可選分組系列。",
    notes: "分類太多時會讓標籤擁擠，建議先用 SQL 篩選或彙總。"
  },
  line: {
    purpose: "觀察數值沿分類或時間順序的變化趨勢。",
    useCase: "適合銷售趨勢、成本變化、指標追蹤與週期性比較。",
    inputs: "X 軸選日期或有順序的分類欄位，數值選數字欄位，分類可選多條線。",
    notes: "X 軸若是文字分類，請先確認排序符合分析需求。"
  },
  pie: {
    purpose: "呈現各分類在總量中的占比。",
    useCase: "適合產品占比、地區占比、支出結構等組成分析。",
    inputs: "X 軸選分類名稱，數值選數字欄位。",
    notes: "分類過多或差異太小時不易判讀，建議保留主要分類。"
  },
  scatter: {
    purpose: "觀察兩個數值欄位之間的關係與離群值。",
    useCase: "適合價格與銷量、成本與利潤、數量與金額等相關性分析。",
    inputs: "X 軸與 Y 軸都要選數字欄位，分類可選點的分組。",
    notes: "資料點過多時可能重疊，可先用 SQL 篩選範圍。"
  },
  stackedBar: {
    purpose: "比較總量，同時看出每個分組的組成。",
    useCase: "適合各月份依產品或地區堆疊的銷售額比較。",
    inputs: "X 軸選主要分類，數值選數字欄位，分類必須選堆疊分組。",
    notes: "堆疊層數過多會降低可讀性，建議控制分類數。"
  },
  timeSeries: {
    purpose: "以時間軸呈現數值變化。",
    useCase: "適合日、月、年的交易量、營收、庫存或 KPI 追蹤。",
    inputs: "X 軸選日期欄位，數值選數字欄位，分類可選多條時間序列。",
    notes: "X 軸必須能被辨識為日期或時間。"
  },
  area: {
    purpose: "強調趨勢變化與累積量感。",
    useCase: "適合流量、營收、產量等隨時間或排序分類累積呈現的指標。",
    inputs: "X 軸選分類或日期欄位，數值選數字欄位，分類可選多個面積系列。",
    notes: "多系列重疊時可能遮蔽細節，必要時改用折線圖。"
  },
  horizontalBar: {
    purpose: "橫向比較分類數值，提升長文字標籤可讀性。",
    useCase: "適合產品名稱、客戶名稱或部門名稱較長的排名比較。",
    inputs: "X 軸選分類欄位，數值選數字欄位，分類可選分組系列。",
    notes: "若要呈現排名，建議先用 SQL 排序或限制筆數。"
  },
  radar: {
    purpose: "比較多個指標構成的輪廓差異。",
    useCase: "適合產品能力、部門績效、客戶評分等多面向比較。",
    inputs: "X 軸選指標名稱，數值選數字欄位，分類可選不同對象。",
    notes: "指標尺度差異太大時會影響判讀，建議先標準化資料。"
  },
  treemap: {
    purpose: "用面積大小呈現分類數值的占比與層級感。",
    useCase: "適合分類多、需要看出大項與小項差距的銷售或成本資料。",
    inputs: "X 軸選分類名稱，數值選數字欄位。",
    notes: "目前使用單層分類；需要多層結構時可用旭日圖。"
  },
  funnel: {
    purpose: "呈現流程階段之間的數量遞減或轉換。",
    useCase: "適合銷售漏斗、案件流程、報名到成交等階段分析。",
    inputs: "X 軸選階段名稱，數值選數字欄位。",
    notes: "資料最好代表有順序的流程階段，否則漏斗語意會不清楚。"
  },
  gauge: {
    purpose: "呈現單一核心指標目前值。",
    useCase: "適合達成率、總銷售額、使用率或單一 KPI 展示。",
    inputs: "數值選數字欄位，程式會彙總目前資料作為指標值。",
    notes: "儀表圖不適合比較多分類，若要比較請改用長條圖。"
  },
  boxplot: {
    purpose: "呈現分類中數值分布、四分位距與離散程度。",
    useCase: "適合比較不同產品、地區或部門的價格、工時、金額分布。",
    inputs: "X 軸選分類欄位，數值選數字欄位。",
    notes: "每個分類需要多筆數值才有分析意義，單筆資料只會形成很窄的分布。"
  },
  heatmap: {
    purpose: "用顏色深淺呈現兩個分類交叉後的數值強度。",
    useCase: "適合產品與地區、部門與月份、狀態與類型的交叉分析。",
    inputs: "X 軸選第一個分類，Y 分類選第二個分類，數值選數字欄位。",
    notes: "分類過多會造成格子太密，建議先篩選主要項目。"
  },
  candlestick: {
    purpose: "呈現開盤、最高、最低、收盤的價格區間。",
    useCase: "適合股票、匯率、商品價格或其他 OHLC 形式資料。",
    inputs: "日期或分類作為 X 軸，並選擇開盤、最高、最低、收盤四個數字欄位。",
    notes: "四個價格欄位都必須是數值，資料排序會依日期或 X 軸文字排序。"
  },
  sankey: {
    purpose: "呈現來源到目標之間的流量或轉移關係。",
    useCase: "適合資金流、客戶旅程、流程轉換、來源去向分析。",
    inputs: "來源選起點欄位，目標選終點欄位，權重選數字欄位。",
    notes: "同一來源與目標會自動加總；節點太多時圖面會較複雜。"
  },
  sunburst: {
    purpose: "用同心圓呈現父子層級與數值占比。",
    useCase: "適合部門到產品、地區到客戶、分類到子分類的階層分析。",
    inputs: "父層選第一層分類，子層選第二層分類，數值選數字欄位。",
    notes: "目前支援二層階層；更深層資料需先整理成適合的欄位。"
  },
  graph: {
    purpose: "呈現節點之間的連結與關係強弱。",
    useCase: "適合客戶關係、交易對手、供應鏈、關聯網路分析。",
    inputs: "來源選起點欄位，目標選終點欄位，權重選數字欄位。",
    notes: "節點數過多會增加閱讀難度，建議先聚焦重要關係。"
  },
  themeRiver: {
    purpose: "呈現多分類數值隨時間流動的相對變化。",
    useCase: "適合多產品銷售趨勢、分類流量變化、議題熱度追蹤。",
    inputs: "日期選時間欄位，分類選系列欄位，數值選數字欄位。",
    notes: "X 軸必須是日期；分類太多時河流會變得擁擠。"
  },
  calendarHeatmap: {
    purpose: "用日曆格呈現每日數值強度。",
    useCase: "適合每日銷售、出勤、交易量、事件次數或工作量分析。",
    inputs: "日期選日期欄位，數值選數字欄位。",
    notes: "資料會依日期彙總；跨年資料會依資料範圍顯示。"
  }
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
  return ["pie", "treemap", "funnel", "gauge", "sankey", "sunburst", "graph"].includes(
    chartType
  );
}

function shouldUseCategorySeries(chartType: ChartType) {
  return [
    "bar",
    "line",
    "stackedBar",
    "timeSeries",
    "area",
    "horizontalBar",
    "radar"
  ].includes(chartType);
}

function shouldRequireCategoryField(chartType: ChartType) {
  return ["stackedBar", "heatmap", "sankey", "sunburst", "graph", "themeRiver"].includes(
    chartType
  );
}

function shouldShowCategoryField(chartType: ChartType) {
  return [
    "bar",
    "line",
    "scatter",
    "stackedBar",
    "timeSeries",
    "area",
    "horizontalBar",
    "radar",
    "heatmap",
    "sankey",
    "sunburst",
    "graph",
    "themeRiver"
  ].includes(chartType);
}

function shouldShowYField(chartType: ChartType) {
  return chartType !== "candlestick";
}

function shouldShowCandlestickFields(chartType: ChartType) {
  return chartType === "candlestick";
}

function getXFieldLabel(chartType: ChartType) {
  if (chartType === "sankey" || chartType === "graph") {
    return "來源";
  }

  if (chartType === "sunburst") {
    return "子層";
  }

  if (chartType === "themeRiver" || chartType === "calendarHeatmap") {
    return "日期";
  }

  return chartType === "scatter" ? "X 軸" : "X 軸";
}

function getYFieldLabel(chartType: ChartType) {
  if (chartType === "scatter") {
    return "Y 軸";
  }

  if (chartType === "sankey" || chartType === "graph") {
    return "權重";
  }

  return "數值";
}

function getCategoryFieldLabel(chartType: ChartType) {
  if (chartType === "sankey" || chartType === "graph") {
    return "目標";
  }

  if (chartType === "sunburst") {
    return "父層";
  }

  if (chartType === "heatmap") {
    return "Y 分類";
  }

  return "分類";
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
      categoryField: parsedSettings.categoryField ?? "",
      openField: parsedSettings.openField ?? "",
      highField: parsedSettings.highField ?? "",
      lowField: parsedSettings.lowField ?? "",
      closeField: parsedSettings.closeField ?? ""
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
  const yLabel = getYFieldLabel(settings.chartType);
  const xError =
    settings.chartType === "candlestick"
      ? validateDimensionField(metadata, settings.xField, getXFieldLabel(settings.chartType))
      : settings.chartType === "scatter"
      ? validateNumericField(metadata, settings.xField, "X 軸")
      : ["timeSeries", "themeRiver", "calendarHeatmap"].includes(settings.chartType)
        ? validateDateField(metadata, settings.xField, getXFieldLabel(settings.chartType))
        : validateDimensionField(metadata, settings.xField, getXFieldLabel(settings.chartType));
  const yError = shouldShowYField(settings.chartType)
    ? validateNumericField(metadata, settings.yField, yLabel)
    : null;

  if (xError || yError) {
    return {
      valid: false,
      message: xError ?? yError ?? "請選擇有效的圖表欄位。",
      warnings
    };
  }

  if (settings.chartType === "candlestick") {
    const candlestickFieldErrors = [
      validateNumericField(metadata, settings.openField, "開盤"),
      validateNumericField(metadata, settings.highField, "最高"),
      validateNumericField(metadata, settings.lowField, "最低"),
      validateNumericField(metadata, settings.closeField, "收盤")
    ];
    const candlestickError = candlestickFieldErrors.find(Boolean);
    if (candlestickError) {
      return {
        valid: false,
        message: candlestickError,
        warnings
      };
    }
  }

  if (shouldRequireCategoryField(settings.chartType)) {
    const categoryError = validateDimensionField(
      metadata,
      settings.categoryField,
      getCategoryFieldLabel(settings.chartType)
    );
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
    shouldShowYField(settings.chartType) ? settings.yField : "",
    shouldUseCategorySeries(settings.chartType) ||
    settings.chartType === "scatter" ||
    shouldRequireCategoryField(settings.chartType)
      ? settings.categoryField
      : "",
    ...(
      settings.chartType === "candlestick"
        ? [settings.openField, settings.highField, settings.lowField, settings.closeField]
        : []
    )
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

function quantile(sortedValues: number[], percentile: number) {
  if (sortedValues.length === 0) {
    return 0;
  }

  const position = (sortedValues.length - 1) * percentile;
  const baseIndex = Math.floor(position);
  const rest = position - baseIndex;
  const nextValue = sortedValues[baseIndex + 1];

  return nextValue === undefined
    ? sortedValues[baseIndex]
    : sortedValues[baseIndex] + rest * (nextValue - sortedValues[baseIndex]);
}

function buildBoxplotData(sheet: ParsedSheet, settings: ChartSettings) {
  const groups = new Map<string, number[]>();

  sheet.rows.forEach((row) => {
    const rawXValue = row[settings.xField] ?? null;
    const yValue = coerceNumber(row[settings.yField] ?? null);
    if (isEmptyCell(rawXValue) || yValue === null) {
      return;
    }

    const xValue = formatDimensionValue(rawXValue);
    groups.set(xValue, [...(groups.get(xValue) ?? []), yValue]);
  });

  const xValues = Array.from(groups.keys());
  const data = xValues.map((xValue) => {
    const values = [...(groups.get(xValue) ?? [])].sort((left, right) => left - right);
    return [
      values[0] ?? 0,
      quantile(values, 0.25),
      quantile(values, 0.5),
      quantile(values, 0.75),
      values[values.length - 1] ?? 0
    ];
  });

  return { xValues, data };
}

function aggregatePairRows(sheet: ParsedSheet, settings: ChartSettings) {
  const xValues: string[] = [];
  const categoryValues: string[] = [];
  const values = new Map<string, number>();

  sheet.rows.forEach((row) => {
    const rawXValue = row[settings.xField] ?? null;
    const rawCategoryValue = row[settings.categoryField] ?? null;
    const yValue = coerceNumber(row[settings.yField] ?? null);
    if (isEmptyCell(rawXValue) || isEmptyCell(rawCategoryValue) || yValue === null) {
      return;
    }

    const xValue = formatDimensionValue(rawXValue);
    const categoryValue = formatDimensionValue(rawCategoryValue);
    if (!xValues.includes(xValue)) {
      xValues.push(xValue);
    }
    if (!categoryValues.includes(categoryValue)) {
      categoryValues.push(categoryValue);
    }

    const key = `${xValue}\u0000${categoryValue}`;
    values.set(key, (values.get(key) ?? 0) + yValue);
  });

  return { xValues, categoryValues, values };
}

function buildNetworkData(sheet: ParsedSheet, settings: ChartSettings) {
  const nodes = new Set<string>();
  const links = new Map<string, { source: string; target: string; value: number }>();

  sheet.rows.forEach((row) => {
    const rawSource = row[settings.xField] ?? null;
    const rawTarget = row[settings.categoryField] ?? null;
    const value = coerceNumber(row[settings.yField] ?? null);
    if (isEmptyCell(rawSource) || isEmptyCell(rawTarget) || value === null) {
      return;
    }

    const source = formatDimensionValue(rawSource);
    const target = formatDimensionValue(rawTarget);
    nodes.add(source);
    nodes.add(target);

    const key = `${source}\u0000${target}`;
    const currentLink = links.get(key);
    links.set(key, {
      source,
      target,
      value: (currentLink?.value ?? 0) + value
    });
  });

  return {
    nodes: Array.from(nodes).map((name) => ({ name })),
    links: Array.from(links.values())
  };
}

function buildSunburstData(sheet: ParsedSheet, settings: ChartSettings) {
  const parentValues = new Map<string, Map<string, number>>();

  sheet.rows.forEach((row) => {
    const rawParent = row[settings.categoryField] ?? null;
    const rawChild = row[settings.xField] ?? null;
    const value = coerceNumber(row[settings.yField] ?? null);
    if (isEmptyCell(rawParent) || isEmptyCell(rawChild) || value === null) {
      return;
    }

    const parent = formatDimensionValue(rawParent);
    const child = formatDimensionValue(rawChild);
    if (!parentValues.has(parent)) {
      parentValues.set(parent, new Map<string, number>());
    }

    const children = parentValues.get(parent);
    children?.set(child, (children.get(child) ?? 0) + value);
  });

  return Array.from(parentValues.entries()).map(([parentName, children]) => ({
    name: parentName,
    children: Array.from(children.entries()).map(([childName, value]) => ({
      name: childName,
      value
    }))
  }));
}

function buildCandlestickData(sheet: ParsedSheet, settings: ChartSettings) {
  const rows = sheet.rows
    .map((row) => {
      const rawXValue = row[settings.xField] ?? null;
      const open = coerceNumber(row[settings.openField] ?? null);
      const high = coerceNumber(row[settings.highField] ?? null);
      const low = coerceNumber(row[settings.lowField] ?? null);
      const close = coerceNumber(row[settings.closeField] ?? null);
      const dateTime = coerceDateTime(rawXValue);
      if (isEmptyCell(rawXValue) || open === null || high === null || low === null || close === null) {
        return null;
      }

      return {
        label: dateTime === null ? formatDimensionValue(rawXValue) : formatDateLabel(dateTime),
        sortValue: dateTime ?? formatDimensionValue(rawXValue),
        value: [open, close, low, high]
      };
    })
    .filter((row): row is { label: string; sortValue: string | number; value: number[] } =>
      Boolean(row)
    )
    .sort((left, right) => {
      if (typeof left.sortValue === "number" && typeof right.sortValue === "number") {
        return left.sortValue - right.sortValue;
      }

      return String(left.sortValue).localeCompare(String(right.sortValue));
    });

  return {
    xValues: rows.map((row) => row.label),
    data: rows.map((row) => row.value)
  };
}

function buildThemeRiverData(sheet: ParsedSheet, settings: ChartSettings) {
  return sheet.rows
    .map((row) => {
      const dateTime = coerceDateTime(row[settings.xField] ?? null);
      const value = coerceNumber(row[settings.yField] ?? null);
      const rawCategory = row[settings.categoryField] ?? null;
      if (dateTime === null || value === null || isEmptyCell(rawCategory)) {
        return null;
      }

      return [formatDateLabel(dateTime), value, formatDimensionValue(rawCategory)];
    })
    .filter((row): row is [string, number, string] => Boolean(row));
}

function buildCalendarHeatmapData(sheet: ParsedSheet, settings: ChartSettings) {
  const values = new Map<string, number>();

  sheet.rows.forEach((row) => {
    const dateTime = coerceDateTime(row[settings.xField] ?? null);
    const value = coerceNumber(row[settings.yField] ?? null);
    if (dateTime === null || value === null) {
      return;
    }

    const dateLabel = formatDateLabel(dateTime);
    values.set(dateLabel, (values.get(dateLabel) ?? 0) + value);
  });

  const data = Array.from(values.entries()).sort(([leftDate], [rightDate]) =>
    leftDate.localeCompare(rightDate)
  );
  const range = data.length > 0 ? [data[0][0], data[data.length - 1][0]] : undefined;

  return { data, range };
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

  if (settings.chartType === "boxplot") {
    const { xValues, data } = buildBoxplotData(sheet, settings);
    return {
      ...baseOption,
      tooltip: {
        trigger: "item"
      },
      xAxis: {
        type: "category",
        data: xValues
      },
      yAxis: {
        type: "value",
        name: settings.yField
      },
      series: [
        {
          name: settings.yField,
          type: "boxplot",
          data
        }
      ]
    };
  }

  if (settings.chartType === "heatmap") {
    const { xValues, categoryValues, values } = aggregatePairRows(sheet, settings);
    const data = xValues.flatMap((xValue, xIndex) =>
      categoryValues.map((categoryValue, categoryIndex) => [
        xIndex,
        categoryIndex,
        values.get(`${xValue}\u0000${categoryValue}`) ?? 0
      ])
    );
    const maxValue = Math.max(0, ...data.map((item) => Number(item[2])));

    return {
      ...baseOption,
      tooltip: {
        trigger: "item"
      },
      grid: {
        left: 88,
        right: 28,
        top: 56,
        bottom: 72,
        containLabel: true
      },
      xAxis: {
        type: "category",
        data: xValues,
        splitArea: {
          show: true
        }
      },
      yAxis: {
        type: "category",
        data: categoryValues,
        splitArea: {
          show: true
        }
      },
      visualMap: {
        min: 0,
        max: maxValue <= 0 ? 100 : maxValue,
        calculable: true,
        orient: "horizontal",
        left: "center",
        bottom: 12
      },
      series: [
        {
          name: settings.yField,
          type: "heatmap",
          data,
          label: {
            show: true
          }
        }
      ]
    };
  }

  if (settings.chartType === "candlestick") {
    const { xValues, data } = buildCandlestickData(sheet, settings);
    return {
      ...baseOption,
      tooltip: {
        trigger: "axis"
      },
      legend: {
        show: false
      },
      xAxis: {
        type: "category",
        data: xValues,
        scale: true
      },
      yAxis: {
        type: "value",
        scale: true
      },
      series: [
        {
          name: chartTypeLabels.candlestick,
          type: "candlestick",
          data
        }
      ]
    };
  }

  if (settings.chartType === "sankey") {
    const { nodes, links } = buildNetworkData(sheet, settings);
    return {
      ...baseOption,
      tooltip: {
        trigger: "item"
      },
      series: [
        {
          name: chartTypeLabels.sankey,
          type: "sankey",
          emphasis: {
            focus: "adjacency"
          },
          data: nodes,
          links
        }
      ]
    };
  }

  if (settings.chartType === "sunburst") {
    return {
      ...baseOption,
      tooltip: {
        trigger: "item"
      },
      series: [
        {
          name: chartTypeLabels.sunburst,
          type: "sunburst",
          radius: [0, "88%"],
          sort: undefined,
          data: buildSunburstData(sheet, settings),
          label: {
            rotate: "radial"
          }
        }
      ]
    };
  }

  if (settings.chartType === "graph") {
    const { nodes, links } = buildNetworkData(sheet, settings);
    return {
      ...baseOption,
      tooltip: {
        trigger: "item"
      },
      series: [
        {
          name: chartTypeLabels.graph,
          type: "graph",
          layout: "force",
          roam: true,
          draggable: true,
          label: {
            show: true
          },
          force: {
            repulsion: 120,
            edgeLength: 80
          },
          data: nodes,
          links: links.map((link) => ({
            ...link,
            lineStyle: {
              width: Math.max(1, Math.min(8, Math.sqrt(link.value)))
            }
          }))
        }
      ]
    };
  }

  if (settings.chartType === "themeRiver") {
    return {
      ...baseOption,
      tooltip: {
        trigger: "axis"
      },
      singleAxis: {
        type: "time",
        top: 56,
        bottom: 56,
        axisLabel: {
          formatter: (value: number) => formatDateLabel(value, getTimeAxisGranularity([value]))
        }
      },
      series: [
        {
          name: chartTypeLabels.themeRiver,
          type: "themeRiver",
          emphasis: {
            itemStyle: {
              shadowBlur: 12,
              shadowColor: "rgba(0, 0, 0, 0.25)"
            }
          },
          data: buildThemeRiverData(sheet, settings)
        }
      ]
    };
  }

  if (settings.chartType === "calendarHeatmap") {
    const { data, range } = buildCalendarHeatmapData(sheet, settings);
    const maxValue = Math.max(0, ...data.map(([, value]) => value));
    return {
      ...baseOption,
      tooltip: {
        trigger: "item"
      },
      visualMap: {
        min: 0,
        max: maxValue <= 0 ? 100 : maxValue,
        calculable: true,
        orient: "horizontal",
        left: "center",
        bottom: 8
      },
      calendar: {
        top: 64,
        left: 48,
        right: 24,
        bottom: 76,
        cellSize: ["auto", 18],
        range,
        itemStyle: {
          borderWidth: 0.5
        },
        yearLabel: {
          show: true
        }
      },
      series: [
        {
          name: settings.yField,
          type: "heatmap",
          coordinateSystem: "calendar",
          data
        }
      ]
    };
  }

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
  const [isChartHelpOpen, setIsChartHelpOpen] = useState(false);
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

    return [
      chartSettings.xField,
      chartSettings.yField,
      chartSettings.categoryField,
      chartSettings.openField,
      chartSettings.highField,
      chartSettings.lowField,
      chartSettings.closeField
    ]
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
      categoryField: "",
      openField: "",
      highField: "",
      lowField: "",
      closeField: ""
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
    <>
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
        <div className="settings-header">
          <h2>圖表設定</h2>
          <button
            type="button"
            className="help-button"
            aria-label="開啟圖形說明"
            onClick={() => setIsChartHelpOpen(true)}
          >
            ?
          </button>
        </div>
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
            <option value="boxplot">箱型圖</option>
            <option value="heatmap">熱力圖</option>
            <option value="candlestick">K 線圖</option>
            <option value="sankey">桑基圖</option>
            <option value="sunburst">旭日圖</option>
            <option value="graph">關係圖</option>
            <option value="themeRiver">主題河流圖</option>
            <option value="calendarHeatmap">日曆熱力圖</option>
          </select>
        </label>
        <label>
          {getXFieldLabel(chartSettings.chartType)}
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
        {shouldShowYField(chartSettings.chartType) ? (
          <label>
            {getYFieldLabel(chartSettings.chartType)}
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
        ) : null}
        {shouldShowCategoryField(chartSettings.chartType) ? (
          <label>
            {getCategoryFieldLabel(chartSettings.chartType)}
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
                {shouldRequireCategoryField(chartSettings.chartType) ? "選擇欄位" : "無"}
              </option>
              {fieldMetadata.map((field) => (
                <option key={field.name} value={field.name}>
                  {field.name}
                </option>
              ))}
            </select>
          </label>
        ) : null}
        {shouldShowCandlestickFields(chartSettings.chartType) ? (
          <>
            <label>
              開盤
              <select
                disabled={fieldMetadata.length === 0}
                value={
                  fieldExists(fieldMetadata, chartSettings.openField)
                    ? chartSettings.openField
                    : ""
                }
                onChange={(event) => updateChartSettings({ openField: event.target.value })}
              >
                <option value="">選擇欄位</option>
                {fieldMetadata.map((field) => (
                  <option key={field.name} value={field.name}>
                    {field.name}{" "}
                    {field.type === "number" ? "" : `(${fieldTypeLabels[field.type]})`}
                  </option>
                ))}
              </select>
            </label>
            <label>
              最高
              <select
                disabled={fieldMetadata.length === 0}
                value={
                  fieldExists(fieldMetadata, chartSettings.highField)
                    ? chartSettings.highField
                    : ""
                }
                onChange={(event) => updateChartSettings({ highField: event.target.value })}
              >
                <option value="">選擇欄位</option>
                {fieldMetadata.map((field) => (
                  <option key={field.name} value={field.name}>
                    {field.name}{" "}
                    {field.type === "number" ? "" : `(${fieldTypeLabels[field.type]})`}
                  </option>
                ))}
              </select>
            </label>
            <label>
              最低
              <select
                disabled={fieldMetadata.length === 0}
                value={
                  fieldExists(fieldMetadata, chartSettings.lowField)
                    ? chartSettings.lowField
                    : ""
                }
                onChange={(event) => updateChartSettings({ lowField: event.target.value })}
              >
                <option value="">選擇欄位</option>
                {fieldMetadata.map((field) => (
                  <option key={field.name} value={field.name}>
                    {field.name}{" "}
                    {field.type === "number" ? "" : `(${fieldTypeLabels[field.type]})`}
                  </option>
                ))}
              </select>
            </label>
            <label>
              收盤
              <select
                disabled={fieldMetadata.length === 0}
                value={
                  fieldExists(fieldMetadata, chartSettings.closeField)
                    ? chartSettings.closeField
                    : ""
                }
                onChange={(event) => updateChartSettings({ closeField: event.target.value })}
              >
                <option value="">選擇欄位</option>
                {fieldMetadata.map((field) => (
                  <option key={field.name} value={field.name}>
                    {field.name}{" "}
                    {field.type === "number" ? "" : `(${fieldTypeLabels[field.type]})`}
                  </option>
                ))}
              </select>
            </label>
          </>
        ) : null}
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
    {isChartHelpOpen ? (
      <div
        className="modal-backdrop"
        role="presentation"
        onClick={() => setIsChartHelpOpen(false)}
      >
        <section
          className="chart-help-dialog"
          role="dialog"
          aria-modal="true"
          aria-labelledby="chart-help-title"
          onClick={(event) => event.stopPropagation()}
        >
          <div className="chart-help-header">
            <h2 id="chart-help-title">圖形說明</h2>
            <button
              type="button"
              className="dialog-close-button"
              aria-label="關閉圖形說明"
              onClick={() => setIsChartHelpOpen(false)}
            >
              ×
            </button>
          </div>
          <div className="chart-help-list">
            {chartTypeOrder.map((chartType) => {
              const helpContent = chartHelpContents[chartType];
              return (
                <article className="chart-help-card" key={chartType}>
                  <h3>{chartTypeLabels[chartType]}</h3>
                  <dl>
                    <div>
                      <dt>目的</dt>
                      <dd>{helpContent.purpose}</dd>
                    </div>
                    <div>
                      <dt>使用情境</dt>
                      <dd>{helpContent.useCase}</dd>
                    </div>
                    <div>
                      <dt>輸入參數</dt>
                      <dd>{helpContent.inputs}</dd>
                    </div>
                    <div>
                      <dt>注意事項</dt>
                      <dd>{helpContent.notes}</dd>
                    </div>
                  </dl>
                </article>
              );
            })}
          </div>
        </section>
      </div>
    ) : null}
  </>
  );
}

installFrontendLogging();

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
