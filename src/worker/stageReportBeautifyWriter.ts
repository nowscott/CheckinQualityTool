import type { FoundSheet } from "./excelReader";
import { buildWorkbook } from "./excelWriter";
import {
  buildAssistantHierarchy,
  buildResearchGroupHierarchy,
  isStageReportAppealed,
  type HierarchyMetricSpec,
  type HierarchyTable,
} from "./stageReportHierarchy";
import type { CellValue, DataRow, SheetDefinition } from "./types";
import { text } from "./utils";

const STYLE = {
  title: 16,
  header: 17,
  singleLineHeader: 30,
  body: 18,
  rate: 19,
  detail: 20,
  detailRate: 21,
  groupTotal: 22,
  groupTotalRate: 23,
  projectTotal: 24,
  projectTotalRate: 25,
  grandTotal: 26,
  grandTotalRate: 27,
  sent: 28,
  unsent: 29,
} as const;

const STAGE_SENT_ALIASES = [
  "是否发送阶段性报告",
  "是否发送阶段性报告（系统数据）",
  "阶段性报告是否发送",
  "是否已发送（申诉+系统）",
  "是否已发送(申诉+系统)",
];
const WINDOW_SENT_ALIASES = ["是否发送窗口期报告", "窗口期报告是否发送"];
const STAGE_RATE_ALIASES = ["阶段性报告发送率", "非窗口期暑期在读阶段性报告发送率"];
const WINDOW_RATE_ALIASES = ["窗口期报告发送率"];
const STAGE_TOTAL_ALIASES = [
  "阶段性报告需发送",
  "阶段性报告需发送数",
  "阶段性报告应发送",
  "阶段性报告应发送数",
  "非窗口期暑期在读阶段性报告需发送",
  "非窗口期暑期在读阶段性报告需发送数",
  "非窗口期暑期在读阶段性报告应发送",
  "非窗口期暑期在读阶段性报告应发送数",
];
const STAGE_SENT_COUNT_ALIASES = [
  "阶段性报告已发送",
  "阶段性报告已发送数",
  "非窗口期暑期在读阶段性报告已发送",
  "非窗口期暑期在读阶段性报告已发送数",
];
const WINDOW_TOTAL_ALIASES = ["窗口期报告需发送", "窗口期报告需发送数", "窗口期报告应发送", "窗口期报告应发送数"];
const WINDOW_SENT_COUNT_ALIASES = ["窗口期报告已发送", "窗口期报告已发送数"];

const ASSISTANT_COLUMNS = [
  "教研组",
  "师训主管/助理主管",
  "阶段性报告应发送数",
  "阶段性报告已发送数",
  "阶段性报告发送率",
  "阶段性报告未发送数",
  "阶段性报告申诉数",
  "窗口期报告应发送数",
  "窗口期报告已发送数",
  "窗口期报告发送率",
  "窗口期报告未发送数",
] as const;

const RESEARCH_GROUP_COLUMNS = [
  "教研组",
  ...ASSISTANT_COLUMNS.slice(2),
] as const;

const HIERARCHY_METRICS: readonly HierarchyMetricSpec[] = [
  {
    totalColumn: "阶段性报告应发送数",
    sentColumn: "阶段性报告已发送数",
    rateColumn: "阶段性报告发送率",
    unsentColumn: "阶段性报告未发送数",
    appealColumn: "阶段性报告申诉数",
  },
  {
    totalColumn: "窗口期报告应发送数",
    sentColumn: "窗口期报告已发送数",
    rateColumn: "窗口期报告发送率",
    unsentColumn: "窗口期报告未发送数",
  },
];

const PERIOD_HIERARCHY_METRICS: readonly HierarchyMetricSpec[] = [
  {
    totalColumn: "窗口期报告应发送数",
    sentColumn: "窗口期报告已发送数",
    rateColumn: "窗口期报告发送率",
    unsentColumn: "窗口期报告未发送数",
  },
  ...["0819", "0805"].map((period) => ({
    totalColumn: `阶段性报告应发送数${period}`,
    sentColumn: `阶段性报告已发送数${period}`,
    rateColumn: `阶段性报告发送率${period}`,
    unsentColumn: `阶段性报告未发送数${period}`,
    appealColumn: `阶段性报告申诉数${period}`,
  })),
];

const PERIOD_ASSISTANT_COLUMNS = [
  "教研组",
  "师训主管/助理主管",
  ...PERIOD_HIERARCHY_METRICS.flatMap((metric) => [
    metric.totalColumn,
    metric.sentColumn,
    metric.rateColumn,
    metric.unsentColumn,
    ...(metric.appealColumn ? [metric.appealColumn] : []),
  ]),
] as const;

const PERIOD_RESEARCH_GROUP_COLUMNS = ["教研组", ...PERIOD_ASSISTANT_COLUMNS.slice(2)] as const;

function generatedDate() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
}

function formatDataTime(date: Date) {
  const pad = (value: number) => String(value).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function parseDataTime(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    return { timestamp: value.valueOf(), label: formatDataTime(value) };
  }
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value) as {
      y: number;
      m: number;
      d: number;
      H?: number;
      M?: number;
      S?: number;
    } | null;
    if (parsed) {
      const date = new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, Math.floor(parsed.S || 0));
      return { timestamp: date.valueOf(), label: formatDataTime(date) };
    }
  }
  const raw = text(value);
  const match = raw.match(/^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?(?:[ T]+(\d{1,2})[:：](\d{2})(?::(\d{2}))?)?\s*$/u);
  if (!match) return null;
  const date = new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4] || 0),
    Number(match[5] || 0),
    Number(match[6] || 0),
  );
  if (
    Number.isNaN(date.valueOf()) ||
    date.getFullYear() !== Number(match[1]) ||
    date.getMonth() !== Number(match[2]) - 1 ||
    date.getDate() !== Number(match[3])
  ) return null;
  return { timestamp: date.valueOf(), label: formatDataTime(date) };
}

function rowsForSheet(workbook: SheetJsWorkbook, name: string): CellValue[][] {
  return XLSX.utils.sheet_to_json(workbook.Sheets[name], {
    header: 1,
    range: 0,
    blankrows: false,
    defval: "",
  });
}

function latestUpdateTime(workbook: SheetJsWorkbook) {
  let latestTimestamp = Number.NEGATIVE_INFINITY;
  let latestLabel = "";
  workbook.SheetNames.forEach((name) => {
    const rows = rowsForSheet(workbook, name);
    if (!rows.length) return;
    const updateColumns = rows[0]
      .map((value, index) => /更新时间/u.test(text(value)) ? index : -1)
      .filter((index) => index >= 0);
    rows.slice(1).forEach((row) => {
      updateColumns.forEach((columnIndex) => {
        const parsed = parseDataTime(row[columnIndex]);
        if (parsed && parsed.timestamp > latestTimestamp) {
          latestTimestamp = parsed.timestamp;
          latestLabel = parsed.label;
        }
      });
    });
  });
  return latestLabel || generatedDate();
}

function findSheet(
  workbook: SheetJsWorkbook,
  required: (headers: string[]) => boolean,
  namePattern?: RegExp,
  description = "目标",
  expected = "所需业务表头",
): FoundSheet {
  const candidates: FoundSheet[] = [];
  workbook.SheetNames.forEach((name) => {
    const rows = rowsForSheet(workbook, name);
    if (!rows.length) return;
    const headers = rows[0].map(text);
    if (required(headers)) candidates.push({ name, rows });
  });
  if (!candidates.length) {
    const sheetNames = workbook.SheetNames.length ? workbook.SheetNames.join("、") : "无";
    throw new Error(
      `公示版整理失败：找不到${description}。${expected}。当前工作表：${sheetNames}。请上传窗口期与非窗口期原始表单，不要上传已生成的公示版。`,
    );
  }
  return candidates.find((sheet) => namePattern?.test(sheet.name)) || candidates[0];
}

function hasAny(headers: string[], aliases: readonly string[]) {
  return aliases.some((alias) => headers.includes(alias));
}

function findDetailSheet(workbook: SheetJsWorkbook, sentAliases: readonly string[], namePattern: RegExp) {
  return findSheet(
    workbook,
    (headers) => ["教师姓名", "学生姓名", "学号"].every((header) => headers.includes(header)) && hasAny(headers, sentAliases),
    namePattern,
    namePattern.test("窗口期报告发送明细") ? "窗口期报告明细" : "阶段性报告明细",
    `需要“教师姓名、学生姓名、学号”以及发送结果列（支持：${sentAliases.join("、")}）`,
  );
}

function findSummarySheet(workbook: SheetJsWorkbook, key: "teacher" | "training") {
  const pattern = key === "teacher" ? /教师维度/u : /组长维度/u;
  return findSheet(
    workbook,
    (headers) => headers.includes(key === "teacher" ? "教师姓名" : "师训组长") &&
      hasStageTotal(headers) && hasWindowTotal(headers),
    pattern,
    key === "teacher" ? "教师维度汇总" : "师训组长维度汇总",
    `需要身份列以及阶段性、窗口期报告的需发送数和已发送数列`,
  );
}

function hasStageTotal(headers: string[]) {
  return hasAny(headers, STAGE_TOTAL_ALIASES) || headers.some((header) => /^阶段性报告[需应]发送(?:数)?\d{4}$/u.test(header));
}

function hasWindowTotal(headers: string[]) {
  return hasAny(headers, WINDOW_TOTAL_ALIASES) || headers.some((header) => /^窗口期报告[需应]发送(?:数)?\d{4}$/u.test(header));
}

function hasPeriodStageMetrics(headers: string[]) {
  return headers.some((header) => /^阶段性报告(?:[需应]发送|已发送|发送率)(?:数)?\d{4}$/u.test(header));
}

function findAppealSheet(workbook: SheetJsWorkbook) {
  return findSheet(
    workbook,
    (headers) => ["教师姓名", "学生姓名"].every((header) => headers.includes(header)) &&
      (headers.includes("申诉情况说明") || headers.includes("申诉情况详情")),
    /申诉/u,
    "申诉情况",
    "需要“教师姓名、学生姓名”以及“申诉情况说明”或“申诉情况详情”",
  );
}

function rowObjects(found: FoundSheet) {
  const columns = found.rows[0].map((value, index) => text(value) || `未命名列${index + 1}`);
  const rows = found.rows.slice(1)
    .filter((row) => row.some((value) => text(value)))
    .map((row) => columns.reduce<DataRow>((result, column, index) => {
      result[column] = row[index] ?? "";
      return result;
    }, {}));
  return { columns, rows };
}

function rateValue(value: unknown) {
  if (typeof value === "number") return value > 1 ? value / 100 : value;
  const raw = text(value).replace(/,/g, "");
  if (!raw) return "";
  const percent = raw.match(/^(-?\d+(?:\.\d+)?)%$/u);
  if (percent) return Number(percent[1]) / 100;
  const numeric = Number(raw);
  return Number.isFinite(numeric) ? numeric > 1 ? numeric / 100 : numeric : raw;
}

function countValue(value: unknown) {
  if (typeof value === "number") return value;
  const raw = text(value);
  const numeric = Number(raw.replace(/,/g, ""));
  return Number.isFinite(numeric) && raw !== "" ? numeric : value;
}

function numericCount(value: unknown) {
  const parsed = Number(countValue(value));
  return Number.isFinite(parsed) ? parsed : 0;
}

function isRateColumn(column: string) {
  return /发送率\d{0,4}$/u.test(column);
}

function isCountColumn(column: string) {
  return /(?:需发送|应发送|已发送|未发送)(?:数)?\d{0,4}$|申诉数\d{0,4}$/u.test(column);
}

function isStatusColumn(column: string) {
  return /是否完成|是否已通知/u.test(column);
}

function metricColumn(columns: string[], rateColumn: string, kind: "总发送" | "已发送") {
  const match = rateColumn.match(/^(.*)发送率(\d{4})?$/u);
  if (!match) return undefined;
  const [, prefix, period = ""] = match;
  const label = kind === "总发送" ? "[需应]发送" : kind;
  return columns.find((column) => new RegExp(`^${prefix}${label}(?:数)?${period}$`, "u").test(column));
}

function requiredColumn(columns: string[], aliases: readonly string[], label: string) {
  const column = aliases.find((alias) => columns.includes(alias));
  if (!column) throw new Error(`教师维度缺少“${label}”列。`);
  return column;
}

function transformSummary(found: FoundSheet, titleLabel: string, dataTime: string) {
  const { columns: sourceColumns, rows: sourceRows } = rowObjects(found);
  const columns = orderedSummaryColumns(sourceColumns.filter((column) => !/是否达标/u.test(column)));

  const rows = sourceRows.map((source) => {
    const row: DataRow = { ...source };
    sourceColumns.forEach((column) => {
      if (isRateColumn(column)) {
        const totalColumn = metricColumn(sourceColumns, column, "总发送");
        const sentColumn = metricColumn(sourceColumns, column, "已发送");
        const total = totalColumn ? Number(countValue(row[totalColumn])) : Number.NaN;
        const sent = sentColumn ? Number(countValue(row[sentColumn])) : Number.NaN;
        row[column] = Number.isFinite(total) && Number.isFinite(sent)
          ? total ? sent / total : 1
          : rateValue(row[column]);
      }
      else if (isCountColumn(column)) row[column] = countValue(row[column]);
    });
    return row;
  });
  return {
    name: titleLabel,
    title: `阶段性报告与窗口期报告发送率【${titleLabel}】（数据时间 ${dataTime}）`,
    rows,
    columns,
  };
}

function periodRank(column: string) {
  if (/窗口期报告/u.test(column)) return 0;
  if (/阶段性报告.*0819/u.test(column)) return 1;
  if (/阶段性报告.*0805/u.test(column)) return 2;
  return -1;
}

function orderedSummaryColumns(columns: string[]) {
  return [
    ...columns.filter((column) => periodRank(column) < 0),
    ...[0, 1, 2].flatMap((rank) => columns.filter((column) => periodRank(column) === rank)),
  ];
}

function compact(value: unknown) {
  return text(value).replace(/\s+/g, "");
}

function teacherIdentity(row: DataRow) {
  return compact(row.教师姓名 || row.授课教师 || row.老师姓名);
}

function hierarchyIdentity(row: DataRow) {
  return [teacherIdentity(row), compact(row.教研组), compact(row.师训组长), compact(row.助理主管)].join("\u0000");
}

function addIndex(index: Map<string, number[]>, key: string, rowIndex: number) {
  if (!key) return;
  if (!index.has(key)) index.set(key, []);
  index.get(key)!.push(rowIndex);
}

function buildHierarchyTeacherRows(teacherFound: FoundSheet, stageDetailRows: DataRow[]) {
  const source = rowObjects(teacherFound);
  const stageTotalColumn = requiredColumn(source.columns, STAGE_TOTAL_ALIASES, "阶段性报告需发送");
  const stageSentColumn = requiredColumn(source.columns, STAGE_SENT_COUNT_ALIASES, "阶段性报告已发送");
  const windowTotalColumn = requiredColumn(source.columns, WINDOW_TOTAL_ALIASES, "窗口期报告需发送");
  const windowSentColumn = requiredColumn(source.columns, WINDOW_SENT_COUNT_ALIASES, "窗口期报告已发送");
  ["教师姓名", "教研组", "师训组长", "助理主管"].forEach((column) => {
    if (!source.columns.includes(column)) throw new Error(`教师维度缺少“${column}”列。`);
  });

  const rows = source.rows
    .filter((row) => !/^(?:总计|合计)$/u.test(text(row.教师姓名)))
    .map((row): DataRow => ({
      教师姓名: row.教师姓名,
      教研组: row.教研组,
      师训组长: row.师训组长,
      助理主管: row.助理主管,
      阶段性报告应发送数: numericCount(row[stageTotalColumn]),
      阶段性报告已发送数: numericCount(row[stageSentColumn]),
      阶段性报告申诉数: 0,
      窗口期报告应发送数: numericCount(row[windowTotalColumn]),
      窗口期报告已发送数: numericCount(row[windowSentColumn]),
    }));

  const exactIndex = new Map<string, number[]>();
  const teacherIndex = new Map<string, number[]>();
  rows.forEach((row, index) => {
    addIndex(exactIndex, hierarchyIdentity(row), index);
    addIndex(teacherIndex, teacherIdentity(row), index);
  });
  stageDetailRows.filter(isStageReportAppealed).forEach((appealRow) => {
    const exactCandidates = exactIndex.get(hierarchyIdentity(appealRow)) || [];
    const teacherCandidates = teacherIndex.get(teacherIdentity(appealRow)) || [];
    const targetIndex = exactCandidates.length === 1
      ? exactCandidates[0]
      : teacherCandidates.length === 1
        ? teacherCandidates[0]
        : -1;
    if (targetIndex >= 0) {
      rows[targetIndex].阶段性报告申诉数 = Number(rows[targetIndex].阶段性报告申诉数) + 1;
      return;
    }
    const fallback: DataRow = {
      教师姓名: appealRow.教师姓名 || appealRow.授课教师 || appealRow.老师姓名 || "未填写",
      教研组: appealRow.教研组 || "未分组",
      师训组长: appealRow.师训组长 || "未填写",
      助理主管: appealRow.助理主管 || "未填写",
      阶段性报告应发送数: 0,
      阶段性报告已发送数: 0,
      阶段性报告申诉数: 1,
      窗口期报告应发送数: 0,
      窗口期报告已发送数: 0,
    };
    const fallbackIndex = rows.length;
    rows.push(fallback);
    addIndex(exactIndex, hierarchyIdentity(fallback), fallbackIndex);
    addIndex(teacherIndex, teacherIdentity(fallback), fallbackIndex);
  });
  return rows;
}

function requiredPeriodColumn(columns: string[], kind: "总发送" | "已发送", period: string) {
  const label = kind === "总发送" ? "[需应]发送" : kind;
  const column = columns.find((value) => new RegExp(`^阶段性报告${label}(?:数)?${period}$`, "u").test(value));
  if (!column) throw new Error(`教师维度缺少“阶段性报告应发送${period}”列。`);
  return column;
}

function buildPeriodHierarchyTeacherRows(teacherFound: FoundSheet, stageDetails: FoundSheet[]) {
  const source = rowObjects(teacherFound);
  const windowTotalColumn = requiredColumn(source.columns, WINDOW_TOTAL_ALIASES, "窗口期报告需发送");
  const windowSentColumn = requiredColumn(source.columns, WINDOW_SENT_COUNT_ALIASES, "窗口期报告已发送");
  ["教师姓名", "教研组", "师训组长", "助理主管"].forEach((column) => {
    if (!source.columns.includes(column)) throw new Error(`教师维度缺少“${column}”列。`);
  });
  const stageColumns = new Map(["0819", "0805"].map((period) => [period, {
    total: requiredPeriodColumn(source.columns, "总发送", period),
    sent: requiredPeriodColumn(source.columns, "已发送", period),
  }]));
  const rows = source.rows
    .filter((row) => !/^(?:总计|合计)$/u.test(text(row.教师姓名)))
    .map((row): DataRow => ({
      教师姓名: row.教师姓名,
      教研组: row.教研组,
      师训组长: row.师训组长,
      助理主管: row.助理主管,
      窗口期报告应发送数: numericCount(row[windowTotalColumn]),
      窗口期报告已发送数: numericCount(row[windowSentColumn]),
      ...Object.fromEntries([...stageColumns].flatMap(([period, columns]) => [
        [`阶段性报告应发送数${period}`, numericCount(row[columns.total])],
        [`阶段性报告已发送数${period}`, numericCount(row[columns.sent])],
        [`阶段性报告申诉数${period}`, 0],
      ])),
    }));
  const exactIndex = new Map<string, number[]>();
  const teacherIndex = new Map<string, number[]>();
  rows.forEach((row, index) => {
    addIndex(exactIndex, hierarchyIdentity(row), index);
    addIndex(teacherIndex, teacherIdentity(row), index);
  });
  stageDetails.forEach((detail) => {
    const period = periodCode(detail);
    const appealColumn = `阶段性报告申诉数${period}`;
    if (!period || !PERIOD_HIERARCHY_METRICS.some((metric) => metric.appealColumn === appealColumn)) return;
    rowObjects(detail).rows.filter(isStageReportAppealed).forEach((appealRow) => {
      const exactCandidates = exactIndex.get(hierarchyIdentity(appealRow)) || [];
      const teacherCandidates = teacherIndex.get(teacherIdentity(appealRow)) || [];
      const targetIndex = exactCandidates.length === 1 ? exactCandidates[0] : teacherCandidates.length === 1 ? teacherCandidates[0] : -1;
      if (targetIndex >= 0) {
        rows[targetIndex][appealColumn] = Number(rows[targetIndex][appealColumn]) + 1;
      }
    });
  });
  return rows;
}

function summaryCellStyle(row: DataRow, column: string, baseStyle: number) {
  if (isRateColumn(column)) return STYLE.rate;
  if (isStatusColumn(column)) return row[column] === "是" ? STYLE.sent : STYLE.unsent;
  return baseStyle;
}

function hierarchyStyle(row: DataRow) {
  if (row.__rowType === "grandTotal") return STYLE.grandTotal;
  if (row.__rowType === "projectTotal") return STYLE.projectTotal;
  if (row.__rowType === "groupTotal") return STYLE.groupTotal;
  return STYLE.detail;
}

function hierarchyCellStyle(row: DataRow, column: string, baseStyle: number) {
  if (!isRateColumn(column)) return baseStyle;
  if (row.__rowType === "grandTotal") return STYLE.grandTotalRate;
  if (row.__rowType === "projectTotal") return STYLE.projectTotalRate;
  if (row.__rowType === "groupTotal") return STYLE.groupTotalRate;
  return STYLE.detailRate;
}

function detailCellStyle(row: DataRow, column: string, baseStyle: number) {
  if (hasAny([column], [...STAGE_SENT_ALIASES, ...WINDOW_SENT_ALIASES])) {
    return text(row[column]) === "是" ? STYLE.sent : STYLE.unsent;
  }
  return baseStyle;
}

function detailWidths(columns: string[]) {
  return columns.reduce<Record<string, number>>((result, column) => {
    if (/申诉情况详情|申诉情况说明/u.test(column)) result[column] = 44;
    else if (/更新时间|时间/u.test(column)) result[column] = 25;
    else if (/学号|学生姓名|学员姓名/u.test(column)) result[column] = 18;
    else if (/是否发送|是否申诉|是否需要发送/u.test(column)) result[column] = 18;
    else result[column] = 18;
    return result;
  }, {});
}

function summaryWidths(columns: readonly string[]) {
  return columns.reduce<Record<string, number>>((result, column) => {
    if (/未发送学生姓名/u.test(column)) result[column] = 46;
    else if (/负责人/u.test(column)) result[column] = 44;
    else if (/教研组/u.test(column)) result[column] = 24;
    else if (/姓名|组长|主管/u.test(column)) result[column] = 20;
    else if (/^(?:窗口期报告|阶段性报告).*(?:应发送|已发送|发送率|未发送|申诉).*$/u.test(column)) result[column] = 24;
    else if (isRateColumn(column) || isStatusColumn(column)) result[column] = 20;
    else if (isCountColumn(column)) result[column] = 18;
    else if (/更新时间/u.test(column)) result[column] = 24;
    else result[column] = 18;
    return result;
  }, {});
}

function makeSummarySheet(summary: ReturnType<typeof transformSummary>): SheetDefinition {
  return {
    name: summary.name,
    title: summary.title,
    rows: summary.rows,
    columns: summary.columns,
    widths: summaryWidths(summary.columns),
    titleStyle: STYLE.title,
    headerStyle: STYLE.singleLineHeader,
    titleHeight: 42,
    headerHeight: 42,
    dataRowHeight: 24,
    rowStyle: () => STYLE.body,
    cellStyle: summaryCellStyle,
    dataBarColumns: summary.columns.filter(isRateColumn),
    dataBarColor: "C68A26",
  };
}

function makeHierarchySheet(
  name: "助理主管维度" | "教研组维度",
  table: HierarchyTable,
  columns: readonly string[],
  dataTime: string,
  metrics = HIERARCHY_METRICS,
): SheetDefinition {
  return {
    name,
    title: `阶段性报告与窗口期报告发送率【${name}】（数据时间 ${dataTime}）`,
    rows: table.rows,
    columns,
    widths: summaryWidths(columns),
    titleStyle: STYLE.title,
    headerStyle: STYLE.singleLineHeader,
    titleHeight: 42,
    headerHeight: 26,
    dataRowHeight: 24,
    mergeCells: table.mergeCells,
    rowStyle: hierarchyStyle,
    cellStyle: hierarchyCellStyle,
    dataBarColumns: metrics.map((metric) => metric.rateColumn),
    dataBarColor: "C68A26",
  };
}

function makeDetailSheet(name: string, found: FoundSheet): SheetDefinition {
  const detail = rowObjects(found);
  const columns = detail.columns.filter((column) => !/数据变动时间/u.test(column));
  const rows = detail.rows.map((source) => columns.reduce<DataRow>((result, column) => {
    result[column] = source[column] ?? "";
    return result;
  }, {}));
  return {
    name,
    rows,
    columns,
    widths: detailWidths(columns),
    headerStyle: STYLE.header,
    headerHeight: 38,
    dataRowHeight: 22,
    rowStyle: () => STYLE.detail,
    cellStyle: detailCellStyle,
  };
}

function periodCode(found: FoundSheet) {
  const match = found.name.match(/8[_月-]?(\d{1,2})/u);
  return match ? `08${match[1].padStart(2, "0")}` : "";
}

function detailName(found: FoundSheet, index: number) {
  const code = periodCode(found);
  return code ? `${code}阶段性报告明细` : `阶段性报告明细${index > 1 ? `第${index}批` : ""}`;
}

function appealName(found: FoundSheet, index: number) {
  const code = periodCode(found);
  return code ? `${code}阶段性报告申诉情况` : `阶段性报告申诉情况${index > 1 ? `第${index}批` : ""}`;
}

function findAllSheets(workbook: SheetJsWorkbook, required: (headers: string[]) => boolean, pattern: RegExp) {
  return workbook.SheetNames.flatMap((name) => {
    const rows = rowsForSheet(workbook, name);
    return rows.length && pattern.test(name) && required(rows[0].map(text)) ? [{ name, rows }] : [];
  });
}

function buildPeriodReportOutput(workbook: SheetJsWorkbook, dataTime: string) {
  const stageDetails = findAllSheets(
    workbook,
    (headers) => ["教师姓名", "学生姓名", "学号"].every((header) => headers.includes(header)) && hasAny(headers, STAGE_SENT_ALIASES),
    /非窗口期|暑期在读|阶段性报告发送明细/u,
  ).sort((left, right) => periodCode(right).localeCompare(periodCode(left)));
  const appeals = findAllSheets(
    workbook,
    (headers) => ["教师姓名", "学生姓名"].every((header) => headers.includes(header)) &&
      (headers.includes("申诉情况说明") || headers.includes("申诉情况详情")),
    /申诉/u,
  ).sort((left, right) => periodCode(right).localeCompare(periodCode(left)));
  const teacherFound = findSummarySheet(workbook, "teacher");
  const hierarchyTeacherRows = buildPeriodHierarchyTeacherRows(teacherFound, stageDetails);
  const group = buildResearchGroupHierarchy(hierarchyTeacherRows, PERIOD_HIERARCHY_METRICS, false);
  const assistant = buildAssistantHierarchy(hierarchyTeacherRows, PERIOD_HIERARCHY_METRICS);
  const training = transformSummary(findSummarySheet(workbook, "training"), "师训组长维度", dataTime);
  const teacher = transformSummary(teacherFound, "教师维度", dataTime);
  const windowDetail = findDetailSheet(workbook, WINDOW_SENT_ALIASES, /窗口期报告发送明细/u);
  const sheets: SheetDefinition[] = [
    makeDetailSheet("窗口期报告明细", windowDetail),
    ...stageDetails.map((detail, index) => makeDetailSheet(detailName(detail, index + 1), detail)),
    ...appeals.map((appeal, index) => ({
      name: appealName(appeal, index + 1),
      title: `${appealName(appeal, index + 1)}（数据时间 ${dataTime}）`,
      rows: rowObjects(appeal).rows,
      columns: rowObjects(appeal).columns,
      widths: detailWidths(rowObjects(appeal).columns),
      titleStyle: STYLE.title,
      headerStyle: STYLE.header,
      titleHeight: 42,
      headerHeight: 38,
      dataRowHeight: 24,
      rowStyle: () => STYLE.detail,
    })),
    makeHierarchySheet("教研组维度", group, PERIOD_RESEARCH_GROUP_COLUMNS, dataTime, PERIOD_HIERARCHY_METRICS),
    makeHierarchySheet("助理主管维度", assistant, PERIOD_ASSISTANT_COLUMNS, dataTime, PERIOD_HIERARCHY_METRICS),
    makeSummarySheet(training),
    makeSummarySheet(teacher),
  ];
  return {
    buffer: buildWorkbook(sheets),
    dataTime,
    counts: {
      stageRows: stageDetails.reduce((total, detail) => total + rowObjects(detail).rows.length, 0),
      windowRows: rowObjects(windowDetail).rows.length,
      teacherRows: teacher.rows.length,
      appealRows: appeals.reduce((total, appeal) => total + rowObjects(appeal).rows.length, 0),
      sheets: sheets.length,
    },
  };
}

export function buildStageReportBeautifyOutput(workbook: SheetJsWorkbook) {
  const dataTime = latestUpdateTime(workbook);
  if (workbook.SheetNames.some((name) => hasPeriodStageMetrics(rowsForSheet(workbook, name)[0]?.map(text) || []))) {
    return buildPeriodReportOutput(workbook, dataTime);
  }
  const stageDetail = findDetailSheet(workbook, STAGE_SENT_ALIASES, /非窗口期|暑期在读|阶段性报告发送明细/u);
  const windowDetail = findDetailSheet(workbook, WINDOW_SENT_ALIASES, /窗口期报告发送明细/u);
  const teacherFound = findSummarySheet(workbook, "teacher");
  const teacher = transformSummary(teacherFound, "教师维度", dataTime);
  const training = transformSummary(findSummarySheet(workbook, "training"), "师训组长维度", dataTime);
  const stageDetailRows = rowObjects(stageDetail).rows;
  const hierarchyTeacherRows = buildHierarchyTeacherRows(teacherFound, stageDetailRows);
  const assistant = buildAssistantHierarchy(hierarchyTeacherRows, HIERARCHY_METRICS);
  const group = buildResearchGroupHierarchy(hierarchyTeacherRows, HIERARCHY_METRICS, false);
  const appeal = rowObjects(findAppealSheet(workbook));
  const sheets: SheetDefinition[] = [
    makeDetailSheet("阶段性报告明细", stageDetail),
    makeDetailSheet("窗口期报告明细", windowDetail),
    {
      name: "阶段性报告申诉情况",
      title: `阶段性报告申诉情况（数据时间 ${dataTime}）`,
      rows: appeal.rows,
      columns: appeal.columns,
      widths: detailWidths(appeal.columns),
      titleStyle: STYLE.title,
      headerStyle: STYLE.header,
      titleHeight: 42,
      headerHeight: 38,
      dataRowHeight: 24,
      rowStyle: () => STYLE.detail,
    },
    makeHierarchySheet("教研组维度", group, RESEARCH_GROUP_COLUMNS, dataTime),
    makeHierarchySheet("助理主管维度", assistant, ASSISTANT_COLUMNS, dataTime),
    makeSummarySheet(training),
    makeSummarySheet(teacher),
  ];
  const buffer = buildWorkbook(sheets);
  return {
    buffer,
    dataTime,
    counts: {
      stageRows: stageDetailRows.length,
      windowRows: rowObjects(windowDetail).rows.length,
      teacherRows: teacher.rows.length,
      appealRows: appeal.rows.length,
      sheets: sheets.length,
    },
  };
}
