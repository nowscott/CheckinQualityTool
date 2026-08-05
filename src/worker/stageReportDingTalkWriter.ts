import { buildWorkbook } from "./excelWriter";
import { findTrainingLeadGroups, trainingPersonKey } from "./stageReportGroups";
import {
  buildAssistantHierarchy,
  buildResearchGroupHierarchy,
  isStageReportAppealed,
  type HierarchyMetricSpec,
} from "./stageReportHierarchy";
import type { DataRow, SheetDefinition } from "./types";
import { text } from "./utils";
import type { StageReportListInfo } from "./stageReportListParser";
import type { StageReportMatchInfo } from "./stageReportMatching";

const TEACHER_COLUMNS = [
  "教师姓名", "教研组", "师训组长", "助理主管",
  "应发送数", "已发送数", "发送率", "未发送数", "申诉数",
] as const;

const TRAINING_COLUMNS = [
  "师训组长", "教研组", "应发送数", "已发送数", "发送率", "未发送数", "申诉数",
] as const;

const ASSISTANT_COLUMNS = [
  "教研组", "师训主管/助理主管", "应发送数", "已发送数", "发送率", "未发送数", "申诉数",
] as const;

const RESEARCH_GROUP_COLUMNS = [
  "教研组", "负责人", "应发送数", "已发送数", "发送率", "未发送数", "申诉数",
] as const;

const HIERARCHY_METRICS: readonly HierarchyMetricSpec[] = [{
  totalColumn: "应发送数",
  sentColumn: "已发送数",
  rateColumn: "发送率",
  unsentColumn: "未发送数",
  appealColumn: "申诉数",
}];

const STYLE = {
  title: 16,
  header: 17,
  referenceBody: 18,
  referenceRate: 19,
  hierarchyDetail: 20,
  hierarchyDetailRate: 21,
  groupTotal: 22,
  groupTotalRate: 23,
  projectTotal: 24,
  projectTotalRate: 25,
  grandTotal: 26,
  grandTotalRate: 27,
  sent: 28,
  unsent: 29,
} as const;

interface SummaryBucket {
  total: number;
  sent: number;
  appealed: number;
}

function createBucket(): SummaryBucket {
  return { total: 0, sent: 0, appealed: 0 };
}

function addSummaryRow(bucket: SummaryBucket, row: DataRow) {
  bucket.total += Number(row.应发送数) || 0;
  bucket.sent += Number(row.已发送数) || 0;
  bucket.appealed += Number(row.申诉数) || 0;
}

function compareText(a: unknown, b: unknown) {
  return text(a).localeCompare(text(b), "zh-CN");
}

function rate(sent: number, total: number) {
  return total ? sent / total : 1;
}

function metrics(bucket: SummaryBucket) {
  return {
    应发送数: bucket.total,
    已发送数: bucket.sent,
    发送率: rate(bucket.sent, bucket.total),
    未发送数: Math.max(0, bucket.total - bucket.sent),
    申诉数: bucket.appealed,
  };
}

function findSentColumn(columns: string[]) {
  const aliases = ["是否发送阶段性报告", "阶段性报告是否发送"];
  const column = columns.find((name) => aliases.includes(text(name)));
  if (!column) throw new Error("钉钉表单分母缺少“是否发送阶段性报告”列。");
  return column;
}

function buildDetailRows(listInfo: StageReportListInfo, matchInfo: StageReportMatchInfo, sentColumn: string) {
  return matchInfo.detailRows.map((row) => ({
    ...row,
    [sentColumn]: row.本次检查结论 === "已发送" ? "是" : "否",
  }));
}

function buildTeacherRows(detailRows: DataRow[], sentColumn: string) {
  const grouped = new Map<string, DataRow>();
  detailRows.forEach((row) => {
    const teacher = text(row.教师姓名 || row.授课教师 || row.老师姓名);
    const teachingGroup = text(row.教研组) || "未分组";
    const trainingLead = text(row.师训组长) || "未填写";
    const assistant = text(row.助理主管) || "未填写";
    const key = [teacher, teachingGroup, trainingLead, assistant].join("\u0000");
    if (!grouped.has(key)) {
      grouped.set(key, {
        教师姓名: teacher,
        教研组: teachingGroup,
        师训组长: trainingLead,
        助理主管: assistant,
        应发送数: 0,
        已发送数: 0,
        申诉数: 0,
      });
    }
    const item = grouped.get(key)!;
    item.应发送数 = Number(item.应发送数) + 1;
    if (text(row[sentColumn]) === "是") item.已发送数 = Number(item.已发送数) + 1;
    if (isStageReportAppealed(row)) item.申诉数 = Number(item.申诉数) + 1;
  });
  return [...grouped.values()]
    .map((row): DataRow => {
      const bucket = createBucket();
      bucket.total = Number(row.应发送数) || 0;
      bucket.sent = Number(row.已发送数) || 0;
      bucket.appealed = Number(row.申诉数) || 0;
      return { ...row, ...metrics(bucket) };
    })
    .sort((a, b) => Number(b.发送率) - Number(a.发送率) || compareText(a.教师姓名, b.教师姓名));
}

function buildTrainingRows(teacherRows: DataRow[], fallbackLeadGroups: Map<string, string[]> = new Map()) {
  const grouped = new Map<string, SummaryBucket>();
  teacherRows.forEach((row) => {
    const lead = text(row.师训组长) || "未填写";
    if (!grouped.has(lead)) grouped.set(lead, createBucket());
    addSummaryRow(grouped.get(lead)!, row);
  });
  return [...grouped.entries()]
    .map(([lead, bucket]): DataRow => ({
      师训组长: lead,
      教研组: fallbackLeadGroups.get(trainingPersonKey(lead))?.join("、") || findTrainingLeadGroups(teacherRows, lead).join("、") || "/",
      ...metrics(bucket),
    }))
    .sort((a, b) => Number(b.发送率) - Number(a.发送率) || compareText(a.师训组长, b.师训组长));
}

function referenceCellStyle(row: DataRow, column: string) {
  if (column === "发送率") return STYLE.referenceRate;
  return STYLE.referenceBody;
}

function hierarchyStyle(row: DataRow) {
  if (row.__rowType === "grandTotal") return STYLE.grandTotal;
  if (row.__rowType === "projectTotal") return STYLE.projectTotal;
  if (row.__rowType === "groupTotal") return STYLE.groupTotal;
  return STYLE.hierarchyDetail;
}

function hierarchyCellStyle(row: DataRow, column: string, baseStyle: number) {
  if (column !== "发送率") return baseStyle;
  if (row.__rowType === "grandTotal") return STYLE.grandTotalRate;
  if (row.__rowType === "projectTotal") return STYLE.projectTotalRate;
  if (row.__rowType === "groupTotal") return STYLE.groupTotalRate;
  return STYLE.hierarchyDetailRate;
}

function detailCellStyle(row: DataRow, column: string, baseStyle: number, sentColumn: string) {
  if (column !== sentColumn) return baseStyle;
  return text(row[sentColumn]) === "是" ? STYLE.sent : STYLE.unsent;
}

function generatedDate() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function summaryTitle(dimension: string) {
  return `阶段性报告发送率【${dimension}】（数据生成 ${generatedDate()}）`;
}

export function buildStageReportDingTalkOutput(
  listInfo: StageReportListInfo,
  matchInfo: StageReportMatchInfo,
) {
  const sentColumn = findSentColumn(listInfo.columns);
  const detailRows = buildDetailRows(listInfo, matchInfo, sentColumn);
  const teacherRows = buildTeacherRows(detailRows, sentColumn);
  const trainingRows = buildTrainingRows(teacherRows, listInfo.trainingLeadGroups);
  const assistantTable = buildAssistantHierarchy(teacherRows, HIERARCHY_METRICS);
  const researchGroupTable = buildResearchGroupHierarchy(teacherRows, HIERARCHY_METRICS);
  const sheets: SheetDefinition[] = [
    {
      name: "师训组长维度",
      title: summaryTitle("师训组长维度"),
      rows: trainingRows,
      columns: TRAINING_COLUMNS,
      widths: { 师训组长: 18, 教研组: 28, 应发送数: 14, 已发送数: 14, 发送率: 18, 未发送数: 14, 申诉数: 14 },
      titleStyle: STYLE.title,
      headerStyle: STYLE.header,
      titleHeight: 42,
      headerHeight: 42,
      dataRowHeight: 24,
      rowStyle: () => STYLE.referenceBody,
      cellStyle: referenceCellStyle,
      dataBarColumns: ["发送率"],
      dataBarColor: "C68A26",
    },
    {
      name: "助理主管维度",
      title: summaryTitle("助理主管维度"),
      rows: assistantTable.rows,
      columns: ASSISTANT_COLUMNS,
      widths: { 教研组: 18, "师训主管/助理主管": 22, 应发送数: 14, 已发送数: 14, 发送率: 18, 未发送数: 14, 申诉数: 14 },
      titleStyle: STYLE.title,
      headerStyle: STYLE.header,
      titleHeight: 42,
      headerHeight: 42,
      dataRowHeight: 24,
      mergeCells: assistantTable.mergeCells,
      rowStyle: hierarchyStyle,
      cellStyle: hierarchyCellStyle,
      dataBarColumns: ["发送率"],
      dataBarColor: "C68A26",
    },
    {
      name: "教研组维度",
      title: summaryTitle("教研组维度"),
      rows: researchGroupTable.rows,
      columns: RESEARCH_GROUP_COLUMNS,
      widths: { 教研组: 18, 负责人: 60, 应发送数: 14, 已发送数: 14, 发送率: 18, 未发送数: 14, 申诉数: 14 },
      titleStyle: STYLE.title,
      headerStyle: STYLE.header,
      titleHeight: 42,
      headerHeight: 42,
      dataRowHeight: 24,
      mergeCells: researchGroupTable.mergeCells,
      rowStyle: hierarchyStyle,
      cellStyle: hierarchyCellStyle,
      dataBarColumns: ["发送率"],
      dataBarColor: "C68A26",
    },
    {
      name: "教师维度",
      title: summaryTitle("教师维度"),
      rows: teacherRows,
      columns: TEACHER_COLUMNS,
      widths: { 教师姓名: 18, 教研组: 18, 师训组长: 18, 助理主管: 18, 应发送数: 14, 已发送数: 14, 发送率: 18, 未发送数: 14, 申诉数: 14 },
      titleStyle: STYLE.title,
      headerStyle: STYLE.header,
      titleHeight: 42,
      headerHeight: 42,
      dataRowHeight: 24,
      rowStyle: () => STYLE.referenceBody,
      cellStyle: referenceCellStyle,
      dataBarColumns: ["发送率"],
      dataBarColor: "C68A26",
    },
    {
      name: "发送明细",
      rows: detailRows,
      columns: listInfo.columns,
      widths: {
        教师姓名: 18, 学生姓名: 18, 学员姓名: 18, 学号: 18, 暑假最后一节课时间: 24,
        师训组长: 18, 助理主管: 18, 教研组: 18, 是否发送阶段性报告: 20,
        是否申诉: 14, 申诉情况详情: 44, 是否需要发送: 14,
      },
      headerStyle: STYLE.header,
      headerHeight: 36,
      dataRowHeight: 22,
      rowStyle: () => STYLE.hierarchyDetail,
      cellStyle: (row, column, baseStyle) => detailCellStyle(row, column, baseStyle, sentColumn),
    },
  ];
  const sheetOrder = ["发送明细", "教师维度", "师训组长维度", "助理主管维度", "教研组维度"];
  sheets.sort((a, b) => sheetOrder.indexOf(a.name) - sheetOrder.indexOf(b.name));
  return {
    output: buildWorkbook(sheets),
    counts: {
      rows: detailRows.length,
      sent: matchInfo.counts.已发送数,
      unsent: matchInfo.counts.未发送数,
      missing: matchInfo.counts.字段缺失数,
      teachers: teacherRows.length,
    },
    sheetName: listInfo.sheetName,
  };
}
