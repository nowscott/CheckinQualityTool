import { buildWorkbook } from "./excelWriter";
import { REMINDER_PASS_RATE } from "./reminderConfig";
import { findTrainingLeadGroups } from "./stageReportGroups";
import type { DataRow, SheetDefinition } from "./types";
import { normalizeMatchText, text } from "./utils";
import type { StageReportListInfo } from "./stageReportListParser";
import type { StageReportMatchInfo } from "./stageReportMatching";

const TEACHER_COLUMNS = [
  "教师姓名", "教研组", "师训组长", "助理主管",
  "应发送数", "已发送数", "发送率", "未发送数", "申诉数", "是否达标（80%）",
] as const;

const TRAINING_COLUMNS = [
  "师训组长", "教研组", "应发送数", "已发送数", "发送率", "未发送数", "申诉数", "是否达标（80%）",
] as const;

const ASSISTANT_COLUMNS = [
  "教研组", "师训主管/助理主管", "应发送数", "已发送数", "发送率", "未发送数", "申诉数", "是否达标（80%）",
] as const;

const RESEARCH_GROUP_COLUMNS = [
  "教研组", "负责人", "应发送数", "已发送数", "发送率", "未发送数", "申诉数", "是否达标（80%）",
] as const;

const PROJECT_ORDER = ["博文项目", "双语项目", "益智项目", "文理综项目", "其他项目"];

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
  assistants: Set<string>;
  groups: Set<string>;
}

interface PublicTable {
  rows: DataRow[];
  mergeCells: string[];
}

function createBucket(): SummaryBucket {
  return { total: 0, sent: 0, appealed: 0, assistants: new Set<string>(), groups: new Set<string>() };
}

function mergeBucket(target: SummaryBucket, source: SummaryBucket) {
  target.total += source.total;
  target.sent += source.sent;
  target.appealed += source.appealed;
  source.assistants.forEach((assistant) => target.assistants.add(assistant));
  source.groups.forEach((group) => target.groups.add(group));
}

function addSummaryRow(bucket: SummaryBucket, row: DataRow) {
  bucket.total += Number(row.应发送数) || 0;
  bucket.sent += Number(row.已发送数) || 0;
  bucket.appealed += Number(row.申诉数) || 0;
  const assistant = text(row.助理主管);
  const group = text(row.教研组);
  if (assistant) bucket.assistants.add(assistant);
  if (group) bucket.groups.add(group);
}

function compareText(a: unknown, b: unknown) {
  return text(a).localeCompare(text(b), "zh-CN");
}

function projectForResearchGroup(value: unknown) {
  const group = text(value).replace(/\s+/g, "");
  if (!group || group === "跨教研组" || group === "未分组") return "其他项目";
  if (/(博文|实验)[A-Za-zＡ-Ｚａ-ｚ]/u.test(group) || /文综|理综|文理综|政史地生|政史地|史地生/u.test(group)) return "文理综项目";
  if (group.includes("博文")) return "博文项目";
  if (group.includes("双语")) return "双语项目";
  if (group.includes("益智")) return "益智项目";
  return "其他项目";
}

function projectCompare(a: string, b: string) {
  return (PROJECT_ORDER.indexOf(a) === -1 ? 99 : PROJECT_ORDER.indexOf(a)) -
    (PROJECT_ORDER.indexOf(b) === -1 ? 99 : PROJECT_ORDER.indexOf(b)) || compareText(a, b);
}

function rate(sent: number, total: number) {
  return total ? sent / total : 1;
}

function passText(sent: number, total: number) {
  return rate(sent, total) >= REMINDER_PASS_RATE ? "是" : "否";
}

function metrics(bucket: SummaryBucket) {
  return {
    应发送数: bucket.total,
    已发送数: bucket.sent,
    发送率: rate(bucket.sent, bucket.total),
    未发送数: Math.max(0, bucket.total - bucket.sent),
    申诉数: bucket.appealed,
    "是否达标（80%）": passText(bucket.sent, bucket.total),
  };
}

function normalizePersonName(value: unknown) {
  return normalizeMatchText(value).replace(/\s+/g, "").replace(/[0-9０-９]+$/u, "");
}

function displayPersonName(value: unknown) {
  return text(value).replace(/[0-9０-９]+$/u, "");
}

function displayPersonNames(values: Iterable<string>) {
  return [...new Set([...values].map(displayPersonName).filter(Boolean))].sort(compareText).join("、") || "/";
}

function assistantTeachingGroup(value: unknown) {
  const group = text(value) || "未分组";
  const project = projectForResearchGroup(group);
  const compact = group.replace(/\s+/g, "");
  if (project !== "文理综项目") return group;
  if (compact.includes("实验P")) return "实验P";
  if (compact.includes("实验C")) return "实验C";
  return "政史地生";
}

function assistantTeachingGroupCompare(a: string, b: string) {
  const order = ["实验P", "实验C", "政史地生"];
  const ai = order.indexOf(a);
  const bi = order.indexOf(b);
  if (ai !== -1 || bi !== -1) return (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi) || compareText(a, b);
  return compareText(a, b);
}

function isAppealed(row: DataRow) {
  const status = text(row.是否申诉).toLocaleLowerCase("zh-CN");
  if (status && !["否", "no", "n", "未申诉"].includes(status)) return true;
  return Boolean(text(row.申诉情况详情 || row.申诉说明));
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
    if (isAppealed(row)) item.申诉数 = Number(item.申诉数) + 1;
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

function buildTrainingRows(teacherRows: DataRow[]) {
  const grouped = new Map<string, SummaryBucket>();
  teacherRows.forEach((row) => {
    const lead = text(row.师训组长) || "未填写";
    if (!grouped.has(lead)) grouped.set(lead, createBucket());
    addSummaryRow(grouped.get(lead)!, row);
  });
  return [...grouped.entries()]
    .map(([lead, bucket]): DataRow => ({
      师训组长: lead,
      教研组: findTrainingLeadGroups(teacherRows, lead).join("、") || "/",
      ...metrics(bucket),
    }))
    .sort((a, b) => Number(b.发送率) - Number(a.发送率) || compareText(a.师训组长, b.师训组长));
}

function assistantOwnTeachingGroups(teacherRows: DataRow[]) {
  const groups = new Map<string, string>();
  teacherRows.forEach((row) => {
    const teacher = normalizePersonName(row.教师姓名);
    const assistant = normalizePersonName(row.助理主管);
    if (teacher && assistant && teacher === assistant && row.教研组) groups.set(assistant, assistantTeachingGroup(row.教研组));
  });
  return groups;
}

function mergeRef(columnIndex: number, startRow: number, endRow: number) {
  const column = String.fromCharCode(65 + columnIndex);
  return `${column}${startRow}:${column}${endRow}`;
}

function buildAssistantTable(teacherRows: DataRow[]): PublicTable {
  const grouped = new Map<string, Map<string, Map<string, SummaryBucket>>>();
  const ownGroups = assistantOwnTeachingGroups(teacherRows);
  teacherRows.forEach((row) => {
    const assistant = text(row.助理主管) || "未填写";
    const group = ownGroups.get(normalizePersonName(assistant)) || assistantTeachingGroup(row.教研组);
    const project = projectForResearchGroup(group);
    if (!grouped.has(project)) grouped.set(project, new Map());
    const groups = grouped.get(project)!;
    if (!groups.has(group)) groups.set(group, new Map());
    const assistants = groups.get(group)!;
    if (!assistants.has(assistant)) assistants.set(assistant, createBucket());
    addSummaryRow(assistants.get(assistant)!, row);
  });

  const rows: DataRow[] = [];
  const mergeCells: string[] = [];
  const actualRow = (index: number) => index + 3;
  [...grouped.keys()].sort(projectCompare).forEach((project) => {
    const projectBucket = createBucket();
    const groups = grouped.get(project)!;
    [...groups.keys()].sort(assistantTeachingGroupCompare).forEach((group) => {
      const start = rows.length;
      const groupBucket = createBucket();
      const assistants = groups.get(group)!;
      [...assistants.keys()].sort(compareText).forEach((assistant) => {
        const item = assistants.get(assistant)!;
        mergeBucket(groupBucket, item);
        rows.push({
          教研组: rows.length === start ? group : "",
          "师训主管/助理主管": displayPersonName(assistant),
          ...metrics(item),
          __rowType: "detail",
        });
      });
      if (rows.length - start > 1) mergeCells.push(mergeRef(0, actualRow(start), actualRow(rows.length - 1)));
      const summaryRow = actualRow(rows.length);
      mergeCells.push(`A${summaryRow}:B${summaryRow}`);
      rows.push({ 教研组: group, "师训主管/助理主管": "", ...metrics(groupBucket), __rowType: "groupTotal" });
      mergeBucket(projectBucket, groupBucket);
    });
    const projectRow = actualRow(rows.length);
    mergeCells.push(`A${projectRow}:B${projectRow}`);
    rows.push({ 教研组: project, "师训主管/助理主管": "", ...metrics(projectBucket), __rowType: "projectTotal" });
  });
  const total = createBucket();
  rows.filter((row) => row.__rowType === "projectTotal").forEach((row) => {
    total.total += Number(row.应发送数) || 0;
    total.sent += Number(row.已发送数) || 0;
    total.appealed += Number(row.申诉数) || 0;
  });
  const totalRow = actualRow(rows.length);
  mergeCells.push(`A${totalRow}:B${totalRow}`);
  rows.push({ 教研组: "总计", "师训主管/助理主管": "", ...metrics(total), __rowType: "grandTotal" });
  return { rows, mergeCells };
}

function buildResearchGroupTable(teacherRows: DataRow[]): PublicTable {
  const projects = new Map<string, Map<string, SummaryBucket>>();
  teacherRows.forEach((row) => {
    const group = assistantTeachingGroup(row.教研组);
    const project = projectForResearchGroup(group);
    if (!projects.has(project)) projects.set(project, new Map());
    const groups = projects.get(project)!;
    if (!groups.has(group)) groups.set(group, createBucket());
    addSummaryRow(groups.get(group)!, row);
  });

  const rows: DataRow[] = [];
  const mergeCells: string[] = [];
  const actualRow = (index: number) => index + 3;
  [...projects.keys()].sort(projectCompare).forEach((project) => {
    const projectBucket = createBucket();
    const groups = projects.get(project)!;
    [...groups.keys()].sort(assistantTeachingGroupCompare).forEach((group) => {
      const item = groups.get(group)!;
      mergeBucket(projectBucket, item);
      rows.push({
        教研组: group,
        负责人: displayPersonNames(item.assistants),
        ...metrics(item),
        __rowType: "detail",
      });
    });
    const projectRow = actualRow(rows.length);
    mergeCells.push(`A${projectRow}:B${projectRow}`);
    rows.push({ 教研组: project, 负责人: "", ...metrics(projectBucket), __rowType: "projectTotal" });
  });
  const total = createBucket();
  rows.filter((row) => row.__rowType === "projectTotal").forEach((row) => {
    total.total += Number(row.应发送数) || 0;
    total.sent += Number(row.已发送数) || 0;
    total.appealed += Number(row.申诉数) || 0;
  });
  const totalRow = actualRow(rows.length);
  mergeCells.push(`A${totalRow}:B${totalRow}`);
  rows.push({ 教研组: "总计", 负责人: "", ...metrics(total), __rowType: "grandTotal" });
  return { rows, mergeCells };
}

function referenceCellStyle(row: DataRow, column: string) {
  if (column === "发送率") return STYLE.referenceRate;
  if (column === "是否达标（80%）") return row[column] === "是" ? STYLE.sent : STYLE.unsent;
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
  const trainingRows = buildTrainingRows(teacherRows);
  const assistantTable = buildAssistantTable(teacherRows);
  const researchGroupTable = buildResearchGroupTable(teacherRows);
  const sheets: SheetDefinition[] = [
    {
      name: "师训组长维度",
      title: summaryTitle("师训组长维度"),
      rows: trainingRows,
      columns: TRAINING_COLUMNS,
      widths: { 师训组长: 18, 教研组: 28, 应发送数: 14, 已发送数: 14, 发送率: 18, 未发送数: 14, 申诉数: 14, "是否达标（80%）": 18 },
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
      widths: { 教研组: 18, "师训主管/助理主管": 22, 应发送数: 14, 已发送数: 14, 发送率: 18, 未发送数: 14, 申诉数: 14, "是否达标（80%）": 18 },
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
      widths: { 教研组: 18, 负责人: 60, 应发送数: 14, 已发送数: 14, 发送率: 18, 未发送数: 14, 申诉数: 14, "是否达标（80%）": 18 },
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
      widths: { 教师姓名: 18, 教研组: 18, 师训组长: 18, 助理主管: 18, 应发送数: 14, 已发送数: 14, 发送率: 18, 未发送数: 14, 申诉数: 14, "是否达标（80%）": 18 },
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
