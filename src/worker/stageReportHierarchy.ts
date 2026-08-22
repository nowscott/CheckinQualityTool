import type { DataRow } from "./types";
import { normalizeMatchText, text } from "./utils";
import { reminderProjectGroup } from "./reminderProjectGroup";

export interface HierarchyMetricSpec {
  totalColumn: string;
  sentColumn: string;
  rateColumn: string;
  unsentColumn: string;
  appealColumn?: string;
}

export interface HierarchyTable {
  rows: DataRow[];
  mergeCells: string[];
}

interface MetricBucket {
  total: number;
  sent: number;
  appealed: number;
}

interface HierarchyBucket {
  metrics: MetricBucket[];
  assistants: Set<string>;
}

const PROJECT_ORDER = ["博文项目", "双语项目", "益智项目", "文理综项目", "其他项目"];

function compareText(a: unknown, b: unknown) {
  return text(a).localeCompare(text(b), "zh-CN");
}

export function normalizePersonName(value: unknown) {
  return normalizeMatchText(value).replace(/\s+/g, "").replace(/[0-9０-９]+$/u, "");
}

function displayPersonName(value: unknown) {
  return text(value).replace(/[0-9０-９]+$/u, "");
}

export function displayPersonNames(values: Iterable<string>) {
  return [...new Set([...values].map(displayPersonName).filter(Boolean))].sort(compareText).join("、") || "/";
}

export function projectForResearchGroup(value: unknown) {
  return reminderProjectGroup(value);
}

function projectCompare(a: string, b: string) {
  return (PROJECT_ORDER.indexOf(a) === -1 ? 99 : PROJECT_ORDER.indexOf(a)) -
    (PROJECT_ORDER.indexOf(b) === -1 ? 99 : PROJECT_ORDER.indexOf(b)) || compareText(a, b);
}

export function assistantTeachingGroup(value: unknown) {
  const group = text(value) || "未分组";
  const project = projectForResearchGroup(group);
  const compact = group.replace(/\s+/g, "");
  if (project !== "文理综项目") return group;
  if (compact.includes("实验P")) return "实验P";
  if (compact.includes("实验C")) return "实验C";
  return "政史地生";
}

export function assistantTeachingGroupCompare(a: string, b: string) {
  const order = ["实验P", "实验C", "政史地生"];
  const ai = order.indexOf(a);
  const bi = order.indexOf(b);
  if (ai !== -1 || bi !== -1) return (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi) || compareText(a, b);
  return compareText(a, b);
}

export function isStageReportAppealed(row: DataRow) {
  const effective = text(row.申诉是否生效).toLocaleLowerCase("zh-CN");
  if (effective) return ["申诉生效", "有效", "是", "通过"].includes(effective);
  const status = text(row.是否申诉).toLocaleLowerCase("zh-CN");
  if (status && !["否", "no", "n", "未申诉"].includes(status)) return true;
  return Boolean(text(row.申诉情况详情 || row.申诉说明));
}

function createBucket(metricSpecs: readonly HierarchyMetricSpec[]): HierarchyBucket {
  return {
    metrics: metricSpecs.map(() => ({ total: 0, sent: 0, appealed: 0 })),
    assistants: new Set<string>(),
  };
}

function addSummaryRow(bucket: HierarchyBucket, row: DataRow, metricSpecs: readonly HierarchyMetricSpec[]) {
  metricSpecs.forEach((spec, index) => {
    bucket.metrics[index].total += Number(row[spec.totalColumn]) || 0;
    bucket.metrics[index].sent += Number(row[spec.sentColumn]) || 0;
    if (spec.appealColumn) bucket.metrics[index].appealed += Number(row[spec.appealColumn]) || 0;
  });
  const assistant = text(row.助理主管);
  if (assistant) bucket.assistants.add(assistant);
}

function mergeBucket(target: HierarchyBucket, source: HierarchyBucket) {
  source.metrics.forEach((metric, index) => {
    target.metrics[index].total += metric.total;
    target.metrics[index].sent += metric.sent;
    target.metrics[index].appealed += metric.appealed;
  });
  source.assistants.forEach((assistant) => target.assistants.add(assistant));
}

function rate(sent: number, total: number) {
  return total ? sent / total : 1;
}

function metricValues(bucket: HierarchyBucket, metricSpecs: readonly HierarchyMetricSpec[]) {
  return metricSpecs.reduce<DataRow>((result, spec, index) => {
    const metric = bucket.metrics[index];
    const currentRate = rate(metric.sent, metric.total);
    result[spec.totalColumn] = metric.total;
    result[spec.sentColumn] = metric.sent;
    result[spec.rateColumn] = currentRate;
    result[spec.unsentColumn] = Math.max(0, metric.total - metric.sent);
    if (spec.appealColumn) result[spec.appealColumn] = metric.appealed;
    return result;
  }, {});
}

export function assistantOwnTeachingGroups(teacherRows: DataRow[]) {
  const groups = new Map<string, string>();
  teacherRows.forEach((row) => {
    const teacher = normalizePersonName(row.教师姓名);
    const assistant = normalizePersonName(row.助理主管);
    if (teacher && assistant && teacher === assistant && row.教研组) {
      groups.set(assistant, assistantTeachingGroup(row.教研组));
    }
  });
  return groups;
}

function mergeRef(columnIndex: number, startRow: number, endRow: number) {
  const column = String.fromCharCode(65 + columnIndex);
  return `${column}${startRow}:${column}${endRow}`;
}

export function buildAssistantHierarchy(
  teacherRows: DataRow[],
  metricSpecs: readonly HierarchyMetricSpec[],
): HierarchyTable {
  const grouped = new Map<string, Map<string, Map<string, HierarchyBucket>>>();
  const ownGroups = assistantOwnTeachingGroups(teacherRows);
  teacherRows.forEach((row) => {
    const assistant = text(row.助理主管) || "未填写";
    const group = ownGroups.get(normalizePersonName(assistant)) || assistantTeachingGroup(row.教研组);
    const project = projectForResearchGroup(group);
    if (!grouped.has(project)) grouped.set(project, new Map());
    const groups = grouped.get(project)!;
    if (!groups.has(group)) groups.set(group, new Map());
    const assistants = groups.get(group)!;
    if (!assistants.has(assistant)) assistants.set(assistant, createBucket(metricSpecs));
    addSummaryRow(assistants.get(assistant)!, row, metricSpecs);
  });

  const rows: DataRow[] = [];
  const mergeCells: string[] = [];
  const totalBucket = createBucket(metricSpecs);
  const actualRow = (index: number) => index + 3;
  [...grouped.keys()].sort(projectCompare).forEach((project) => {
    const projectBucket = createBucket(metricSpecs);
    const groups = grouped.get(project)!;
    [...groups.keys()].sort(assistantTeachingGroupCompare).forEach((group) => {
      const start = rows.length;
      const groupBucket = createBucket(metricSpecs);
      const assistants = groups.get(group)!;
      [...assistants.keys()].sort(compareText).forEach((assistant) => {
        const item = assistants.get(assistant)!;
        mergeBucket(groupBucket, item);
        rows.push({
          教研组: rows.length === start ? group : "",
          "师训主管/助理主管": displayPersonName(assistant),
          ...metricValues(item, metricSpecs),
          __rowType: "detail",
        });
      });
      if (rows.length - start > 1) mergeCells.push(mergeRef(0, actualRow(start), actualRow(rows.length - 1)));
      const summaryRow = actualRow(rows.length);
      mergeCells.push(`A${summaryRow}:B${summaryRow}`);
      rows.push({
        教研组: group,
        "师训主管/助理主管": "",
        ...metricValues(groupBucket, metricSpecs),
        __rowType: "groupTotal",
      });
      mergeBucket(projectBucket, groupBucket);
    });
    const projectRow = actualRow(rows.length);
    mergeCells.push(`A${projectRow}:B${projectRow}`);
    rows.push({
      教研组: project,
      "师训主管/助理主管": "",
      ...metricValues(projectBucket, metricSpecs),
      __rowType: "projectTotal",
    });
    mergeBucket(totalBucket, projectBucket);
  });
  const totalRow = actualRow(rows.length);
  mergeCells.push(`A${totalRow}:B${totalRow}`);
  rows.push({
    教研组: "总计",
    "师训主管/助理主管": "",
    ...metricValues(totalBucket, metricSpecs),
    __rowType: "grandTotal",
  });
  return { rows, mergeCells };
}

export function buildResearchGroupHierarchy(
  teacherRows: DataRow[],
  metricSpecs: readonly HierarchyMetricSpec[],
  includeOwner = true,
): HierarchyTable {
  const projects = new Map<string, Map<string, HierarchyBucket>>();
  teacherRows.forEach((row) => {
    const group = assistantTeachingGroup(row.教研组);
    const project = projectForResearchGroup(group);
    if (!projects.has(project)) projects.set(project, new Map());
    const groups = projects.get(project)!;
    if (!groups.has(group)) groups.set(group, createBucket(metricSpecs));
    addSummaryRow(groups.get(group)!, row, metricSpecs);
  });

  const rows: DataRow[] = [];
  const mergeCells: string[] = [];
  const totalBucket = createBucket(metricSpecs);
  const actualRow = (index: number) => index + 3;
  [...projects.keys()].sort(projectCompare).forEach((project) => {
    const projectBucket = createBucket(metricSpecs);
    const groups = projects.get(project)!;
    [...groups.keys()].sort(assistantTeachingGroupCompare).forEach((group) => {
      const item = groups.get(group)!;
      mergeBucket(projectBucket, item);
      rows.push({
        教研组: group,
        ...(includeOwner ? { 负责人: displayPersonNames(item.assistants) } : {}),
        ...metricValues(item, metricSpecs),
        __rowType: "detail",
      });
    });
    const projectRow = actualRow(rows.length);
    if (includeOwner) mergeCells.push(`A${projectRow}:B${projectRow}`);
    rows.push({
      教研组: project,
      ...(includeOwner ? { 负责人: "" } : {}),
      ...metricValues(projectBucket, metricSpecs),
      __rowType: "projectTotal",
    });
    mergeBucket(totalBucket, projectBucket);
  });
  const totalRow = actualRow(rows.length);
  if (includeOwner) mergeCells.push(`A${totalRow}:B${totalRow}`);
  rows.push({
    教研组: "总计",
    ...(includeOwner ? { 负责人: "" } : {}),
    ...metricValues(totalBucket, metricSpecs),
    __rowType: "grandTotal",
  });
  return { rows, mergeCells };
}
