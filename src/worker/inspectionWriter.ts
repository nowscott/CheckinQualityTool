import { buildWorkbook } from "./excelWriter";
import { historyItems } from "./inspectionSampler";
import type { DataRow, SheetDefinition } from "./types";
import type { InspectionOutput, InspectionSelection } from "./inspectionTypes";

const TEMPLATE_COLUMNS = [
  "教师姓名", "学员姓名", "课次日期", "课次时间", "是否反馈", "是否具有个性化点评", "是否具有图文并茂", "是否具有行课计划",
  "拟扣罚", "拟扣罚积分\n（超过48h反馈扣20/条 ）", "拟扣罚积分\n（反馈内容不具有个性化点评扣20/条 ）",
  "拟扣罚积分\n（反馈内容不具有图文并茂扣10/条 ）", "拟扣罚积分\n（反馈内容不具有行课计划扣10/条 ）",
] as const;
const TRACE_COLUMNS = ["抽检序号", "入选原因", "业务周", "抽取尝试次数", "源表行号", "课次ID", "老师邮箱", "报告链接H5", "系统是否生成报告", "报告生成时间", "产品分组", "项目组", "校区", "科目"] as const;
const RISK_TRACE_COLUMNS = ["是否抽检", "抽检序号", "抽检原因", "风险判断", "源表行号", "课次ID", "老师邮箱", "报告链接H5", "系统是否生成报告", "报告生成时间", "产品分组", "项目组", "校区", "科目"] as const;

const MANUAL_VALIDATIONS = [
  { type: "list" as const, formula1: '"是,否（未加群）,否（未反馈）,否（超过周一24点反馈）,否（多条反馈合并）,否（超过48h反馈）"', ranges: [] as string[] },
  { type: "list" as const, formula1: '"是,否"', ranges: [] as string[] },
];

function manualFormula(_row: DataRow, column: string, rowIndex: number) {
  if (column === "拟扣罚") return `IF(E${rowIndex}="是","0","-50")`;
  if (column === TEMPLATE_COLUMNS[9]) return `IF(E${rowIndex}="是","0","-20")`;
  if (column === TEMPLATE_COLUMNS[10]) return `IF(F${rowIndex}="是","0","-20")`;
  if (column === TEMPLATE_COLUMNS[11]) return `IF(G${rowIndex}="是","0","-20")`;
  if (column === TEMPLATE_COLUMNS[12]) return `IF(H${rowIndex}="是","0","-20")`;
  return undefined;
}

function manualDate(value: string) {
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})/u);
  return match ? `${match[1]}/${Number(match[2])}/${Number(match[3])}` : value;
}

function manualTime(start: string, end: string) {
  const first = start.match(/\b(\d{2}:\d{2})/u)?.[1] || "";
  const last = end.match(/\b(\d{2}:\d{2})/u)?.[1] || "";
  return first && last ? `${first}-${last}` : `${first}${last}`;
}

function sourceColumns(selection: InspectionSelection) {
  if (selection.sourceColumns.length) return selection.sourceColumns;
  const first = selection.allEligibleRows[0]?.source || {};
  return Object.keys(first);
}

function selectedRows(selection: InspectionSelection) {
  return selection.selectedRows.map((row) => ({
    教师姓名: row.teacherName,
    学员姓名: row.studentName,
    课次日期: manualDate(row.lessonStart),
    课次时间: manualTime(row.lessonStart, row.lessonEnd),
    是否反馈: "",
    是否具有个性化点评: "",
    是否具有图文并茂: "",
    是否具有行课计划: "",
    拟扣罚: "",
    [TEMPLATE_COLUMNS[9]]: "",
    [TEMPLATE_COLUMNS[10]]: "",
    [TEMPLATE_COLUMNS[11]]: "",
    [TEMPLATE_COLUMNS[12]]: "",
    抽检序号: row.selectionOrder,
    入选原因: row.selectionReason,
    业务周: `${selection.businessWeekStart}~${selection.businessWeekEnd}`,
    抽取尝试次数: selection.attempt,
    源表行号: row.sourceRowNumber,
    课次ID: row.courseId,
    教师邮箱: row.teacherEmail,
    报告链接H5: row.source.报告链接H5 || "",
    系统是否生成报告: row.submittedValue,
    报告生成时间: row.source.报告生成时间 || "",
    产品分组: row.productGroup,
    项目组: row.projectGroup,
    校区: row.campus,
    科目: row.source.科目 || "",
  })) as DataRow[];
}

function riskRows(selection: InspectionSelection) {
  return selection.riskRows.map((row) => ({
    教师姓名: row.teacherName,
    学员姓名: row.studentName,
    课次日期: manualDate(row.lessonStart),
    课次时间: manualTime(row.lessonStart, row.lessonEnd),
    是否反馈: "",
    是否具有个性化点评: "",
    是否具有图文并茂: "",
    是否具有行课计划: "",
    拟扣罚: "",
    [TEMPLATE_COLUMNS[9]]: "",
    [TEMPLATE_COLUMNS[10]]: "",
    [TEMPLATE_COLUMNS[11]]: "",
    [TEMPLATE_COLUMNS[12]]: "",
    是否抽检: row.inspected ? "是" : "否",
    抽检序号: row.inspectionOrder,
    抽检原因: row.inspectionReason,
    风险判断: "是否生成报告=否",
    源表行号: row.sourceRowNumber,
    课次ID: row.courseId,
    教师邮箱: row.teacherEmail,
    报告链接H5: row.source.报告链接H5 || "",
    系统是否生成报告: row.submittedValue,
    报告生成时间: row.source.报告生成时间 || "",
    产品分组: row.productGroup,
    项目组: row.projectGroup,
    校区: row.campus,
    科目: row.source.科目 || "",
  })) as DataRow[];
}

export function buildInspectionOutput(selection: InspectionSelection, includeExplanation: boolean): InspectionOutput {
  const selectedColumns = [...TEMPLATE_COLUMNS, ...TRACE_COLUMNS];
  const riskColumns = [...TEMPLATE_COLUMNS, ...RISK_TRACE_COLUMNS];
  const validationEnd = 1 + selection.selectedRows.length;
  const riskValidationEnd = 1 + selection.riskRows.length;
  const selectedValidations = [
    { ...MANUAL_VALIDATIONS[0], ranges: [`E2:E${validationEnd}`] },
    { ...MANUAL_VALIDATIONS[1], ranges: [`F2:H${validationEnd}`] },
  ];
  const riskValidations = [
    { ...MANUAL_VALIDATIONS[0], ranges: [`E2:E${riskValidationEnd}`] },
    { ...MANUAL_VALIDATIONS[1], ranges: [`F2:H${riskValidationEnd}`] },
  ];
  const sheets: SheetDefinition[] = [
    {
      name: "抽检名单",
      rows: selectedRows(selection),
      columns: selectedColumns,
      headerStyle: 31,
      headerHeight: 56,
      freezeRows: 1,
      dataRowHeight: 21,
      formula: manualFormula,
      dataValidations: selectedValidations,
      widths: {
        教师姓名: 32, 学员姓名: 20, 课次日期: 18, 课次时间: 20, 是否反馈: 22,
        是否具有个性化点评: 26, 是否具有图文并茂: 24, 是否具有行课计划: 24, 拟扣罚: 16,
        [TEMPLATE_COLUMNS[9]]: 36, [TEMPLATE_COLUMNS[10]]: 50, [TEMPLATE_COLUMNS[11]]: 50, [TEMPLATE_COLUMNS[12]]: 50,
        抽检序号: 10, 入选原因: 32, 业务周: 24, 抽取尝试次数: 12, 源表行号: 10,
        课次ID: 20, 老师邮箱: 30, 报告链接H5: 34, 系统是否生成报告: 18, 报告生成时间: 22,
        产品分组: 30, 项目组: 28, 校区: 24, 科目: 12,
      },
      rowStyle: () => 32,
    },
    {
      name: "未生成报告风险",
      rows: riskRows(selection),
      columns: riskColumns,
      headerStyle: 31,
      headerHeight: 56,
      freezeRows: 1,
      dataRowHeight: 21,
      formula: manualFormula,
      dataValidations: riskValidations,
      widths: {
        教师姓名: 32, 学员姓名: 20, 课次日期: 18, 课次时间: 20, 是否反馈: 22,
        是否具有个性化点评: 26, 是否具有图文并茂: 24, 是否具有行课计划: 24, 拟扣罚: 16,
        [TEMPLATE_COLUMNS[9]]: 36, [TEMPLATE_COLUMNS[10]]: 50, [TEMPLATE_COLUMNS[11]]: 50, [TEMPLATE_COLUMNS[12]]: 50,
        是否抽检: 12, 抽检序号: 10, 抽检原因: 34, 风险判断: 24, 源表行号: 10,
        课次ID: 20, 老师邮箱: 30, 报告链接H5: 34, 系统是否生成报告: 18, 报告生成时间: 22,
        产品分组: 30, 项目组: 28, 校区: 24, 科目: 12,
      },
      rowStyle: () => 32,
    },
  ];

  if (includeExplanation) {
    const explanation: DataRow[] = [
      { 项目: "生成时间", 值: new Date().toLocaleString("zh-CN", { hour12: false }) },
      { 项目: "课程反馈文件", 值: selection.sourceName },
      { 项目: "在职明细文件", 值: selection.roster.sourceName },
      { 项目: "在职快照日期", 值: selection.roster.snapshotDate || "文件名未识别日期" },
      { 项目: "业务周", 值: `${selection.businessWeekStart}~${selection.businessWeekEnd}` },
      { 项目: "抽取尝试次数", 值: selection.attempt },
      { 项目: "抽检规则版本", 值: selection.ruleVersion },
      { 项目: "抽检数上限", 值: selection.sampleLimit },
      { 项目: "原始课程条数", 值: selection.stats.sourceRows },
      { 项目: "在职教师课程条数", 值: selection.stats.eligibleRows },
      { 项目: "在职教师人数", 值: selection.stats.eligibleTeachers },
      { 项目: "实际抽检课程条数", 值: selection.stats.selectedRows },
      { 项目: "实际覆盖教师人数", 值: selection.stats.selectedTeachers },
      { 项目: "未生成报告课程数", 值: selection.stats.unsubmittedRows },
      { 项目: "未生成报告且被抽检数", 值: selection.stats.unsubmittedSelectedRows },
      { 项目: "排除课程数", 值: selection.stats.excludedRows },
      { 项目: "排除原因", 值: `无邮箱 ${selection.stats.excludedNoEmailRows} 条；邮箱不在职 ${selection.stats.excludedNotInRosterRows} 条` },
      { 项目: "提交字段未知值", 值: selection.stats.unknownSubmissionRows },
      { 项目: "数据保存范围", 值: "数据库只保存抽检审计记录，不保存原始 Excel、课堂反馈正文或报告链接" },
    ];
    sheets.push({
      name: "处理说明",
      rows: explanation,
      columns: ["项目", "值"],
      title: "抽检处理说明",
      titleStyle: 5,
      headerStyle: 1,
      titleHeight: 34,
      headerHeight: 30,
      freezeRows: 2,
      dataRowHeight: 24,
      widths: { 项目: 30, 值: 90 },
    });
  }

  const chunks = buildWorkbook(sheets, 6);
  return {
    chunks,
    filename: `课堂反馈抽检_${selection.businessWeekStart.replaceAll("-", "")}_第${selection.attempt}次.xlsx`,
    summary: selection.stats,
    historyPayload: {
      batch: {
        businessWeekStart: selection.businessWeekStart,
        businessWeekEnd: selection.businessWeekEnd,
        sourceName: selection.sourceName,
        sourceSha256: selection.sourceSha256,
        rosterName: selection.roster.sourceName,
        rosterSha256: selection.rosterSha256,
        rosterSnapshotDate: selection.roster.snapshotDate,
        sampleLimit: selection.sampleLimit,
        eligibleCount: selection.stats.eligibleRows,
        selectedCount: selection.stats.selectedRows,
        teacherCount: selection.stats.selectedTeachers,
        unsubmittedCount: selection.stats.unsubmittedRows,
        ruleVersion: selection.ruleVersion,
        attempt: selection.attempt,
        batchKind: selection.batchKind,
      },
      items: historyItems(selection.selectedRows),
    },
  };
}
