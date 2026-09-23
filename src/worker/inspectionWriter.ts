import { buildWorkbook } from "./excelWriter";
import type { DataRow, SheetDefinition } from "./types";
import type { InspectionOutput, InspectionSelection } from "./inspectionTypes";

const INSPECTION_COLUMNS = [
  "抽检序号", "抽检类型", "入选原因", "业务周", "抽取尝试次数", "源表行号", "课次ID", "教师邮箱", "报告链接H5",
  "系统是否生成报告", "报告生成时间", "产品分组", "项目组", "校区", "科目",
] as const;

function selectedRows(selection: InspectionSelection) {
  return selection.selectedRows.map((row) => ({
    抽检序号: row.selectionOrder,
    抽检类型: row.unsubmitted ? "容量外加抽" : "普通抽检",
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

export function buildInspectionOutput(selection: InspectionSelection, includeExplanation: boolean): InspectionOutput {
  const sheets: SheetDefinition[] = [{
    name: "抽检名单",
    rows: selectedRows(selection),
    columns: INSPECTION_COLUMNS,
    headerStyle: 31,
    headerHeight: 42,
    freezeRows: 1,
    dataRowHeight: 22,
    widths: {
      抽检序号: 10,
      抽检类型: 16,
      入选原因: 28,
      业务周: 24,
      抽取尝试次数: 12,
      源表行号: 10,
      课次ID: 20,
      教师邮箱: 30,
      报告链接H5: 38,
      系统是否生成报告: 18,
      报告生成时间: 22,
      产品分组: 30,
      项目组: 28,
      校区: 24,
      科目: 12,
    },
  }];

  if (includeExplanation) {
    const explanation: DataRow[] = [
      { 项目: "生成时间", 值: new Date().toLocaleString("zh-CN", { hour12: false }) },
      { 项目: "课程反馈文件", 值: selection.sourceName },
      { 项目: "在职明细文件", 值: selection.roster.sourceName },
      { 项目: "在职快照日期", 值: selection.roster.snapshotDate || "文件名未识别日期" },
      { 项目: "业务周", 值: `${selection.businessWeekStart}~${selection.businessWeekEnd}` },
      { 项目: "导出优先级", 值: selection.priority.mode === "unreported" ? "优先本月未反馈教师，再覆盖其他教师" : "优先覆盖不同教师，剩余名额用于未反馈教师加频" },
      { 项目: "重点关注表未反馈教师", 值: selection.priority.focusTeacherCount },
      { 项目: "匹配到的未反馈教师", 值: selection.priority.matchedFocusTeacherCount },
      { 项目: "未匹配的未反馈教师", 值: selection.priority.unmatchedFocusTeacherCount },
      { 项目: "同名冲突教师", 值: selection.priority.ambiguousFocusTeacherCount },
      { 项目: "评分来源教师数", 值: selection.priority.scoreSourceTeacherCount },
      { 项目: "已匹配评分教师数", 值: selection.priority.matchedScoreTeacherCount },
      { 项目: "缺少评分教师数", 值: selection.priority.missingScoreTeacherCount },
      { 项目: "评分身份歧义教师数", 值: selection.priority.ambiguousScoreTeacherCount },
      { 项目: "抽检规则版本", 值: selection.ruleVersion },
      { 项目: "普通抽检条数上限", 值: selection.sampleLimit },
      { 项目: "原始课程条数", 值: selection.stats.sourceRows },
      { 项目: "在职教师课程条数", 值: selection.stats.eligibleRows },
      { 项目: "在职教师人数", 值: selection.stats.eligibleTeachers },
      { 项目: "经理岗位排除课程数", 值: selection.stats.excludedManagementRows },
      { 项目: "经理岗位排除教师数", 值: selection.stats.excludedManagementTeachers },
      { 项目: "实际普通抽检课程条数", 值: selection.stats.normalSelectedRows },
      { 项目: "容量外加抽课程条数", 值: selection.stats.extraSelectedRows },
      { 项目: "实际抽检课程总条数", 值: selection.stats.selectedRows },
      { 项目: "实际覆盖教师人数", 值: selection.stats.selectedTeachers },
      { 项目: "未反馈教师加频课程数", 值: selection.stats.focusTeacherExtraRows },
      { 项目: "未生成报告课程数", 值: selection.stats.unsubmittedRows },
      { 项目: "未生成报告且被抽检数", 值: selection.stats.unsubmittedSelectedRows },
      { 项目: "排除课程数", 值: selection.stats.excludedRows },
      { 项目: "排除原因", 值: `无邮箱 ${selection.stats.excludedNoEmailRows} 条；邮箱不在职 ${selection.stats.excludedNotInRosterRows} 条；岗位含“经理” ${selection.stats.excludedManagementRows} 条，不参与抽检；主管岗位参与抽检` },
      { 项目: "提交字段未知值", 值: selection.stats.unknownSubmissionRows },
      { 项目: "数据保存范围", 值: "上传的 Excel 只在当前浏览器处理；应用不保存课程明细或抽检历史" },
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
    filename: `课堂反馈抽检_${selection.businessWeekStart.replaceAll("-", "")}_${selection.priority.mode === "unreported" ? "未反馈优先" : "教师覆盖优先"}.xlsx`,
    summary: selection.stats,
    priority: selection.priority,
  };
}
