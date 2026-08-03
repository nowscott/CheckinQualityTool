import { buildWorkbook } from "./excelWriter";
import type { ChatInfo, DataRow, SheetDefinition } from "./types";
import type { StageReportListInfo } from "./stageReportListParser";
import type { StageReportMatchInfo } from "./stageReportMatching";

const APPENDED_COLUMNS = [
  "本次检查结论", "命中关键词", "命中聊天内容", "命中群名", "命中时间", "来源文件", "源聊天行号", "原分母行号", "去重合并行号",
] as const;

const TEACHER_COLUMNS = ["教师姓名", "教师邮箱", "应发送数", "已发送数", "未发送数", "发送率", "申诉数"] as const;

export function buildStageReportOutput(
  listInfo: StageReportListInfo,
  chatInfo: ChatInfo,
  matchInfo: StageReportMatchInfo,
  sourceNames: { list: string; chats: string },
) {
  const explanation: DataRow[] = [
    { 项目: "生成时间", 值: new Date().toLocaleString("zh-CN", { hour12: false }) },
    { 项目: "阶段性报告分母文件", 值: sourceNames.list },
    { 项目: "聊天质检结果文件", 值: sourceNames.chats },
    { 项目: "分母工作表", 值: listInfo.sheetName },
    { 项目: "聊天工作表", 值: chatInfo.sheetName },
    { 项目: "分母识别字段", 值: "教师姓名、邮箱、学号、学员姓名（支持常见别名）" },
    { 项目: "分母去重规则", 值: "同一工作表中整行完全相同的记录只保留第一行，并记录合并行号" },
    { 项目: "匹配规则", 值: "聊天文件已按“阶段性报告”关键词检索导出，不再重复验证正文关键词；分母有邮箱时按邮箱匹配，无邮箱时按教师姓名去空格、去末尾数字和“老师”后缀兜底；学员按打卡口径优先匹配清洗后姓名的后两字（不足两字自动匹配末字），群名或聊天内容均可作为学员证据" },
    { 项目: "聊天清洗规则", 值: "删除私聊、无有效邮箱、发送方非员工、引用回复" },
    { 项目: "申诉规则", 值: "原“是否申诉”和“申诉情况详情”列不覆盖；教师汇总单列统计申诉数，不从分母剔除" },
  ];
  Object.entries(listInfo.counts).forEach(([key, value]) => explanation.push({ 项目: `分母_${key}`, 值: value }));
  Object.entries(chatInfo.counts).forEach(([key, value]) => explanation.push({ 项目: `聊天_${key}`, 值: value }));
  Object.entries(matchInfo.counts).forEach(([key, value]) => explanation.push({ 项目: `检查_${key}`, 值: value }));

  const sheets: SheetDefinition[] = [
    {
      name: "阶段性报告检查明细",
      rows: matchInfo.detailRows,
      columns: [...listInfo.columns, ...APPENDED_COLUMNS],
      widths: { 邮箱: 28, 教师邮箱: 28, 命中聊天内容: 80, 命中群名: 38, 来源文件: 34, 申诉情况详情: 48 },
      rowStyle: (row) => row.本次检查结论 === "已发送" ? 2 : row.本次检查结论 === "字段缺失" ? 4 : 3,
    },
    {
      name: "教师发送汇总",
      rows: matchInfo.teacherRows,
      columns: TEACHER_COLUMNS,
      widths: { 教师姓名: 18, 教师邮箱: 28 },
      rowStyle: (row) => Number(row.发送率) >= 1 ? 2 : 3,
      cellStyle: (row, column, style) => column === "发送率" ? (style === 2 ? 12 : 12) : style,
      dataBarColumns: ["发送率"],
    },
    {
      name: "处理说明",
      rows: explanation,
      columns: ["项目", "值"],
      widths: { 项目: 32, 值: 100 },
    },
  ];
  return buildWorkbook(sheets);
}
