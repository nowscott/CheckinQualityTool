import { buildWorkbook } from "./excelWriter";
import { REMINDER_PASS_RATE } from "./reminderConfig";
import { reminderProjectGroup } from "./reminderProjectGroup";
import type { ChatInfo, DataRow, SheetDefinition, SourceNames } from "./types";
import type { ReminderListInfo } from "./reminderListParser";
import type { ReminderMatchInfo } from "./reminderMatching";
import { text } from "./utils";

const STUDENT_COLUMNS = [
  "教师姓名", "教研组", "师训组长", "助理主管", "学员姓名",
  "年级", "学管姓名", "是否发送", "申诉情况说明",
] as const;

const TEACHER_COLUMNS = [
  "教师姓名", "教研组", "师训组长", "助理主管",
  "应发送人次", "已发送人次", "发送差额", "发送进度", "是否达标",
] as const;

const TRAINING_COLUMNS = [
  "教研组", "应发送数", "已发送数", "发送率", "是否达标（80%）",
] as const;

const ASSISTANT_COLUMNS = [
  "项目组", "教研组", "助理主管", "应发送数", "已发送数", "发送率", "是否达标（80%）",
] as const;

const PROJECT_COLUMNS = [
  "教研组", "负责人", "应发送数", "已发送数", "发送率", "是否达标（80%）",
] as const;

const CAMPUS_COLUMNS = [
  "校区", "区域教学负责人", "应发送数", "已发送数", "发送率", "是否达标（80%）",
] as const;

const EXCEPTION_COLUMNS = [
  "异常类型", "异常原因", "质检序号", "授课教师", "学员姓名", "匹配学员姓名",
  "教研组", "师训组长", "助理主管", "源名单行号", "保留名单行号",
  "命中群名", "命中聊天内容", "命中聊天时间", "命中质检文件", "源聊天行号",
] as const;

interface PublicTable {
  rows: DataRow[];
  mergeCells: string[];
}

interface SummaryBucket {
  total: number;
  sent: number;
  counselorSent: number;
  assistants: Set<string>;
}

function createBucket(): SummaryBucket {
  return { total: 0, sent: 0, counselorSent: 0, assistants: new Set() };
}

function addStudent(bucket: SummaryBucket, row: DataRow) {
  if (row.是否发送 === "已申诉通过" || row.匹配状态 === "已申诉通过") return;
  bucket.total += 1;
  if (row.是否发送 === "是") bucket.sent += 1;
  if (row.学管是否发送 === "是") bucket.counselorSent += 1;
  const assistant = text(row.助理主管);
  if (assistant) bucket.assistants.add(assistant);
}

function rate(sent: number, total: number) {
  return total ? sent / total : 0;
}

function rateText(sent: number, total: number) {
  return `${(rate(sent, total) * 100).toFixed(1)}%`;
}

function passText(sent: number, total: number) {
  return rate(sent, total) >= REMINDER_PASS_RATE ? "是" : "否";
}

function compareText(a: unknown, b: unknown) {
  return text(a).localeCompare(text(b), "zh-CN");
}

function displayTeachingGroup(value: unknown) {
  const group = text(value);
  if (!group) return "未分组";
  return group.endsWith("组") || group.endsWith("项目") ? group : `${group}组`;
}

function studentOutputRows(rows: DataRow[]) {
  return rows.map((row) => ({
    教师姓名: row.教师姓名 || "",
    教研组: row.教研组 || "",
    师训组长: row.师训组长 || "",
    助理主管: row.助理主管 || "",
    学员姓名: row.学员姓名 || "",
    年级: row.年级 || "",
    学管姓名: row.学管姓名 || "",
    是否发送: row.是否发送 || "",
    申诉情况说明: row.申诉情况说明 || "/",
  }));
}

function teacherOutputRows(rows: DataRow[]) {
  return rows.map((row) => ({
    教师姓名: row.教师姓名 || "",
    教研组: row.教研组 || "",
    师训组长: row.师训组长 || "",
    助理主管: row.助理主管 || "",
    应发送人次: row.应发送数 || 0,
    已发送人次: row.已发送数 || 0,
    发送差额: row.发送差额 || 0,
    发送进度: row.发送率 || "0.0%",
    是否达标: row.是否达标 || "否",
  }));
}

function mergeRef(columnIndex: number, startRow: number, endRow: number) {
  const column = String.fromCharCode(65 + columnIndex);
  return `${column}${startRow}:${column}${endRow}`;
}

function pushTotalRow(rows: DataRow[], labelColumn: string, label: string, bucket: SummaryBucket, rowType: string) {
  rows.push({
    [labelColumn]: label,
    应发送数: bucket.total,
    已发送数: bucket.sent,
    发送率: rateText(bucket.sent, bucket.total),
    "是否达标（80%）": passText(bucket.sent, bucket.total),
    __rowType: rowType,
  });
}

function passFailStyle(passed: boolean, includeResultColors: boolean) {
  if (!includeResultColors) return 8;
  return passed ? 2 : 3;
}

function summaryAwareStyle(row: DataRow, includeResultColors: boolean) {
  if (row.__rowType === "grandTotal") return 10;
  if (row.__rowType === "projectTotal") return 10;
  if (row.__rowType === "groupTotal") return 9;
  return passFailStyle(row["是否达标（80%）"] === "是", includeResultColors);
}

function buildPublicTable(matchInfo: ReminderMatchInfo): PublicTable {
  const projectOrder = ["博文项目", "双语项目", "益智项目", "文理综项目", "其他项目"];
  const grouped = new Map<string, Map<string, Map<string, SummaryBucket>>>();
  for (const row of matchInfo.studentRows) {
    const teachingGroup = String(row.教研组 || "未分组");
    const project = reminderProjectGroup(teachingGroup);
    const assistant = String(row.助理主管 || "未填写");
    if (!grouped.has(project)) grouped.set(project, new Map());
    const teachingGroups = grouped.get(project)!;
    if (!teachingGroups.has(teachingGroup)) teachingGroups.set(teachingGroup, new Map());
    const assistants = teachingGroups.get(teachingGroup)!;
    if (!assistants.has(assistant)) assistants.set(assistant, createBucket());
    addStudent(assistants.get(assistant)!, row);
  }

  const actualRow = (dataIndex: number) => dataIndex + 3;
  const rows: DataRow[] = [];
  const mergeCells: string[] = [];
  const sortedProjects = [...grouped.keys()].sort(
    (a, b) =>
      (projectOrder.indexOf(a) === -1 ? 99 : projectOrder.indexOf(a)) -
        (projectOrder.indexOf(b) === -1 ? 99 : projectOrder.indexOf(b)) ||
      a.localeCompare(b, "zh-CN"),
  );

  for (const project of sortedProjects) {
    const projectStart = rows.length;
    const teachingGroups = grouped.get(project)!;
    const projectBucket = createBucket();
    const sortedTeachingGroups = [...teachingGroups.keys()].sort((a, b) => a.localeCompare(b, "zh-CN"));
    for (const teachingGroup of sortedTeachingGroups) {
      const groupStart = rows.length;
      const assistants = teachingGroups.get(teachingGroup)!;
      const groupBucket = createBucket();
      const sortedAssistants = [...assistants.keys()].sort((a, b) => a.localeCompare(b, "zh-CN"));
      sortedAssistants.forEach((assistant) => {
        const item = assistants.get(assistant)!;
        if (!item.total) return;
        groupBucket.total += item.total;
        groupBucket.sent += item.sent;
        groupBucket.counselorSent += item.counselorSent;
        rows.push({
          项目组: rows.length === projectStart ? project : "",
          教研组: rows.length === groupStart ? teachingGroup : "",
          助理主管: assistant,
          应发送数: item.total,
          已发送数: item.sent,
          发送率: rateText(item.sent, item.total),
          "是否达标（80%）": passText(item.sent, item.total),
          __rowType: "detail",
        });
      });
      if (!groupBucket.total) continue;
      if (rows.length - groupStart > 1) mergeCells.push(mergeRef(1, actualRow(groupStart), actualRow(rows.length - 1)));
      rows.push({
        项目组: "",
        教研组: "",
        助理主管: `${displayTeachingGroup(teachingGroup)}小计`,
        应发送数: groupBucket.total,
        已发送数: groupBucket.sent,
        发送率: rateText(groupBucket.sent, groupBucket.total),
        "是否达标（80%）": passText(groupBucket.sent, groupBucket.total),
        __rowType: "groupTotal",
      });
      projectBucket.total += groupBucket.total;
      projectBucket.sent += groupBucket.sent;
      projectBucket.counselorSent += groupBucket.counselorSent;
    }
    if (!projectBucket.total) continue;
    rows.push({
      项目组: "",
      教研组: "",
      助理主管: `${project}合计`,
      应发送数: projectBucket.total,
      已发送数: projectBucket.sent,
      发送率: rateText(projectBucket.sent, projectBucket.total),
      "是否达标（80%）": passText(projectBucket.sent, projectBucket.total),
      __rowType: "projectTotal",
    });
    if (rows.length - projectStart > 1) mergeCells.push(mergeRef(0, actualRow(projectStart), actualRow(rows.length - 1)));
  }

  const grandBucket = createBucket();
  rows.forEach((row) => {
    if (row.__rowType !== "projectTotal") return;
    grandBucket.total += Number(row.应发送数) || 0;
    grandBucket.sent += Number(row.已发送数) || 0;
  });
  rows.push({
    项目组: "总计",
    教研组: "",
    助理主管: "总计",
    应发送数: grandBucket.total,
    已发送数: grandBucket.sent,
    发送率: rateText(grandBucket.sent, grandBucket.total),
    "是否达标（80%）": passText(grandBucket.sent, grandBucket.total),
    __rowType: "grandTotal",
  });

  return { rows, mergeCells };
}

function buildTrainingRows(studentRows: DataRow[]) {
  const projectOrder = ["博文项目", "双语项目", "益智项目", "文理综项目", "其他项目"];
  const projects = new Map<string, Map<string, Map<string, SummaryBucket>>>();
  for (const row of studentRows) {
    const project = reminderProjectGroup(row.教研组);
    const group = text(row.教研组) || "未分组";
    const trainingLead = text(row.师训组长) || "未填写";
    if (!projects.has(project)) projects.set(project, new Map());
    const groups = projects.get(project)!;
    if (!groups.has(group)) groups.set(group, new Map());
    const leads = groups.get(group)!;
    if (!leads.has(trainingLead)) leads.set(trainingLead, createBucket());
    addStudent(leads.get(trainingLead)!, row);
  }
  const rows: DataRow[] = [];
  [...projects.keys()]
    .sort((a, b) =>
      (projectOrder.indexOf(a) === -1 ? 99 : projectOrder.indexOf(a)) -
        (projectOrder.indexOf(b) === -1 ? 99 : projectOrder.indexOf(b)) ||
      compareText(a, b),
    )
    .forEach((project) => {
      const projectBucket = createBucket();
      const groups = projects.get(project)!;
      const projectRows: DataRow[] = [];
      [...groups.keys()].sort(compareText).forEach((group) => {
        const groupBucket = createBucket();
        const detailRows: DataRow[] = [];
        const leads = groups.get(group)!;
        [...leads.keys()].sort(compareText).forEach((lead) => {
          const item = leads.get(lead)!;
          if (!item.total) return;
          groupBucket.total += item.total;
          groupBucket.sent += item.sent;
          groupBucket.counselorSent += item.counselorSent;
          detailRows.push({
            教研组: lead,
            应发送数: item.total,
            已发送数: item.sent,
            发送率: rateText(item.sent, item.total),
            "是否达标（80%）": passText(item.sent, item.total),
            __rowType: "detail",
          });
        });
        if (!groupBucket.total) return;
        projectBucket.total += groupBucket.total;
        projectBucket.sent += groupBucket.sent;
        projectBucket.counselorSent += groupBucket.counselorSent;
        projectRows.push({
          教研组: displayTeachingGroup(group),
          应发送数: groupBucket.total,
          已发送数: groupBucket.sent,
          发送率: rateText(groupBucket.sent, groupBucket.total),
          "是否达标（80%）": passText(groupBucket.sent, groupBucket.total),
          __rowType: "groupTotal",
        });
        projectRows.push(...detailRows);
      });
      rows.push({
        教研组: project,
        应发送数: projectBucket.total,
        已发送数: projectBucket.sent,
        发送率: rateText(projectBucket.sent, projectBucket.total),
        "是否达标（80%）": passText(projectBucket.sent, projectBucket.total),
        __rowType: "projectTotal",
      });
      rows.push(...projectRows);
    });
  const grandBucket = createBucket();
  rows.forEach((row) => {
    if (row.__rowType !== "projectTotal") return;
    grandBucket.total += Number(row.应发送数) || 0;
    grandBucket.sent += Number(row.已发送数) || 0;
  });
  pushTotalRow(rows, "教研组", "总计", grandBucket, "grandTotal");
  return rows;
}

function buildProjectRows(studentRows: DataRow[]) {
  const projectOrder = ["博文项目", "双语项目", "益智项目", "文理综项目", "其他项目"];
  const projects = new Map<string, Map<string, SummaryBucket>>();
  for (const row of studentRows) {
    const project = reminderProjectGroup(row.教研组);
    const group = text(row.教研组) || "未分组";
    if (!projects.has(project)) projects.set(project, new Map());
    const groups = projects.get(project)!;
    if (!groups.has(group)) groups.set(group, createBucket());
    addStudent(groups.get(group)!, row);
  }
  const rows: DataRow[] = [];
  [...projects.keys()]
    .sort((a, b) =>
      (projectOrder.indexOf(a) === -1 ? 99 : projectOrder.indexOf(a)) -
        (projectOrder.indexOf(b) === -1 ? 99 : projectOrder.indexOf(b)) ||
      compareText(a, b),
    )
    .forEach((project) => {
      const groups = projects.get(project)!;
      const projectBucket = createBucket();
      [...groups.keys()].sort(compareText).forEach((group) => {
        const item = groups.get(group)!;
        if (!item.total) return;
        projectBucket.total += item.total;
        projectBucket.sent += item.sent;
        projectBucket.counselorSent += item.counselorSent;
        item.assistants.forEach((assistant) => projectBucket.assistants.add(assistant));
        rows.push({
          教研组: displayTeachingGroup(group),
          负责人: [...item.assistants].sort(compareText).join("、") || "/",
          应发送数: item.total,
          已发送数: item.sent,
          发送率: rateText(item.sent, item.total),
          "是否达标（80%）": passText(item.sent, item.total),
          __rowType: "detail",
        });
      });
      if (!projectBucket.total) return;
      rows.push({
        教研组: project,
        负责人: [...projectBucket.assistants].sort(compareText).join("、") || "/",
        应发送数: projectBucket.total,
        已发送数: projectBucket.sent,
        发送率: rateText(projectBucket.sent, projectBucket.total),
        "是否达标（80%）": passText(projectBucket.sent, projectBucket.total),
        __rowType: "projectTotal",
      });
    });
  const grandBucket = createBucket();
  rows.forEach((row) => {
    if (row.__rowType !== "projectTotal") return;
    grandBucket.total += Number(row.应发送数) || 0;
    grandBucket.sent += Number(row.已发送数) || 0;
    text(row.负责人).split("、").filter((item) => item && item !== "/").forEach((item) => grandBucket.assistants.add(item));
  });
  rows.push({
    教研组: "总计",
    负责人: [...grandBucket.assistants].sort(compareText).join("、") || "/",
    应发送数: grandBucket.total,
    已发送数: grandBucket.sent,
    发送率: rateText(grandBucket.sent, grandBucket.total),
    "是否达标（80%）": passText(grandBucket.sent, grandBucket.total),
    __rowType: "grandTotal",
  });
  return rows;
}

function buildCampusRows(studentRows: DataRow[]) {
  const campus = new Map<string, SummaryBucket>();
  for (const row of studentRows) {
    const key = text(row.校区) || "未填写";
    if (!campus.has(key)) campus.set(key, createBucket());
    addStudent(campus.get(key)!, row);
  }
  return [...campus.entries()]
    .filter(([, item]) => item.total)
    .sort((a, b) => compareText(a[0], b[0]))
    .map(([name, item]) => ({
      校区: name,
      区域教学负责人: "/",
      应发送数: item.total,
      已发送数: item.sent,
      发送率: rateText(item.sent, item.total),
      "是否达标（80%）": passText(item.sent, item.total),
    }));
}

function explanationRows(
  listInfo: ReminderListInfo,
  chatInfo: ChatInfo,
  matchInfo: ReminderMatchInfo,
  sourceNames: SourceNames,
  includeCleanChats: boolean,
) {
  const rows: DataRow[] = [
    { 项目: "生成时间", 值: new Date().toLocaleString("zh-CN", { hour12: false }) },
    { 项目: "开课提醒学员明细名单", 值: sourceNames.list },
    { 项目: "聊天质检结果", 值: sourceNames.chat },
    { 项目: "分母工作表", 值: listInfo.sheetName },
    { 项目: "聊天工作表", 值: chatInfo.sheetName },
    { 项目: "是否输出清洗后聊天", 值: includeCleanChats ? "是" : "否" },
    { 项目: "达标阈值", 值: `${REMINDER_PASS_RATE * 100}%` },
    { 项目: "分母去重规则", 值: "仅当分母表格一整行所有字段完全一致时去重；同一授课教师+学员姓名不会被自动合并" },
    { 项目: "匹配优先级1", 值: "聊天发送人与授课教师一致，且群聊名称或聊天内容包含学员姓名，自动判定该老师-学员记录已发送；发送人姓名会忽略末尾数字" },
    { 项目: "匹配优先级2", 值: "群聊名称或聊天内容同时包含学员姓名和授课教师/教师姓名，自动判定已发送" },
    { 项目: "匹配优先级3", 值: "群聊名称包含学员姓名，且该学员在去重后分母中唯一，自动判定已发送" },
    { 项目: "匹配优先级4", 值: "聊天内容包含学员姓名，且该学员在去重后分母中唯一，自动判定已发送" },
    { 项目: "异常处理", 值: "分母整行重复直接剔除，只在处理说明记录数量；多条质检命中同一分母时正常判定已发送；同名学员需通过授课教师+学员姓名定位，无法定位时保持未发送且不写入异常；字段缺失写入“匹配核对-异常明细”" },
    { 项目: "申诉剔除规则", 值: "上传申诉文件且“申诉是否通过”为“是/通过”时，按教师姓名+学生姓名匹配分母；命中行保留在学员名单并标注“已申诉通过”，但不计入应发送、未发送和发送率分母" },
  ];
  Object.entries(listInfo.counts).forEach(([key, value]) => rows.push({ 项目: `分母_${key}`, 值: value }));
  Object.entries(chatInfo.counts).forEach(([key, value]) => rows.push({ 项目: `聊天_${key}`, 值: value }));
  Object.entries(matchInfo.counts).forEach(([key, value]) => rows.push({ 项目: `匹配_${key}`, 值: value }));
  return rows;
}

export function buildReminderOutput(
  listInfo: ReminderListInfo,
  chatInfo: ChatInfo,
  matchInfo: ReminderMatchInfo,
  sourceNames: SourceNames,
  includeCleanChats: boolean,
  includeResultColors = false,
) {
  const publicTable = buildPublicTable(matchInfo);
  const sheets: SheetDefinition[] = [
    {
      name: "学员名单",
      title: "开课提醒发送明细",
      rows: studentOutputRows(matchInfo.studentRows),
      columns: STUDENT_COLUMNS,
      widths: { 教研组: 18, 师训组长: 18, 助理主管: 18, 申诉情况说明: 32 },
      rowStyle: (row) => row.是否发送 === "已申诉通过"
        ? 8
        : passFailStyle(row.是否发送 === "是", includeResultColors),
    },
    {
      name: "教师维度发送进度",
      title: "开课提醒发送进度（教师维度）",
      rows: teacherOutputRows(matchInfo.teacherRows),
      columns: TEACHER_COLUMNS,
      widths: { 教研组: 18, 师训组长: 18, 助理主管: 18 },
      rowStyle: (row) => passFailStyle(row.是否达标 === "是", includeResultColors),
    },
    {
      name: "师训组维度",
      title: "开课提醒发送进度（师训组维度）",
      rows: buildTrainingRows(matchInfo.studentRows),
      columns: TRAINING_COLUMNS,
      widths: { 教研组: 26 },
      rowStyle: (row) => summaryAwareStyle(row, includeResultColors),
    },
    {
      name: "助理主管维度",
      title: "开课提醒话术发送进度（助理主管维度）",
      rows: publicTable.rows,
      columns: ASSISTANT_COLUMNS,
      widths: { 项目组: 18, 教研组: 18, 助理主管: 22 },
      mergeCells: publicTable.mergeCells,
      rowStyle: (row) => summaryAwareStyle(row, includeResultColors),
      cellStyle: (row, column, baseStyle) => {
        if (column === "项目组" || column === "教研组") return baseStyle || 8;
        return baseStyle;
      },
    },
    {
      name: "项目组维度",
      title: "开课提醒发送（教研组维度）",
      rows: buildProjectRows(matchInfo.studentRows),
      columns: PROJECT_COLUMNS,
      widths: { 教研组: 24, 负责人: 42 },
      rowStyle: (row) => summaryAwareStyle(row, includeResultColors),
    },
    {
      name: "校区维度",
      title: "开课提醒发送进度（校区维度）",
      rows: buildCampusRows(matchInfo.studentRows),
      columns: CAMPUS_COLUMNS,
      widths: { 校区: 30, 区域教学负责人: 18 },
      rowStyle: (row) => passFailStyle(row["是否达标（80%）"] === "是", includeResultColors),
    },
    {
      name: "匹配核对-异常明细",
      rows: matchInfo.exceptionRows,
      columns: EXCEPTION_COLUMNS,
      widths: { 异常原因: 48, 命中群名: 38, 命中聊天内容: 80, 命中质检文件: 36 },
      rowStyle: () => 8,
    },
  ];
  if (includeCleanChats) {
    sheets.push(
      {
        name: "清洗后聊天",
        rows: chatInfo.chats,
        columns: ["来源文件", "有效教师邮箱", "邮箱来源", "发送人名称", "群名/好友昵称", "聊天时间", "聊天内容", "源聊天行号"],
        widths: { 来源文件: 36, 有效教师邮箱: 28, "群名/好友昵称": 38, 聊天内容: 80 },
      },
    );
  }
  sheets.push(
    {
      name: "处理说明",
      rows: explanationRows(listInfo, chatInfo, matchInfo, sourceNames, includeCleanChats),
      columns: ["项目", "值"],
      widths: { 项目: 34, 值: 100 },
    },
  );
  return buildWorkbook(sheets);
}
