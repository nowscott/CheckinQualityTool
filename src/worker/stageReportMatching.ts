import type { ChatRow, CountMap, DataRow } from "./types";
import { normalizeMatchText } from "./utils";
import type { StageReportListInfo, StageReportTarget } from "./stageReportListParser";

const KEYWORD = "阶段性报告";

export interface StageReportMatchInfo {
  detailRows: DataRow[];
  teacherRows: DataRow[];
  counts: CountMap;
}

function teacherKey(target: StageReportTarget) {
  return `${target.teacher}\u0000${target.email}`;
}

export function matchStageReportData(listInfo: StageReportListInfo, chats: ChatRow[]): StageReportMatchInfo {
  const chatsByEmail = new Map<string, ChatRow[]>();
  chats.forEach((chat) => {
    const email = normalizeMatchText(chat.有效教师邮箱);
    if (!chatsByEmail.has(email)) chatsByEmail.set(email, []);
    chatsByEmail.get(email)!.push(chat);
  });
  const counts: CountMap = {
    应检查数: listInfo.targets.length,
    已发送数: 0,
    未发送数: 0,
    字段缺失数: 0,
    关键词命中但无学员命中: 0,
    申诉数: 0,
  };
  const detailRows: DataRow[] = [];
  const summary = new Map<string, DataRow>();

  listInfo.targets.forEach((target) => {
    const email = normalizeMatchText(target.email);
    const student = normalizeMatchText(target.matchedStudent);
    const candidates = chatsByEmail.get(email) || [];
    const keywordChats = candidates.filter((chat) => normalizeMatchText(chat.聊天内容).includes(KEYWORD));
    const matched = student
      ? keywordChats.find((chat) => normalizeMatchText(chat.聊天内容).includes(student))
      : undefined;
    const missing = !target.teacher || !email || !student;
    const conclusion = missing ? "字段缺失" : matched ? "已发送" : "未发送";
    if (conclusion === "已发送") counts.已发送数 += 1;
    else if (conclusion === "字段缺失") counts.字段缺失数 += 1;
    else counts.未发送数 += 1;
    if (!matched && keywordChats.length && !missing) counts.关键词命中但无学员命中 += 1;
    if (target.appeal) counts.申诉数 += 1;

    detailRows.push({
      ...target.original,
      本次检查结论: conclusion,
      命中关键词: matched ? `${KEYWORD}；${target.matchedStudent}` : "",
      命中聊天内容: matched?.聊天内容 || "",
      命中群名: matched?.["群名/好友昵称"] || "",
      命中时间: matched?.聊天时间 || "",
      来源文件: matched?.来源文件 || "",
      源聊天行号: matched?.源聊天行号 || "",
      原分母行号: target.sourceRowNumber,
      去重合并行号: target.duplicateRows,
    });

    const key = teacherKey(target);
    if (!summary.has(key)) {
      summary.set(key, { 教师姓名: target.teacher, 教师邮箱: target.email, 应发送数: 0, 已发送数: 0, 未发送数: 0, 申诉数: 0 });
    }
    const teacherRow = summary.get(key)!;
    teacherRow.应发送数 = Number(teacherRow.应发送数) + 1;
    if (conclusion === "已发送") teacherRow.已发送数 = Number(teacherRow.已发送数) + 1;
    else teacherRow.未发送数 = Number(teacherRow.未发送数) + 1;
    if (target.appeal) teacherRow.申诉数 = Number(teacherRow.申诉数) + 1;
  });

  const teacherRows: DataRow[] = [...summary.values()]
    .map((row): DataRow => ({ ...row, 发送率: Number(row.应发送数) ? Number(row.已发送数) / Number(row.应发送数) : 1 }))
    .sort((a, b) => String(a.教师姓名).localeCompare(String(b.教师姓名), "zh-CN") || String(a.教师邮箱).localeCompare(String(b.教师邮箱), "zh-CN"));
  return { detailRows, teacherRows, counts };
}
