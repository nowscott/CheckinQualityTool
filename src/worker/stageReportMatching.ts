import type { ChatRow, CountMap, DataRow } from "./types";
import { normalizeMatchText, normalizeTeacherName } from "./utils";
import type { StageReportListInfo, StageReportTarget } from "./stageReportListParser";

export interface StageReportMatchInfo {
  detailRows: DataRow[];
  teacherRows: DataRow[];
  counts: CountMap;
  unresolvedRows: Array<{ sourceRowNumber: number; reason: string }>;
}

interface NormalizedStageChat {
  chat: ChatRow;
  group: string;
  content: string;
}

interface StageReportMatchedChat {
  chat: ChatRow;
  studentKeyword: string;
  strength: "强匹配" | "弱匹配";
  locations: string[];
}

function teacherKey(target: StageReportTarget) {
  return `${target.teacher}\u0000${target.email}`;
}

function studentKeywords(student: string) {
  const characters = [...student];
  const strong = characters.length >= 2 ? characters.slice(-2).join("") : "";
  const weak = characters.length < 2 ? characters.at(-1) || "" : "";
  return { strong: normalizeMatchText(strong), weak: normalizeMatchText(weak), displayStrong: strong, displayWeak: weak };
}

function matchLocations(chat: NormalizedStageChat, keyword: string) {
  return [chat.group.includes(keyword) ? "群名" : "", chat.content.includes(keyword) ? "聊天内容" : ""].filter(Boolean);
}

function findStudentMatch(chats: NormalizedStageChat[], student: string): StageReportMatchedChat | undefined {
  const { strong, weak, displayStrong, displayWeak } = studentKeywords(student);
  for (const [keyword, display, strength] of [[strong, displayStrong, "强匹配"], [weak, displayWeak, "弱匹配"]] as const) {
    if (!keyword) continue;
    const chat = chats.find((item) => matchLocations(item, keyword).length > 0);
    if (chat) return { chat: chat.chat, studentKeyword: display, strength, locations: matchLocations(chat, keyword) };
  }
  return undefined;
}

export function matchStageReportData(listInfo: StageReportListInfo, chats: ChatRow[]): StageReportMatchInfo {
  const chatsByEmail = new Map<string, NormalizedStageChat[]>();
  const chatsByTeacher = new Map<string, NormalizedStageChat[]>();
  const emailsByTeacher = new Map<string, Set<string>>();
  chats.forEach((chat) => {
    const email = normalizeMatchText(chat.有效教师邮箱);
    const normalized: NormalizedStageChat = {
      chat,
      group: normalizeMatchText(chat["群名/好友昵称"]),
      content: normalizeMatchText(chat.聊天内容),
    };
    if (!chatsByEmail.has(email)) chatsByEmail.set(email, []);
    chatsByEmail.get(email)!.push(normalized);
    const teacher = normalizeTeacherName(chat.发送人名称);
    if (!teacher) return;
    if (!chatsByTeacher.has(teacher)) chatsByTeacher.set(teacher, []);
    chatsByTeacher.get(teacher)!.push(normalized);
    if (!emailsByTeacher.has(teacher)) emailsByTeacher.set(teacher, new Set<string>());
    emailsByTeacher.get(teacher)!.add(email);
  });
  const counts: CountMap = {
    应检查数: listInfo.targets.length,
    已发送数: 0,
    未发送数: 0,
    字段缺失数: 0,
    姓名兜底匹配数: 0,
    姓名匹配歧义数: 0,
    同教师聊天但无学员命中: 0,
    申诉数: 0,
  };
  const detailRows: DataRow[] = [];
  const summary = new Map<string, DataRow>();
  const unresolvedRows: Array<{ sourceRowNumber: number; reason: string }> = [];

  listInfo.targets.forEach((target) => {
    const email = normalizeMatchText(target.email);
    const student = normalizeMatchText(target.matchedStudent);
    const teacher = normalizeTeacherName(target.teacher);
    const nameEmails = emailsByTeacher.get(teacher);
    const nameAmbiguous = !email && Boolean(nameEmails && nameEmails.size > 1);
    const candidates = email ? chatsByEmail.get(email) || [] : chatsByTeacher.get(teacher) || [];
    const matched = student ? findStudentMatch(candidates, student) : undefined;
    const missing = !target.teacher || !student || nameAmbiguous;
    const missingReasons = [
      !target.teacher ? "教师姓名为空" : "",
      !student ? "学员姓名为空" : "",
      nameAmbiguous ? "教师姓名对应多个聊天邮箱" : "",
    ].filter(Boolean);
    if (missing) unresolvedRows.push({ sourceRowNumber: target.sourceRowNumber, reason: missingReasons.join("、") });
    const conclusion = missing ? "字段缺失" : matched ? "已发送" : "未发送";
    if (conclusion === "已发送") counts.已发送数 += 1;
    else if (conclusion === "字段缺失") counts.字段缺失数 += 1;
    else counts.未发送数 += 1;
    if (!email && matched) counts.姓名兜底匹配数 += 1;
    if (nameAmbiguous) counts.姓名匹配歧义数 += 1;
    if (!matched && candidates.length && !missing) counts.同教师聊天但无学员命中 += 1;
    if (target.appeal) counts.申诉数 += 1;

    detailRows.push({
      ...target.original,
      本次检查结论: conclusion,
      命中关键词: matched ? `${matched.studentKeyword}（${matched.strength}）` : "",
      命中聊天内容: matched?.chat.聊天内容 || "",
      命中群名: matched?.chat["群名/好友昵称"] || "",
      命中时间: matched?.chat.聊天时间 || "",
      来源文件: matched?.chat.来源文件 || "",
      源聊天行号: matched?.chat.源聊天行号 || "",
      原分母行号: target.sourceRowNumber,
      去重合并行号: target.duplicateRows,
      教师匹配方式: email ? "邮箱匹配" : nameAmbiguous ? "姓名匹配歧义" : "姓名兜底匹配",
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
  return { detailRows, teacherRows, counts, unresolvedRows };
}
