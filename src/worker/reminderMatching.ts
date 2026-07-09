import { REMINDER_MATCH_RULES, REMINDER_PASS_RATE } from "./reminderConfig";
import type { ChatRow, CountMap, DataRow } from "./types";
import { displayValue, normalizeMatchText, sortDate } from "./utils";
import { appealKey, type ReminderAppealInfo } from "./reminderAppealParser";
import type { ReminderListInfo, ReminderTarget } from "./reminderListParser";

export interface ReminderMatchedChat extends ChatRow {
  匹配优先级: number;
  匹配方式: string;
  命中位置: string;
  命中关键词: string;
}

export interface ReminderMatchInfo {
  studentRows: DataRow[];
  teacherRows: DataRow[];
  trainingRows: DataRow[];
  assistantRows: DataRow[];
  exceptionRows: DataRow[];
  counts: CountMap;
}

function rate(sent: number, total: number) {
  return total ? sent / total : 0;
}

function rateText(value: number) {
  return `${(value * 100).toFixed(1)}%`;
}

function passText(value: number) {
  return value >= REMINDER_PASS_RATE ? "是" : "否";
}

function isAppealed(row: DataRow) {
  return row.是否发送 === "已申诉" || row.匹配状态 === "已申诉" ||
    row.是否发送 === "已申诉通过" || row.匹配状态 === "已申诉通过";
}

function compactKey(parts: unknown[]) {
  return parts.map((part) => String(part || "")).join("\u0000");
}

function normalizeTeacherName(value: unknown) {
  return normalizeMatchText(value).replace(/\s+/g, "").replace(/[0-9０-９]+$/u, "");
}

function assistantOwnTeachingGroups(rows: DataRow[]) {
  const groups = new Map<string, unknown>();
  rows.forEach((row) => {
    const teacher = normalizeTeacherName(row.教师姓名);
    const assistant = normalizeTeacherName(row.助理主管);
    if (teacher && assistant && teacher === assistant && row.教研组) {
      groups.set(assistant, row.教研组);
    }
  });
  return groups;
}

function addException(rows: DataRow[], target: ReminderTarget, type: string, reason: string, chat?: ChatRow) {
  rows.push({
    异常类型: type,
    异常原因: reason,
    质检序号: target.id,
    授课教师: target.授课教师,
    学员姓名: target.学员姓名,
    匹配学员姓名: target.匹配学员姓名,
    教研组: target.教研组,
    师训组长: target.师训组长,
    助理主管: target.助理主管,
    源名单行号: target.源名单行号,
    命中群名: chat?.["群名/好友昵称"] || "",
    命中聊天内容: chat?.聊天内容 || "",
    命中聊天时间: chat ? displayValue(chat.聊天时间) : "",
    命中质检文件: chat?.来源文件 || "",
    源聊天行号: chat?.源聊天行号 || "",
  });
}

export function buildReminderGroupSummary(
  rows: DataRow[],
  keyColumns: readonly string[],
  outputColumns: readonly string[],
) {
  const grouped = new Map<string, DataRow & { 应发送数: number; 已发送数: number; 申诉数: number }>();
  for (const row of rows) {
    const key = compactKey(keyColumns.map((column) => row[column]));
    if (!grouped.has(key)) {
      const item: DataRow & { 应发送数: number; 已发送数: number; 申诉数: number } = {
        应发送数: 0,
        已发送数: 0,
        申诉数: 0,
      };
      outputColumns.forEach((column) => {
        item[column] = row[column] || "";
      });
      grouped.set(key, item);
    }
    const item = grouped.get(key)!;
    if (isAppealed(row)) {
      item.申诉数 += 1;
      continue;
    }
    item.应发送数 += 1;
    if (row.是否发送 === "是") item.已发送数 += 1;
  }
  return [...grouped.values()]
    .map((item) => {
      const currentRate = rate(item.已发送数, item.应发送数);
      return {
        ...item,
        发送差额: item.应发送数 - item.已发送数,
        发送率: rateText(currentRate),
        是否达标: passText(currentRate),
      };
    })
    .sort((a, b) => outputColumns.map((column) =>
      String((a as DataRow)[column] || "").localeCompare(String((b as DataRow)[column] || ""), "zh-CN"),
    ).find((value) => value !== 0) || 0);
}

export function matchReminderData(
  listInfo: ReminderListInfo,
  chats: ChatRow[],
  appealInfo?: ReminderAppealInfo,
): ReminderMatchInfo {
  const normalizedChats = chats.map((chat) => ({
    chat,
    group: normalizeMatchText(chat["群名/好友昵称"]),
    content: normalizeMatchText(chat.聊天内容),
  }));
  const studentCounts = new Map<string, number>();
  listInfo.targets.forEach((target) => {
    const student = normalizeMatchText(target.匹配学员姓名);
    studentCounts.set(student, (studentCounts.get(student) || 0) + 1);
  });

  const studentRows: DataRow[] = [];
  const exceptionRows: DataRow[] = [];
  const counts: CountMap = {
    应发送数: 0,
    已发送数: 0,
    未发送数: 0,
    申诉数: 0,
    待核对数: 0,
    发送人教师学员命中: 0,
    群名教师学员命中: 0,
    群名唯一学员命中: 0,
    内容唯一学员命中: 0,
    多条质检命中: 0,
    无法唯一匹配: 0,
    字段缺失: 0,
    异常明细行数: 0,
  };

  for (const target of listInfo.targets) {
    const appeal = appealInfo?.byKey.get(appealKey(target.授课教师, target.学员姓名)) ||
      appealInfo?.byKey.get(appealKey(target.授课教师, target.匹配学员姓名));
    if (appeal) {
      counts.申诉数 += 1;
      const appealReason = [appeal.reason, appeal.description].filter(Boolean).join("：") || "已申诉";
      studentRows.push({
        质检序号: target.id,
        教师姓名: target.授课教师,
        教师邮箱: target.教师邮箱,
        教研组: target.教研组,
        师训组长: target.师训组长,
        助理主管: target.助理主管,
        学员姓名: target.学员姓名,
        匹配学员姓名: target.匹配学员姓名,
        姓名清洗说明: target.姓名清洗说明,
        校区: target.校区,
        年级: target.年级,
        学管姓名: target.学管,
        新老生季度: target.新老生季度,
        课时: target.课时,
        是否发送: "已申诉",
        匹配状态: "已申诉",
        匹配方式: "已申诉，剔除分母",
        申诉情况说明: appealReason,
        申诉状态: appeal.status,
        申诉源行号: appeal.sourceRowNumber,
        异常原因: "",
        命中位置: "",
        命中关键词: "",
        命中群名: "",
        命中聊天时间: "",
        命中质检文件: "",
        发送人名称: "",
        发送人邮箱: "",
        源名单行号: target.源名单行号,
        去重合并行号: target.去重合并行号,
        源聊天行号: "",
        匹配消息数: 0,
      });
      continue;
    }
    counts.应发送数 += 1;
    const student = normalizeMatchText(target.匹配学员姓名);
    const teacher = normalizeTeacherName(target.授课教师);
    const isUniqueStudent = (studentCounts.get(student) || 0) === 1;
    const matches: ReminderMatchedChat[] = [];

    for (const item of normalizedChats) {
      const groupHasStudent = Boolean(student && item.group.includes(student));
      const contentHasStudent = Boolean(student && item.content.includes(student));
      const groupHasTeacher = Boolean(teacher && item.group.includes(teacher));
      const contentHasTeacher = Boolean(teacher && item.content.includes(teacher));
      const hasStudent = groupHasStudent || contentHasStudent;
      const hasTeacher = groupHasTeacher || contentHasTeacher;
      const senderTeacher = normalizeTeacherName(item.chat.发送人名称);
      const senderMatchesThisTeacher = Boolean(teacher && senderTeacher && senderTeacher === teacher);
      if (hasStudent && senderMatchesThisTeacher) {
        matches.push({
          ...item.chat,
          匹配优先级: 1,
          匹配方式: REMINDER_MATCH_RULES.senderTeacherAndStudent,
          命中位置: [groupHasStudent ? "群聊名称" : "", contentHasStudent ? "聊天内容" : ""]
            .filter(Boolean)
            .join("+"),
          命中关键词: `${target.授课教师}+${target.匹配学员姓名}`,
        });
        continue;
      }
      if (hasStudent && hasTeacher) {
        matches.push({
          ...item.chat,
          匹配优先级: 2,
          匹配方式: REMINDER_MATCH_RULES.groupStudentAndTeacher,
          命中位置: [
            groupHasStudent || groupHasTeacher ? "群聊名称" : "",
            contentHasStudent || contentHasTeacher ? "聊天内容" : "",
          ].filter(Boolean).join("+"),
          命中关键词: `${target.授课教师}+${target.匹配学员姓名}`,
        });
        continue;
      }
      if (groupHasStudent && isUniqueStudent) {
        matches.push({
          ...item.chat,
          匹配优先级: 3,
          匹配方式: REMINDER_MATCH_RULES.uniqueStudentInGroup,
          命中位置: "群聊名称",
          命中关键词: target.匹配学员姓名,
        });
        continue;
      }
      if (contentHasStudent && isUniqueStudent) {
        matches.push({
          ...item.chat,
          匹配优先级: 4,
          匹配方式: REMINDER_MATCH_RULES.uniqueStudentInContent,
          命中位置: "聊天内容",
          命中关键词: target.匹配学员姓名,
        });
        continue;
      }
    }

    matches.sort(
      (a, b) =>
        a.匹配优先级 - b.匹配优先级 ||
        sortDate(a.聊天时间) - sortDate(b.聊天时间) ||
        String(a.来源文件 || "").localeCompare(String(b.来源文件 || ""), "zh-CN") ||
        a.源聊天行号 - b.源聊天行号,
    );
    const best = matches[0];
    const missing = [
      target.教研组 ? "" : "教研组",
      target.师训组长 ? "" : "师训组长",
      target.助理主管 ? "" : "助理主管",
    ].filter(Boolean);
    const sent = Boolean(best);
    const matchStatus = sent ? "已发送" : "未发送";
    const reason = sent ? "" : "未找到可自动判定的聊天记录";

    if (sent) {
      counts.已发送数 += 1;
      if (best.匹配优先级 === 1) counts.发送人教师学员命中 += 1;
      if (best.匹配优先级 === 2) counts.群名教师学员命中 += 1;
      if (best.匹配优先级 === 3) counts.群名唯一学员命中 += 1;
      if (best.匹配优先级 === 4) counts.内容唯一学员命中 += 1;
    } else {
      counts.未发送数 += 1;
    }
    if (matches.length > 1) counts.多条质检命中 += 1;
    if (missing.length) {
      counts.字段缺失 += 1;
      addException(exceptionRows, target, "字段缺失", `分母记录缺少：${missing.join("、")}`);
    }

    studentRows.push({
      质检序号: target.id,
      教师姓名: target.授课教师,
      教师邮箱: target.教师邮箱,
      教研组: target.教研组,
      师训组长: target.师训组长,
      助理主管: target.助理主管,
      学员姓名: target.学员姓名,
      匹配学员姓名: target.匹配学员姓名,
      姓名清洗说明: target.姓名清洗说明,
      校区: target.校区,
      年级: target.年级,
      学管姓名: target.学管,
      新老生季度: target.新老生季度,
      课时: target.课时,
      是否发送: sent ? "是" : "否",
      匹配状态: matchStatus,
      匹配方式: best?.匹配方式 || "",
      异常原因: reason,
      命中位置: best?.命中位置 || "",
      命中关键词: best?.命中关键词 || "",
      命中群名: best?.["群名/好友昵称"] || "",
      命中聊天时间: best ? displayValue(best.聊天时间) : "",
      命中质检文件: best?.来源文件 || "",
      发送人名称: best?.发送人名称 || "",
      发送人邮箱: best?.有效教师邮箱 || "",
      源名单行号: target.源名单行号,
      去重合并行号: target.去重合并行号,
      源聊天行号: best?.源聊天行号 || "",
      匹配消息数: matches.length,
    });
  }

  const teacherRows = buildReminderGroupSummary(
    studentRows,
    ["教师姓名", "教师邮箱", "教研组", "师训组长", "助理主管", "校区"],
    ["教师姓名", "教师邮箱", "教研组", "师训组长", "助理主管", "校区"],
  );
  const trainingRows = buildReminderGroupSummary(
    studentRows,
    ["教研组", "师训组长"],
    ["教研组", "师训组长"],
  );
  const assistantGroups = assistantOwnTeachingGroups(studentRows);
  const assistantSourceRows = studentRows.map((row) => {
    const ownGroup = assistantGroups.get(normalizeTeacherName(row.助理主管));
    return ownGroup ? { ...row, 教研组: ownGroup } : row;
  });
  const assistantRows = buildReminderGroupSummary(
    assistantSourceRows,
    ["助理主管"],
    ["教研组", "助理主管"],
  );
  counts.异常明细行数 = exceptionRows.length;
  return { studentRows, teacherRows, trainingRows, assistantRows, exceptionRows, counts };
}
