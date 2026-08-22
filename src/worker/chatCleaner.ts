import { headerMap, sheetCandidates, type FoundSheet } from "./excelReader";
import type { ChatInfo, ChatRow } from "./types";
import { emailValue, text } from "./utils";

const QUOTE_SEPARATOR = /(?:-\s*){8,}|[—－-]{12,}/;
const QUOTE_PREFIX = /^\s*[「『][\s\S]{0,500}?[：:]/;

const CHAT_HEADERS = {
  name: ["姓名", "发送人名称", "发送人"],
  email: ["邮箱", "发送人邮箱"],
  type: ["聊天类型", "类型"],
  sender: ["发送方", "消息发送方"],
  group: ["群名/好友昵称", "群聊名称", "群名", "好友昵称"],
  groupSender: ["群聊发送人名称", "发送人名称", "发送人"],
  groupEmail: ["群聊发送人邮箱", "发送人邮箱", "邮箱"],
  time: ["聊天时间", "发送时间", "消息时间"],
  content: ["聊天内容", "消息内容", "内容"],
} as const;

function firstIndex(map: Map<string, number[]>, aliases: readonly string[]) {
  for (const alias of aliases) {
    const index = (map.get(alias) || [])[0];
    if (index != null) return index;
  }
  return -1;
}

function hasAnyHeader(map: Map<string, number[]>, aliases: readonly string[]) {
  return aliases.some((alias) => map.has(alias));
}

function findChatSheet(workbook: SheetJsWorkbook): FoundSheet {
  const required = [CHAT_HEADERS.type, CHAT_HEADERS.sender, CHAT_HEADERS.content];
  const candidates = sheetCandidates(workbook);
  const ranked = candidates
    .filter((candidate) => required.every((aliases) => hasAnyHeader(candidate.map, aliases)))
    .map((candidate) => {
      const email = firstIndex(candidate.map, CHAT_HEADERS.groupEmail) >= 0
        ? firstIndex(candidate.map, CHAT_HEADERS.groupEmail)
        : firstIndex(candidate.map, CHAT_HEADERS.email);
      const content = firstIndex(candidate.map, CHAT_HEADERS.content);
      const usableRows = candidate.rows.slice(1).filter((row) => email >= 0 && content >= 0 && emailValue(row[email]) && text(row[content])).length;
      return {
        candidate,
        score: Number(hasAnyHeader(candidate.map, CHAT_HEADERS.email)) +
          Number(hasAnyHeader(candidate.map, CHAT_HEADERS.groupEmail)) +
          Number(hasAnyHeader(candidate.map, CHAT_HEADERS.group)) +
          Number(hasAnyHeader(candidate.map, CHAT_HEADERS.time)),
        usableRows,
      };
    })
    .sort((a, b) => b.score - a.score || b.usableRows - a.usableRows || b.candidate.rows.length - a.candidate.rows.length || a.candidate.order - b.candidate.order);
  if (ranked.length && (ranked.length === 1 || ranked[0].score > ranked[1].score || ranked[0].usableRows > ranked[1].usableRows || ranked[0].candidate.rows.length > ranked[1].candidate.rows.length)) {
    return ranked[0].candidate;
  }
  const seenSheets = candidates.map((candidate) => candidate.name);
  throw new Error(
    "找不到聊天记录工作表：需要第一行包含“聊天类型/类型”、“发送方/消息发送方”和“聊天内容/消息内容/内容”。" +
      `已检查 Sheet：${seenSheets.join("、") || "无"}`,
  );
}

export function preprocessChats(workbook: SheetJsWorkbook, sourceFile = ""): ChatInfo {
  const found = findChatSheet(workbook);
  const rows = found.rows;
  const map = headerMap(rows[0]);
  const columns = {
    name: firstIndex(map, CHAT_HEADERS.name),
    email: firstIndex(map, CHAT_HEADERS.email),
    type: firstIndex(map, CHAT_HEADERS.type),
    sender: firstIndex(map, CHAT_HEADERS.sender),
    group: firstIndex(map, CHAT_HEADERS.group),
    groupSender: firstIndex(map, CHAT_HEADERS.groupSender),
    groupEmail: firstIndex(map, CHAT_HEADERS.groupEmail),
    time: firstIndex(map, CHAT_HEADERS.time),
    content: firstIndex(map, CHAT_HEADERS.content),
  };
  const chats: ChatRow[] = [];
  const counts = {
    原始聊天行数: Math.max(0, rows.length - 1),
    检测到引用回复: 0,
    删除私聊: 0,
    删除无有效邮箱: 0,
    删除发送方非员工: 0,
    删除引用回复: 0,
    清洗后聊天行数: 0,
  };

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const chatType = text(row[columns.type]);
    const sender = text(row[columns.sender]);
    const primaryEmail = emailValue(row[columns.groupEmail]);
    const fallbackEmail = emailValue(row[columns.email]);
    const effectiveEmail = primaryEmail || fallbackEmail;
    const content = text(row[columns.content]);
    const isQuotedReply = QUOTE_SEPARATOR.test(content) && QUOTE_PREFIX.test(content);
    if (isQuotedReply) counts.检测到引用回复 += 1;
    if (chatType === "私聊") {
      counts.删除私聊 += 1;
      continue;
    }
    if (!effectiveEmail) {
      counts.删除无有效邮箱 += 1;
      continue;
    }
    if (sender && sender !== "员工") {
      counts.删除发送方非员工 += 1;
      continue;
    }
    if (isQuotedReply) {
      counts.删除引用回复 += 1;
      continue;
    }
    chats.push({
      来源文件: sourceFile,
      有效教师邮箱: effectiveEmail,
      邮箱来源: primaryEmail ? "群聊发送人邮箱" : "邮箱",
      发送人名称: text(row[columns.groupSender]) || text(row[columns.name]),
      "群名/好友昵称": text(row[columns.group]),
      聊天时间: row[columns.time],
      聊天内容: content,
      源聊天行号: rowIndex + 1,
    });
  }
  counts.清洗后聊天行数 = chats.length;
  return { chats, counts, sheetName: found.name };
}
