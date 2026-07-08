import { findSheet, headerMap } from "./excelReader";
import type { ChatInfo, ChatRow } from "./types";
import { emailValue, text } from "./utils";

const QUOTE_SEPARATOR = /(?:-\s*){8,}|[—－-]{12,}/;
const QUOTE_PREFIX = /^\s*[「『][\s\S]{0,500}?[：:]/;

export function preprocessChats(workbook: SheetJsWorkbook, sourceFile = ""): ChatInfo {
  const found = findSheet(workbook, ["聊天类型", "发送方", "聊天内容"]);
  const rows = found.rows;
  const map = headerMap(rows[0]);
  const index = (name: string) => (map.get(name) || [])[0] ?? -1;
  const columns = {
    name: index("姓名"),
    email: index("邮箱"),
    type: index("聊天类型"),
    sender: index("发送方"),
    group: index("群名/好友昵称"),
    groupSender: index("群聊发送人名称"),
    groupEmail: index("群聊发送人邮箱"),
    time: index("聊天时间"),
    content: index("聊天内容"),
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
