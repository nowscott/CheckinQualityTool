import type { TargetRow, Whitelist, WhitelistEntry } from "./types";
import { cleanStudentName, normalizeMatchText, text } from "./utils";

export const WHITELIST_COLUMNS = [
  "学员号",
  "学员姓名",
  "匹配学员姓名",
  "处理方式",
  "匹配别名",
  "说明",
] as const;

function parseCsv(value: string) {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  const source = String(value || "").replace(/^\uFEFF/, "");
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (quoted) {
      if (character === '"' && source[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        field += character;
      }
    } else if (character === '"') {
      quoted = true;
    } else if (character === ",") {
      row.push(field);
      field = "";
    } else if (character === "\n") {
      row.push(field.replace(/\r$/, ""));
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += character;
    }
  }
  if (field || row.length) {
    row.push(field.replace(/\r$/, ""));
    rows.push(row);
  }
  return rows;
}

export function buildWhitelist(csvText: string): Whitelist {
  const rows = parseCsv(csvText).filter((row) => row.some((value) => text(value)));
  const empty = {
    entries: [],
    byStudentId: new Map<string, WhitelistEntry>(),
    byOriginalStudentName: new Map<string, WhitelistEntry>(),
    byStudentName: new Map<string, WhitelistEntry>(),
  };
  if (!rows.length) return empty;

  const headers = rows[0].map(text);
  const index = (name: string) => headers.indexOf(name);
  const entries: WhitelistEntry[] = [];
  const byStudentId = new Map<string, WhitelistEntry>();
  const byOriginalStudentName = new Map<string, WhitelistEntry>();
  const byStudentName = new Map<string, WhitelistEntry>();
  for (const row of rows.slice(1)) {
    const studentName = cleanStudentName(row[index("学员姓名")]);
    const studentId = text(row[index("学员号")]);
    if (!studentId && !studentName.original) continue;
    const mode = text(row[index("处理方式")]) || "已发送";
    const aliases = text(row[index("匹配别名")])
      .split(/[|；;]/u)
      .map((value) => cleanStudentName(value).cleaned)
      .filter(Boolean);
    const entry: WhitelistEntry = {
      学员号: studentId,
      学员姓名: studentName.original,
      匹配学员姓名: mode === "保留原名" ? studentName.original.replace(/\s+/g, "") : studentName.cleaned,
      处理方式: mode,
      匹配别名: aliases.join("|"),
      匹配别名关键词: aliases.map((value) => normalizeMatchText(value.slice(-2))),
      说明: text(row[index("说明")]),
    };
    entries.push(entry);
    const studentKey = normalizeMatchText(entry.匹配学员姓名 || entry.学员姓名);
    const originalStudentKey = normalizeMatchText(entry.学员姓名);
    if (entry.学员号) byStudentId.set(normalizeMatchText(entry.学员号), entry);
    if (originalStudentKey) byOriginalStudentName.set(originalStudentKey, entry);
    if (studentKey) byStudentName.set(studentKey, entry);
  }
  return { entries, byStudentId, byOriginalStudentName, byStudentName };
}

export function findPreCleanWhitelistEntry(
  studentId: string,
  originalStudentName: string,
  whitelist: Whitelist,
) {
  return (
    (studentId && whitelist.byStudentId.get(normalizeMatchText(studentId))) ||
    whitelist.byOriginalStudentName.get(normalizeMatchText(originalStudentName)) ||
    null
  );
}

export function findWhitelistEntry(target: TargetRow, whitelist: Whitelist) {
  const studentKey = normalizeMatchText(target.学员姓名 || target.原始学员姓名);
  return (
    (target.学员号 && whitelist.byStudentId.get(normalizeMatchText(target.学员号))) ||
    whitelist.byStudentName.get(studentKey) ||
    null
  );
}
