import type { DataRow } from "./types";
import { normalizeMatchText, text } from "./utils";

function personKey(value: unknown) {
  return normalizeMatchText(value)
    .replace(/\s+/g, "")
    .replace(/[0-9０-９]+$/u, "");
}

export function findTrainingLeadGroups(teacherRows: DataRow[], lead: unknown) {
  const leadKey = personKey(lead);
  return [...new Set(
    teacherRows
      .filter((row) => personKey(row.教师姓名) === leadKey)
      .map((row) => text(row.教研组))
      .filter(Boolean),
  )].sort((a, b) => a.localeCompare(b, "zh-CN"));
}
