import type { DataRow } from "./types";
import { normalizeMatchText, text } from "./utils";

export function trainingPersonKey(value: unknown) {
  return normalizeMatchText(value)
    .replace(/\s+/g, "")
    .replace(/老师$/u, "");
}

export function findTrainingLeadGroups(teacherRows: DataRow[], lead: unknown) {
  const leadKey = trainingPersonKey(lead);
  return [...new Set(
    teacherRows
      .filter((row) => trainingPersonKey(row.教师姓名) === leadKey)
      .map((row) => text(row.教研组))
      .filter(Boolean),
  )].sort((a, b) => a.localeCompare(b, "zh-CN"));
}
