import { findSheet, headerMap } from "./excelReader";
import type { CellValue, CountMap } from "./types";
import { cleanStudentName, normalizeMatchText, text } from "./utils";

export interface ReminderAppealRecord {
  teacher: string;
  student: string;
  matchedStudent: string;
  reason: string;
  description: string;
  status: string;
  sourceRowNumber: number;
}

export interface ReminderAppealInfo {
  records: ReminderAppealRecord[];
  byKey: Map<string, ReminderAppealRecord>;
  counts: CountMap;
  sheetName: string;
}

function findFirstIndex(map: Map<string, number[]>, aliases: readonly string[]) {
  for (const alias of aliases) {
    const index = (map.get(alias) || [])[0];
    if (index != null) return index;
  }
  return -1;
}

function splitStudentNames(value: unknown) {
  const raw = text(value);
  if (!raw) return [];
  return raw
    .split(/[、，,;；/／\n\r]+/u)
    .flatMap((part) => {
      const current = part.trim();
      if (!current) return [];
      const spaceParts = current.split(/\s+/u).filter(Boolean);
      const shouldSplitSpaces = spaceParts.length > 1 && spaceParts.every((item) => item.length >= 2);
      return shouldSplitSpaces ? spaceParts : [current];
    })
    .map((item) => item.trim())
    .filter(Boolean);
}

export function appealKey(teacher: unknown, student: unknown) {
  const normalizedTeacher = normalizeMatchText(teacher).replace(/\s+/g, "").replace(/[0-9０-９]+$/u, "");
  return `${normalizedTeacher}\u0000${normalizeMatchText(student)}`;
}

export function isSentAppeal(record: ReminderAppealRecord) {
  return [record.reason, record.status].some((value) => normalizeMatchText(value).includes("已发送"));
}

export function buildReminderAppeals(workbook: SheetJsWorkbook): ReminderAppealInfo {
  const found = findSheet(workbook, ["教师姓名", "学生姓名"]);
  const rows = found.rows;
  const map = headerMap(rows[0]);
  const column = {
    teacher: findFirstIndex(map, ["教师姓名", "授课教师", "老师姓名"]),
    student: findFirstIndex(map, ["学生姓名", "学员姓名"]),
    reason: findFirstIndex(map, ["申诉原因"]),
    description: findFirstIndex(map, ["申诉原因描述", "申诉说明"]),
    passed: findFirstIndex(map, ["申诉是否通过", "是否通过"]),
  };
  const counts: CountMap = {
    原始申诉行数: Math.max(0, rows.length - 1),
    申诉行数: 0,
    申诉拆分学员数: 0,
    申诉跳过行数: 0,
    申诉去重行数: 0,
  };
  const records: ReminderAppealRecord[] = [];
  const byKey = new Map<string, ReminderAppealRecord>();

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] as CellValue[];
    const teacher = text(row[column.teacher]);
    const students = splitStudentNames(row[column.student]);
    if (!teacher || !students.length) {
      counts.申诉跳过行数 += 1;
      continue;
    }
    counts.申诉拆分学员数 += students.length - 1;
    students.forEach((rawStudent) => {
      const cleaned = cleanStudentName(rawStudent);
      const record: ReminderAppealRecord = {
        teacher,
        student: cleaned.original,
        matchedStudent: cleaned.cleaned || cleaned.original,
        reason: text(row[column.reason]),
        description: text(row[column.description]),
        status: column.passed >= 0 ? text(row[column.passed]) : "",
        sourceRowNumber: rowIndex + 1,
      };
      const keys = [
        appealKey(record.teacher, record.student),
        appealKey(record.teacher, record.matchedStudent),
      ];
      let duplicated = false;
      keys.forEach((key) => {
        if (byKey.has(key)) duplicated = true;
        byKey.set(key, record);
      });
      if (duplicated) counts.申诉去重行数 += 1;
      records.push(record);
      counts.申诉行数 += 1;
    });
  }

  return { records, byKey, counts, sheetName: found.name };
}
