import { findSheet, headerMap } from "./excelReader";
import type { CellValue, CountMap } from "./types";
import { cleanStudentName, normalizeMatchText, text } from "./utils";

export interface ReminderAppealRecord {
  teacher: string;
  student: string;
  matchedStudent: string;
  reason: string;
  description: string;
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

function passed(value: unknown) {
  const current = text(value);
  return ["是", "通过", "已通过", "申诉通过", "true", "TRUE", "1"].includes(current);
}

export function appealKey(teacher: unknown, student: unknown) {
  return `${normalizeMatchText(teacher)}\u0000${normalizeMatchText(student)}`;
}

export function buildReminderAppeals(workbook: SheetJsWorkbook): ReminderAppealInfo {
  const found = findSheet(workbook, ["教师姓名", "申诉是否通过"]);
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
    申诉通过行数: 0,
    申诉跳过行数: 0,
    申诉去重行数: 0,
  };
  const records: ReminderAppealRecord[] = [];
  const byKey = new Map<string, ReminderAppealRecord>();

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] as CellValue[];
    const teacher = text(row[column.teacher]);
    const rawStudent = text(row[column.student]);
    if (!teacher || !rawStudent || !passed(row[column.passed])) {
      counts.申诉跳过行数 += 1;
      continue;
    }
    const cleaned = cleanStudentName(rawStudent);
    const record: ReminderAppealRecord = {
      teacher,
      student: cleaned.original,
      matchedStudent: cleaned.cleaned || cleaned.original,
      reason: text(row[column.reason]),
      description: text(row[column.description]),
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
    counts.申诉通过行数 += 1;
  }

  return { records, byKey, counts, sheetName: found.name };
}
