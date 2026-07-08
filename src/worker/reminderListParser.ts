import { findSheet, headerMap } from "./excelReader";
import { REMINDER_LIST_HEADERS } from "./reminderConfig";
import type { CellValue, CountMap } from "./types";
import { cleanStudentName, text } from "./utils";

export interface ReminderTarget {
  id: number;
  授课教师: string;
  教师邮箱: string;
  教研组: string;
  师训组长: string;
  助理主管: string;
  学员姓名: string;
  匹配学员姓名: string;
  姓名清洗说明: string;
  新老生季度: string;
  年级: string;
  学管: string;
  课时: string;
  校区: string;
  源名单行号: number;
  去重合并行号: string;
}

export interface ReminderListInfo {
  targets: ReminderTarget[];
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

function rowDedupKey(row: CellValue[], width: number) {
  return Array.from({ length: width }, (_, index) => text(row[index])).join("\u0000");
}

export function buildReminderTargets(workbook: SheetJsWorkbook): ReminderListInfo {
  const found = findSheet(workbook, ["授课教师", "学员姓名"]);
  const rows = found.rows;
  const map = headerMap(rows[0]);
  const column = {
    teacher: findFirstIndex(map, REMINDER_LIST_HEADERS.teacher),
    teacherEmail: findFirstIndex(map, REMINDER_LIST_HEADERS.teacherEmail),
    teachingGroup: findFirstIndex(map, REMINDER_LIST_HEADERS.teachingGroup),
    trainingLead: findFirstIndex(map, REMINDER_LIST_HEADERS.trainingLead),
    assistantLead: findFirstIndex(map, REMINDER_LIST_HEADERS.assistantLead),
    student: findFirstIndex(map, REMINDER_LIST_HEADERS.student),
    studentType: findFirstIndex(map, REMINDER_LIST_HEADERS.studentType),
    grade: findFirstIndex(map, REMINDER_LIST_HEADERS.grade),
    counselor: findFirstIndex(map, REMINDER_LIST_HEADERS.counselor),
    classHours: findFirstIndex(map, REMINDER_LIST_HEADERS.classHours),
    campus: findFirstIndex(map, REMINDER_LIST_HEADERS.campus),
  };
  const counts: CountMap = {
    原始分母行数: Math.max(0, rows.length - 1),
    整行重复行数: 0,
    跳过教师或学员为空: 0,
    字段缺失记录: 0,
    去重后应发送数: 0,
  };
  const width = Math.max(...Object.values(column), rows[0].length - 1, 0) + 1;
  const seen = new Map<string, ReminderTarget>();
  const targets: ReminderTarget[] = [];

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const sourceRowNumber = rowIndex + 1;
    const key = rowDedupKey(row, width);
    const existed = seen.get(key);
    if (existed) {
      counts.整行重复行数 += 1;
      existed.去重合并行号 = [existed.去重合并行号, sourceRowNumber].filter(Boolean).join("、");
      continue;
    }

    const teacher = text(row[column.teacher]);
    const rawStudent = text(row[column.student]);
    if (!teacher || !rawStudent) {
      counts.跳过教师或学员为空 += 1;
      continue;
    }

    const cleaned = cleanStudentName(rawStudent);
    const missingFields = [
      column.teachingGroup < 0 ? "教研组" : "",
      column.trainingLead < 0 ? "师训组长" : "",
      column.assistantLead < 0 ? "师训助理主管/主管" : "",
    ].filter(Boolean);
    if (missingFields.length) counts.字段缺失记录 += 1;

    const target: ReminderTarget = {
      id: targets.length + 1,
      授课教师: teacher,
      教师邮箱: text(row[column.teacherEmail]),
      教研组: text(row[column.teachingGroup]),
      师训组长: text(row[column.trainingLead]),
      助理主管: text(row[column.assistantLead]),
      学员姓名: cleaned.original,
      匹配学员姓名: cleaned.cleaned || cleaned.original,
      姓名清洗说明: cleaned.note,
      新老生季度: text(row[column.studentType]),
      年级: text(row[column.grade]),
      学管: text(row[column.counselor]),
      课时: text(row[column.classHours]),
      校区: text(row[column.campus]),
      源名单行号: sourceRowNumber,
      去重合并行号: "",
    };
    seen.set(key, target);
    targets.push(target);
  }

  counts.去重后应发送数 = targets.length;
  return { targets, counts, sheetName: found.name };
}
