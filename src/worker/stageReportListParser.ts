import { headerMap, type FoundSheet } from "./excelReader";
import type { CellValue, CountMap, DataRow } from "./types";
import { cleanStudentName, emailValue, text } from "./utils";

const HEADERS = {
  teacher: ["教师姓名", "授课教师", "老师姓名"],
  email: ["邮箱", "教师邮箱", "老师邮箱"],
  studentId: ["学号", "学员号", "学生编号", "学员编号", "学生ID", "学员ID"],
  student: ["学员姓名", "学生姓名"],
  sent: ["是否发送阶段性报告"],
  appeal: ["是否申诉"],
  appealDetail: ["申诉情况详情", "申诉说明"],
} as const;

export interface StageReportTarget {
  id: number;
  original: DataRow;
  teacher: string;
  email: string;
  studentId: string;
  student: string;
  matchedStudent: string;
  cleanNote: string;
  appeal: boolean;
  sourceRowNumber: number;
  duplicateRows: string;
}

export interface StageReportListInfo {
  targets: StageReportTarget[];
  columns: string[];
  counts: CountMap;
  sheetName: string;
}

function firstIndex(map: Map<string, number[]>, aliases: readonly string[]) {
  for (const alias of aliases) {
    const index = (map.get(alias) || [])[0];
    if (index != null) return index;
  }
  return -1;
}

function hasHeader(map: Map<string, number[]>, aliases: readonly string[]) {
  return firstIndex(map, aliases) >= 0;
}

function findStageReportSheet(workbook: SheetJsWorkbook): FoundSheet {
  const seen: string[] = [];
  for (const name of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
      header: 1,
      range: 0,
      blankrows: false,
      defval: "",
    });
    if (!rows.length) continue;
    seen.push(name);
    const map = headerMap(rows[0]);
    if ([HEADERS.teacher, HEADERS.studentId, HEADERS.student].every((aliases) => hasHeader(map, aliases))) {
      return { name, rows };
    }
  }
  throw new Error(
    "找不到阶段性报告分母工作表：需要包含教师姓名、学号、学员姓名（支持常见别名）；邮箱可选。" +
      `已检查 Sheet：${seen.join("、") || "无"}`,
  );
}

function isAppealed(value: CellValue, detail: CellValue) {
  const status = text(value).toLocaleLowerCase("zh-CN");
  if (status && !["否", "no", "n", "未申诉"].includes(status)) return true;
  return Boolean(text(detail));
}

function rowKey(row: CellValue[], width: number) {
  return Array.from({ length: width }, (_, index) => text(row[index])).join("\u0000");
}

export function buildStageReportTargets(workbook: SheetJsWorkbook): StageReportListInfo {
  const found = findStageReportSheet(workbook);
  const rows = found.rows;
  const columns = rows[0].map((value, index) => text(value) || `未命名列${index + 1}`);
  const map = headerMap(rows[0]);
  const column = {
    teacher: firstIndex(map, HEADERS.teacher),
    email: firstIndex(map, HEADERS.email),
    studentId: firstIndex(map, HEADERS.studentId),
    student: firstIndex(map, HEADERS.student),
    appeal: firstIndex(map, HEADERS.appeal),
    appealDetail: firstIndex(map, HEADERS.appealDetail),
  };
  const counts: CountMap = {
    原始分母行数: Math.max(0, rows.length - 1),
    整行重复行数: 0,
    跳过空行数: 0,
    教师或学员为空: 0,
    邮箱为空: 0,
    申诉数: 0,
    去重后应检查数: 0,
  };
  const targets: StageReportTarget[] = [];
  const seen = new Map<string, StageReportTarget>();

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const sourceRowNumber = rowIndex + 1;
    if (!row.some((value) => text(value))) {
      counts.跳过空行数 += 1;
      continue;
    }
    const key = rowKey(row, columns.length);
    const existed = seen.get(key);
    if (existed) {
      counts.整行重复行数 += 1;
      existed.duplicateRows = [existed.duplicateRows, sourceRowNumber].filter(Boolean).join("、");
      continue;
    }
    const teacher = text(row[column.teacher]);
    const student = text(row[column.student]);
    const email = emailValue(row[column.email]);
    if (!teacher || !student) counts.教师或学员为空 += 1;
    if (!email) counts.邮箱为空 += 1;
    const cleaned = cleanStudentName(student);
    const appealed = isAppealed(row[column.appeal], row[column.appealDetail]);
    if (appealed) counts.申诉数 += 1;
    const original = columns.reduce<DataRow>((result, name, index) => {
      result[name] = row[index] ?? "";
      return result;
    }, {});
    const target: StageReportTarget = {
      id: targets.length + 1,
      original,
      teacher,
      email,
      studentId: text(row[column.studentId]),
      student,
      matchedStudent: cleaned.cleaned || student.replace(/\s+/g, ""),
      cleanNote: cleaned.note,
      appeal: appealed,
      sourceRowNumber,
      duplicateRows: "",
    };
    seen.set(key, target);
    targets.push(target);
  }
  counts.去重后应检查数 = targets.length;
  return { targets, columns, counts, sheetName: found.name };
}
