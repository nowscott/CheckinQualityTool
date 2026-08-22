import { headerMap, sheetCandidates, type FoundSheet } from "./excelReader";
import { trainingPersonKey } from "./stageReportGroups";
import type { CellValue, CountMap, DataRow } from "./types";
import { cleanStudentName, emailValue, text } from "./utils";

const HEADERS = {
  teacher: ["教师姓名", "授课教师", "老师姓名"],
  email: ["邮箱", "教师邮箱", "老师邮箱"],
  studentId: ["学号", "学员号", "学生编号", "学员编号", "学生ID", "学员ID"],
  student: ["学员姓名", "学生姓名"],
  teachingGroup: ["教研组", "教研组名称"],
  sent: ["是否发送阶段性报告"],
  stageNeed: ["是否需要发送"],
  appeal: ["是否申诉"],
  appealDetail: ["申诉情况详情", "申诉说明"],
  windowReport: ["是否发送窗口期报告", "窗口期报告是否发送"],
  courseEnd: ["暑假最后一节课时间", "最后一节课时间", "结课时间"],
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
  trainingLeadGroups: Map<string, string[]>;
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

function scoreStageReportSheet(name: string, headers: string[], map: Map<string, number[]>) {
  const hasStageHeader = headers.some((header) => header.includes("阶段性报告"));
  const hasWindowHeader = HEADERS.windowReport.some((alias) => headers.includes(alias)) || headers.some((header) => header.includes("窗口期报告"));
  let score = 0;
  if (hasHeader(map, HEADERS.teachingGroup)) score += 8;
  if (hasHeader(map, HEADERS.sent)) score += 8;
  if (hasHeader(map, HEADERS.stageNeed)) score += 4;
  if (hasStageHeader) score += 2;
  if (hasHeader(map, HEADERS.courseEnd)) score += 2;
  if (hasWindowHeader) score -= 8;
  if (/(阶段性报告|非窗口期|暑期在读|结课)/u.test(name)) score += 3;
  if (/(窗口期报告|汇总|维度)/u.test(name)) score -= 3;
  return score;
}

function findStageReportSheet(workbook: SheetJsWorkbook): FoundSheet {
  const candidates: Array<{ found: FoundSheet; score: number; hasGroup: boolean; order: number }> = [];
  for (const candidate of sheetCandidates(workbook)) {
    const { name, rows, map } = candidate;
    if ([HEADERS.teacher, HEADERS.studentId, HEADERS.student].every((aliases) => hasHeader(map, aliases))) {
      const headers = rows[0].map(text);
      candidates.push({
        found: { name, rows },
        score: scoreStageReportSheet(name, headers, map),
        hasGroup: hasHeader(map, HEADERS.teachingGroup),
        order: candidates.length,
      });
    }
  }
  if (!candidates.length) {
    throw new Error("找不到阶段性报告分母工作表：需要包含教师姓名、学号、学员姓名（支持常见别名）；邮箱可选。");
  }
  const preferred = candidates.some((candidate) => candidate.hasGroup)
    ? candidates.filter((candidate) => candidate.hasGroup)
    : candidates;
  preferred.sort((a, b) => b.score - a.score || a.order - b.order);
  return preferred[0].found;
}

function addLeadGroup(groups: Map<string, Set<string>>, lead: CellValue, group: CellValue) {
  const leadKey = trainingPersonKey(lead);
  const groupName = text(group);
  if (!leadKey || !groupName) return;
  if (!groups.has(leadKey)) groups.set(leadKey, new Set<string>());
  groups.get(leadKey)!.add(groupName);
}

function findFallbackTrainingLeadGroups(workbook: SheetJsWorkbook): Map<string, string[]> {
  const ownGroups = new Map<string, Set<string>>();
  const summaryGroups = new Map<string, Set<string>>();
  for (const candidate of sheetCandidates(workbook)) {
    const { rows, map, headers } = candidate;
    const leadIndex = firstIndex(map, ["师训组长"]);
    const groupIndex = firstIndex(map, HEADERS.teachingGroup);
    if (leadIndex < 0 || groupIndex < 0) continue;

    const teacherIndex = firstIndex(map, HEADERS.teacher);
    const isLeadSummary = teacherIndex < 0 && headers.some((header) => /报告需发送|报告已发送|报告发送率/u.test(header));
    if (isLeadSummary) {
      for (const row of rows.slice(1)) addLeadGroup(summaryGroups, row[leadIndex], row[groupIndex]);
      continue;
    }

    if (teacherIndex < 0) continue;
    for (const row of rows.slice(1)) {
      if (trainingPersonKey(row[teacherIndex]) === trainingPersonKey(row[leadIndex])) {
        addLeadGroup(ownGroups, row[leadIndex], row[groupIndex]);
      }
    }
  }
  const resolved = new Map<string, Set<string>>(ownGroups);
  for (const [lead, values] of summaryGroups) {
    if (!resolved.has(lead)) resolved.set(lead, values);
  }
  return new Map([...resolved.entries()].map(([lead, values]) => [lead, [...values].sort((a, b) => a.localeCompare(b, "zh-CN"))]));
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
  return {
    targets,
    columns,
    counts,
    sheetName: found.name,
    trainingLeadGroups: findFallbackTrainingLeadGroups(workbook),
  };
}
