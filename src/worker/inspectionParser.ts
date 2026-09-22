import { headerMap, sheetCandidates } from "./excelReader";
import { emailValue, excelDate, excelTime, text } from "./utils";
import type { CellValue, DataRow } from "./types";
import type { InspectionSourceRow, RosterInfo } from "./inspectionTypes";

function firstIndex(map: Map<string, number[]>, aliases: readonly string[]) {
  for (const alias of aliases) {
    const index = map.get(alias)?.[0];
    if (index != null) return index;
  }
  return -1;
}

function findInspectionSheet(workbook: SheetJsWorkbook) {
  const required = {
    teacher: ["老师姓名", "教师姓名"],
    student: ["学员姓名", "学生姓名"],
    studentId: ["学员号", "学生号", "学员ID"],
    courseId: ["课次ID", "课程ID", "课程编码"],
    start: ["课次开始时间", "课次开始时", "课程开始时间"],
    end: ["课次结束时间", "课次结束时", "课程结束时间"],
    email: ["老师邮箱", "教师邮箱", "邮箱"],
    submitted: ["是否生成报告", "是否提交", "高中工作台提交情况"],
  } as const;
  const candidates = sheetCandidates(workbook).map((candidate) => {
    const indexes = Object.fromEntries(Object.entries(required).map(([key, aliases]) => [key, firstIndex(candidate.map, aliases)]));
    const score = Object.values(indexes).filter((index) => index >= 0).length;
    return { candidate, indexes, score };
  }).filter((item) => item.score === Object.keys(required).length).sort((left, right) => right.candidate.rows.length - left.candidate.rows.length);
  if (!candidates.length) {
    throw new Error("找不到课堂反馈明细工作表。需要老师、学员、学员号、课次ID、开始/结束时间、老师邮箱和是否生成报告字段。");
  }
  return candidates[0];
}

function rowData(headers: string[], row: CellValue[]) {
  const result: DataRow = {};
  headers.forEach((header, index) => {
    result[header || `未命名列${index + 1}`] = row[index] ?? "";
  });
  return result;
}

function sourceRows(workbook: SheetJsWorkbook): InspectionSourceRow[] {
  const found = findInspectionSheet(workbook);
  const headers = found.candidate.headers;
  const { indexes } = found;
  return found.candidate.rows.slice(1).map((row, offset) => {
    const teacherName = text(row[indexes.teacher]);
    const teacherEmail = emailValue(row[indexes.email]);
    const studentName = text(row[indexes.student]);
    const studentId = text(row[indexes.studentId]);
    const courseId = text(row[indexes.courseId]);
    const lessonStartValue = row[indexes.start];
    const lessonEndValue = row[indexes.end];
    const submittedValue = text(row[indexes.submitted]);
    const sourceRowNumber = offset + 2;
    return {
      sourceRowNumber,
      source: rowData(headers, row),
      teacherName,
      teacherEmail,
      studentName,
      studentId,
      courseId,
      lessonStart: `${excelDate(lessonStartValue)} ${excelTime(lessonStartValue)}`.trim(),
      lessonEnd: `${excelDate(lessonEndValue)} ${excelTime(lessonEndValue)}`.trim(),
      submittedValue,
      productGroup: text(row[found.candidate.map.get("产品分组")?.[0] ?? -1]),
      campus: text(row[found.candidate.map.get("校区")?.[0] ?? -1]),
      projectGroup: text(row[found.candidate.map.get("项目组")?.[0] ?? -1]),
      unsubmitted: submittedValue === "否",
      selectionKey: `${teacherEmail}|${courseId}|${sourceRowNumber}`,
    };
  }).filter((row) => row.teacherName || row.teacherEmail || row.courseId);
}

function rosterDate(fileName: string) {
  const match = fileName.match(/(20\d{2})[-_年]?(\d{2})[-_月]?(\d{2})/u);
  return match ? `${match[1]}-${match[2]}-${match[3]}` : "";
}

export function parseInspectionRows(workbook: SheetJsWorkbook) {
  const found = findInspectionSheet(workbook);
  return { rows: sourceRows(workbook), columns: found.candidate.headers };
}

export function parseInspectionRoster(workbook: SheetJsWorkbook, fileName: string): RosterInfo {
  const candidates = sheetCandidates(workbook).map((candidate) => {
    const emailIndex = firstIndex(candidate.map, ["邮箱", "教师邮箱", "老师邮箱"]);
    return { candidate, emailIndex };
  }).filter((item) => item.emailIndex >= 0).sort((left, right) => right.candidate.rows.length - left.candidate.rows.length);
  if (!candidates.length) throw new Error("找不到在职明细中的邮箱列。需要“邮箱”或“教师邮箱”字段。");
  const found = candidates[0];
  const emails = new Set<string>();
  let matchedEmailRows = 0;
  for (const row of found.candidate.rows.slice(1)) {
    const email = emailValue(row[found.emailIndex]);
    if (!email) continue;
    matchedEmailRows += 1;
    emails.add(email);
  }
  if (!emails.size) throw new Error("在职明细的邮箱列没有有效邮箱，无法进行在职教师过滤。");
  return {
    emails,
    sourceName: fileName,
    snapshotDate: rosterDate(fileName),
    rowCount: Math.max(0, found.candidate.rows.length - 1),
    matchedEmailRows,
  };
}

export async function sha256File(file: File) {
  const digest = await crypto.subtle.digest("SHA-256", await file.arrayBuffer());
  return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, "0")).join("");
}
