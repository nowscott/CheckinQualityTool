import { findSheet, headerMap } from "./excelReader";
import type { ListInfo, TargetRow, Whitelist } from "./types";
import { cleanStudentName, emailValue, sortDate, text, weekOfMonth } from "./utils";
import { findPreCleanWhitelistEntry } from "./whitelist";

export function buildTargets(workbook: SheetJsWorkbook, whitelist?: Whitelist): ListInfo {
  const found = findSheet(workbook, ["老师姓名", "学员姓名", "老师邮箱", "课次开始时"]);
  const rows = found.rows;
  const map = headerMap(rows[0]);
  const index = (name: string, occurrence = 0) => (map.get(name) || [])[occurrence] ?? -1;
  const columns = {
    teacher: index("老师姓名"),
    student: index("学员姓名"),
    studentId: index("学员号"),
    lessonDate: index("课次开始时"),
    start: index("间", 0),
    end: index("间", 1),
    campus: index("校区"),
    project: index("项目组"),
    subject: index("科目"),
    email: index("老师邮箱"),
  };
  const grouped = new Map<string, Array<Omit<TargetRow, "该周课次数">>>();
  const weekCounts = new Map<number, number>();
  const counts = {
    原始课次行数: Math.max(0, rows.length - 1),
    跳过教师或学员为空: 0,
    名单教师邮箱为空: 0,
    姓名已清洗课次: 0,
    姓名不足两字课次: 0,
    去重后质检人数: 0,
    合并的重复课次: 0,
  };

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const teacher = text(row[columns.teacher]);
    const rawStudentName = text(row[columns.student]);
    const studentId = text(row[columns.studentId]);
    const whitelistEntry = whitelist ? findPreCleanWhitelistEntry(studentId, rawStudentName, whitelist) : null;
    const preserveOriginalName = whitelistEntry?.处理方式 === "保留原名";
    const studentName = cleanStudentName(rawStudentName);
    const student = preserveOriginalName ? studentName.original.replace(/\s+/g, "") : studentName.cleaned;
    const studentNote =
      preserveOriginalName && student !== studentName.cleaned
        ? `${studentName.original}（白名单保留原名）`
        : studentName.note;
    const teacherEmail = emailValue(row[columns.email]);
    if (!teacher || !studentName.original) {
      counts.跳过教师或学员为空 += 1;
      continue;
    }
    if (studentNote) counts.姓名已清洗课次 += 1;
    if ([...student].length < 2) counts.姓名不足两字课次 += 1;
    if (!teacherEmail) counts.名单教师邮箱为空 += 1;
    const lessonWeek = weekOfMonth(row[columns.lessonDate]);
    if (lessonWeek) weekCounts.set(lessonWeek, (weekCounts.get(lessonWeek) || 0) + 1);
    const key = `${teacherEmail || `__name__:${teacher}`}\u0000${student || `__raw__:${studentName.original}`}`;
    const record: Omit<TargetRow, "该周课次数"> = {
      教师姓名: teacher,
      教师邮箱: teacherEmail,
      学员姓名: student,
      原始学员姓名: studentName.original,
      姓名清洗说明: studentNote,
      学员号: studentId,
      上课日期: row[columns.lessonDate],
      上课开始: text(row[columns.start]),
      上课结束: text(row[columns.end]),
      校区: text(row[columns.campus]),
      项目组: text(row[columns.project]),
      科目: text(row[columns.subject]),
      源名单行号: rowIndex + 1,
    };
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key)!.push(record);
  }

  const targets: TargetRow[] = [];
  for (const records of grouped.values()) {
    records.sort(
      (a, b) =>
        sortDate(a.上课日期) - sortDate(b.上课日期) ||
        a.上课开始.localeCompare(b.上课开始) ||
        a.源名单行号 - b.源名单行号,
    );
    targets.push({ ...records[0], 该周课次数: records.length });
  }
  targets.sort(
    (a, b) =>
      a.教师姓名.localeCompare(b.教师姓名, "zh-CN") ||
      a.学员姓名.localeCompare(b.学员姓名, "zh-CN"),
  );
  counts.去重后质检人数 = targets.length;
  counts.合并的重复课次 =
    counts.原始课次行数 - targets.length - counts.跳过教师或学员为空;
  return { targets, counts, weekCounts, sheetName: found.name };
}
