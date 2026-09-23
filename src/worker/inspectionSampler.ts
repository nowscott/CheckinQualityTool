import { excelDate, text } from "./utils";
import { displayTeacherName } from "../lib/teacherDisplay.js";
import type {
  InspectionCandidateRow,
  InspectionHistoryItem,
  InspectionRiskRow,
  InspectionSelectedRow,
  InspectionSelection,
  InspectionSourceRow,
  InspectionBatchKind,
  RosterInfo,
} from "./inspectionTypes";

export const INSPECTION_RULE_VERSION = "inspection-v3-role-risk-only";

function hashString(value: string) {
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

function hashScore(value: string) {
  return hashString(value).toString(16).padStart(8, "0");
}

function compareRows(left: InspectionSourceRow, right: InspectionSourceRow) {
  return (
    Number(right.unsubmitted) - Number(left.unsubmitted) ||
    left.selectionKey.localeCompare(right.selectionKey) ||
    left.sourceRowNumber - right.sourceRowNumber
  );
}

function uniqueTeacherKey(row: InspectionSourceRow) {
  return row.teacherEmail || `${row.teacherName}\u0000${row.sourceRowNumber}`;
}

function businessWeek(value: string) {
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.valueOf())) throw new Error(`无法识别课次日期：${value || "空值"}。`);
  const day = date.getDay() || 7;
  const start = new Date(date);
  start.setDate(date.getDate() - day + 1);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  const format = (item: Date) => {
    const pad = (number: number) => String(number).padStart(2, "0");
    return `${item.getFullYear()}-${pad(item.getMonth() + 1)}-${pad(item.getDate())}`;
  };
  return { start: format(start), end: format(end) };
}

function chooseRow(rows: InspectionSourceRow[]) {
  return [...rows].sort(compareRows)[0];
}

function compareDisplayRows(left: InspectionSourceRow, right: InspectionSourceRow) {
  return (
    (left.teacherEmail || left.teacherName).localeCompare(right.teacherEmail || right.teacherName) ||
    left.teacherName.localeCompare(right.teacherName) ||
    left.lessonStart.localeCompare(right.lessonStart) ||
    left.lessonEnd.localeCompare(right.lessonEnd) ||
    left.sourceRowNumber - right.sourceRowNumber
  );
}

function selectionReason(row: InspectionSourceRow, coverage: boolean, fill: boolean) {
  const reasons: string[] = [];
  if (row.unsubmitted) reasons.push("高中工作台未生成报告");
  if (coverage) reasons.push("教师覆盖");
  if (fill) reasons.push("补足抽检数");
  return reasons.join("；") || "稳定抽检排序";
}

export function buildInspectionSelection(
  rows: InspectionSourceRow[],
  roster: RosterInfo,
  options: {
    sampleCount: number;
    attempt: number;
    sourceSha256: string;
    rosterSha256: string;
    sourceName: string;
    sourceColumns?: string[];
    batchKind?: InspectionBatchKind;
  },
): InspectionSelection {
  const sampleCount = Math.max(0, Math.floor(options.sampleCount));
  const seed = `${options.sourceSha256}|${options.rosterSha256}|${options.attempt}`;
  const sourceRows = rows.map((row) => ({
    ...row,
    selectionKey: hashScore(`${seed}|${row.teacherEmail}|${row.courseId}|${row.sourceRowNumber}`),
  }));
  const roleExcludedEmails = roster.roleExcludedEmails || new Set<string>();
  const activeRows = sourceRows.filter((row) => row.teacherEmail && roster.emails.has(row.teacherEmail));
  const excludedNoEmailRows = sourceRows.filter((row) => !row.teacherEmail).length;
  const excludedNotInRosterRows = sourceRows.filter((row) =>
    Boolean(row.teacherEmail) && !roster.emails.has(row.teacherEmail),
  ).length;
  const excludedRoleRows = activeRows.filter((row) => row.unsubmitted && roleExcludedEmails.has(row.teacherEmail)).length;

  const groups = new Map<string, InspectionSourceRow[]>();
  for (const row of activeRows) {
    const key = uniqueTeacherKey(row);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }
  const teacherGroups = [...groups.entries()].map(([key, teacherRows]) => ({
    key,
    rows: teacherRows,
    priority: teacherRows.some((row) => row.unsubmitted) ? 0 : 1,
    sortKey: [...teacherRows].sort((left, right) => left.selectionKey.localeCompare(right.selectionKey))[0]?.selectionKey || "",
  })).sort((left, right) => left.priority - right.priority || left.sortKey.localeCompare(right.sortKey) || left.key.localeCompare(right.key));

  const selected = new Map<string, InspectionSelectedRow>();
  const selectedSourceRows = new Set<string>();
  const rowKey = (row: InspectionSourceRow) => `${row.teacherEmail}\u0000${row.courseId}\u0000${row.sourceRowNumber}`;
  const add = (row: InspectionSourceRow, coverage: boolean, fill: boolean) => {
    if (selected.size >= sampleCount) return false;
    const key = rowKey(row);
    if (selectedSourceRows.has(key)) return false;
    selectedSourceRows.add(key);
    selected.set(key, {
      ...row,
      selectionOrder: selected.size + 1,
      selectionReason: selectionReason(row, coverage, fill),
    });
    return true;
  };

  for (const group of teacherGroups) {
    if (selected.size >= sampleCount) break;
    if (sampleCount >= teacherGroups.length || selected.size < sampleCount) {
      add(chooseRow(group.rows), true, false);
    }
  }

  const remaining = activeRows
    .filter((row) => !selectedSourceRows.has(rowKey(row)))
    .sort(compareRows);
  for (const row of remaining) {
    if (selected.size >= sampleCount) break;
    add(row, false, true);
  }

  const selectedRows = [...selected.values()]
    .sort(compareDisplayRows)
    .map((row, index) => ({ ...row, selectionOrder: index + 1 }));
  const selectedOrders = new Map(selectedRows.map((row) => [rowKey(row), row]));
  const riskRows: InspectionRiskRow[] = activeRows
    .filter((row) => row.unsubmitted && !roleExcludedEmails.has(row.teacherEmail))
    .sort(compareDisplayRows)
    .map((row) => {
      const picked = selectedOrders.get(rowKey(row));
      return {
        ...row,
        inspected: Boolean(picked),
        inspectionOrder: picked?.selectionOrder ?? "",
        inspectionReason: picked?.selectionReason || "未进入抽检上限，列入风险清单",
      };
    });

  const dates = activeRows.map((row) => excelDate(row.lessonStart)).filter(Boolean).sort();
  const week = businessWeek(dates[0] || "");
  const selectedTeachers = new Set(selectedRows.map(uniqueTeacherKey));
  return {
    selectedRows,
    riskRows,
    allEligibleRows: activeRows.map((row) => ({ ...row, active: true, excludedReason: "" })),
    stats: {
      sourceRows: rows.length,
      eligibleRows: activeRows.length,
      eligibleTeachers: teacherGroups.length,
      selectedRows: selectedRows.length,
      selectedTeachers: selectedTeachers.size,
      unsubmittedRows: riskRows.length,
      unsubmittedSelectedRows: selectedRows.filter((row) => row.unsubmitted && !roleExcludedEmails.has(row.teacherEmail)).length,
      excludedRows: sourceRows.length - activeRows.length,
      excludedNoEmailRows,
      excludedNotInRosterRows,
      excludedRoleRows,
      unknownSubmissionRows: activeRows.filter((row) => !["是", "否"].includes(row.submittedValue)).length,
    },
    sourceName: options.sourceName,
    roster,
    businessWeekStart: week.start,
    businessWeekEnd: week.end,
    attempt: options.attempt,
    ruleVersion: INSPECTION_RULE_VERSION,
    sourceSha256: options.sourceSha256,
    rosterSha256: options.rosterSha256,
    sourceColumns: options.sourceColumns || [],
    sampleLimit: sampleCount,
    batchKind: options.batchKind || "formal",
  };
}

export function historyItems(rows: InspectionSelectedRow[]): InspectionHistoryItem[] {
  return rows.map((row) => ({
    position: row.selectionOrder,
    teacherName: displayTeacherName(row.teacherName, row.teacherEmail),
    teacherEmail: row.teacherEmail,
    studentName: row.studentName,
    studentId: row.studentId,
    courseId: row.courseId,
    lessonStart: row.lessonStart,
    lessonEnd: row.lessonEnd,
    submittedValue: row.submittedValue,
    productGroup: row.productGroup,
    campus: row.campus,
    projectGroup: row.projectGroup,
    selectionReason: row.selectionReason,
  }));
}

export function normalizeInspectionNumber(value: unknown) {
  const number = Number(text(value));
  if (!Number.isFinite(number) || number < 0) return 0;
  return Math.floor(number);
}
