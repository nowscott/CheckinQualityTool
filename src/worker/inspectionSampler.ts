import { excelDate, text } from "./utils";
import { displayTeacherName } from "../lib/teacherDisplay.js";
import { INSPECTION_RULE_VERSION } from "./inspectionTypes";
import type {
  InspectionCandidateRow,
  InspectionHistoryItem,
  InspectionRiskRow,
  InspectionSelectedRow,
  InspectionSelection,
  InspectionSourceRow,
  InspectionPriorityMode,
  RosterInfo,
} from "./inspectionTypes";

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

function normalizeTeacherName(value: string) {
  return value.normalize("NFKC").replace(/\s+/gu, "").trim().toLocaleLowerCase();
}

function teacherNameKeys(row: InspectionSourceRow) {
  return new Set([
    normalizeTeacherName(row.teacherName),
    normalizeTeacherName(displayTeacherName(row.teacherName, row.teacherEmail)),
  ].filter(Boolean));
}

function selectionReason(
  row: InspectionSourceRow,
  coverage: boolean,
  fill: boolean,
  focus: boolean,
  focusReason: string,
) {
  const reasons: string[] = [];
  if (row.unsubmitted) reasons.push("报告未生成容量外加抽");
  if (focus) reasons.push(focusReason);
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
    priorityMode?: InspectionPriorityMode;
    focusTeacherNames?: string[];
  },
): InspectionSelection {
  const sampleCount = Math.max(0, Math.floor(options.sampleCount));
  const seed = `${options.sourceSha256}|${options.rosterSha256}|${options.attempt}`;
  const sourceRows = rows.map((row) => ({
    ...row,
    selectionKey: hashScore(`${seed}|${row.teacherEmail}|${row.courseId}|${row.sourceRowNumber}`),
  }));
  const roleExcludedEmails = roster.roleExcludedEmails || new Set<string>();
  const inRosterRows = sourceRows.filter((row) => row.teacherEmail && roster.emails.has(row.teacherEmail));
  const excludedManagementRows = inRosterRows.filter((row) => roleExcludedEmails.has(row.teacherEmail));
  const activeRows = inRosterRows.filter((row) => !roleExcludedEmails.has(row.teacherEmail));
  const excludedNoEmailRows = sourceRows.filter((row) => !row.teacherEmail).length;
  const excludedNotInRosterRows = sourceRows.filter((row) =>
    Boolean(row.teacherEmail) && !roster.emails.has(row.teacherEmail),
  ).length;
  const normalRows = activeRows.filter((row) => !row.unsubmitted);

  const groups = new Map<string, InspectionSourceRow[]>();
  for (const row of normalRows) {
    const key = uniqueTeacherKey(row);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }
  const teacherGroups = [...groups.entries()].map(([key, teacherRows]) => ({
    key,
    rows: teacherRows,
    sortKey: [...teacherRows].sort((left, right) => left.selectionKey.localeCompare(right.selectionKey))[0]?.selectionKey || "",
  })).sort((left, right) => left.sortKey.localeCompare(right.sortKey) || left.key.localeCompare(right.key));

  const focusNameKeys = new Set((options.focusTeacherNames || []).map(normalizeTeacherName).filter(Boolean));
  const groupsByName = new Map<string, Set<string>>();
  for (const row of activeRows) {
    const teacherKey = uniqueTeacherKey(row);
    for (const nameKey of teacherNameKeys(row)) {
      if (!groupsByName.has(nameKey)) groupsByName.set(nameKey, new Set());
      groupsByName.get(nameKey)!.add(teacherKey);
    }
  }
  const normalGroupKeys = new Set(teacherGroups.map((group) => group.key));
  const focusedGroupKeys = new Set<string>();
  const matchedFocusTeacherKeys = new Set<string>();
  let matchedFocusNameCount = 0;
  let ambiguousFocusTeacherCount = 0;
  for (const focusName of focusNameKeys) {
    const candidates = groupsByName.get(focusName) || new Set<string>();
    if (candidates.size === 1) {
      matchedFocusNameCount += 1;
      const teacherKey = [...candidates][0];
      matchedFocusTeacherKeys.add(teacherKey);
      if (normalGroupKeys.has(teacherKey)) focusedGroupKeys.add(teacherKey);
    } else if (candidates.size > 1) ambiguousFocusTeacherCount += 1;
  }
  const priorityMode = options.priorityMode || "coverage";
  const priority: InspectionSelection["priority"] = {
    mode: priorityMode,
    focusTeacherCount: focusNameKeys.size,
    matchedFocusTeacherCount: matchedFocusTeacherKeys.size,
    unmatchedFocusTeacherCount: Math.max(0, focusNameKeys.size - matchedFocusNameCount - ambiguousFocusTeacherCount),
    ambiguousFocusTeacherCount,
  };
  const focusReason = priorityMode === "coverage"
    ? "未反馈教师剩余名额加频"
    : "本月未反馈教师优先";
  const orderedGroups = priorityMode === "unreported"
    ? [
      ...teacherGroups.filter((group) => focusedGroupKeys.has(group.key)),
      ...teacherGroups.filter((group) => !focusedGroupKeys.has(group.key)),
    ]
    : teacherGroups;

  const selectedNormal = new Map<string, InspectionSelectedRow>();
  const selectedSourceRows = new Set<string>();
  let focusTeacherExtraRows = 0;
  const rowKey = (row: InspectionSourceRow) => `${row.teacherEmail}\u0000${row.courseId}\u0000${row.sourceRowNumber}`;
  const add = (row: InspectionSourceRow, coverage: boolean, fill: boolean, focus: boolean) => {
    if (selectedNormal.size >= sampleCount) return false;
    const key = rowKey(row);
    if (selectedSourceRows.has(key)) return false;
    selectedSourceRows.add(key);
    selectedNormal.set(key, {
      ...row,
      selectionOrder: selectedNormal.size + 1,
      selectionReason: selectionReason(row, coverage, fill, focus, focusReason),
    });
    if (fill && focus) focusTeacherExtraRows += 1;
    return true;
  };

  for (const group of orderedGroups) {
    if (selectedNormal.size >= sampleCount) break;
    add(chooseRow(group.rows), true, false, priorityMode === "unreported" && focusedGroupKeys.has(group.key));
  }

  const remaining = normalRows
    .filter((row) => !selectedSourceRows.has(rowKey(row)))
    .sort((left, right) => {
      const leftFocused = focusedGroupKeys.has(uniqueTeacherKey(left));
      const rightFocused = focusedGroupKeys.has(uniqueTeacherKey(right));
      return Number(rightFocused) - Number(leftFocused) || compareRows(left, right);
    });
  for (const row of remaining) {
    if (selectedNormal.size >= sampleCount) break;
    add(row, false, true, focusedGroupKeys.has(uniqueTeacherKey(row)));
  }

  const selectedExtra: InspectionSelectedRow[] = activeRows
    .filter((row) => row.unsubmitted)
    .sort(compareDisplayRows)
    .map((row, index) => ({
      ...row,
      selectionOrder: selectedNormal.size + index + 1,
      selectionReason: selectionReason(
        row,
        false,
        false,
        priorityMode === "unreported" && focusedGroupKeys.has(uniqueTeacherKey(row)),
        focusReason,
      ),
    }));
  const unorderedRows = [...selectedNormal.values(), ...selectedExtra];
  const teacherOrder = new Map<string, number>();
  for (const row of unorderedRows) {
    const teacherKey = uniqueTeacherKey(row);
    if (!teacherOrder.has(teacherKey)) teacherOrder.set(teacherKey, teacherOrder.size);
  }
  const selectedRows: InspectionSelectedRow[] = unorderedRows
    .sort((left, right) =>
      (teacherOrder.get(uniqueTeacherKey(left)) ?? Number.MAX_SAFE_INTEGER)
        - (teacherOrder.get(uniqueTeacherKey(right)) ?? Number.MAX_SAFE_INTEGER)
      || compareDisplayRows(left, right),
    )
    .map((row, index) => ({ ...row, selectionOrder: index + 1 }));
  const selectedOrders = new Map(selectedRows.map((row) => [rowKey(row), row]));
  const finalizedRiskRows: InspectionRiskRow[] = selectedExtra.map((row) => {
    const picked = selectedOrders.get(rowKey(row));
    return {
      ...row,
      inspected: Boolean(picked),
      inspectionOrder: picked?.selectionOrder ?? "",
      inspectionReason: picked?.selectionReason || "报告未生成容量外加抽",
    };
  });

  const dates = activeRows.map((row) => excelDate(row.lessonStart)).filter(Boolean).sort();
  const week = businessWeek(dates[0] || "");
  const selectedTeachers = new Set(selectedRows.map(uniqueTeacherKey));
  const eligibleTeachers = new Set(activeRows.map(uniqueTeacherKey));
  const excludedManagementTeachers = new Set(excludedManagementRows.map(uniqueTeacherKey));
  return {
    selectedRows,
    riskRows: finalizedRiskRows,
    allEligibleRows: activeRows.map((row) => ({ ...row, active: true, excludedReason: "" })),
    stats: {
      sourceRows: rows.length,
      eligibleRows: activeRows.length,
      eligibleTeachers: eligibleTeachers.size,
      selectedRows: selectedRows.length,
      normalSelectedRows: selectedNormal.size,
      extraSelectedRows: finalizedRiskRows.length,
      selectedTeachers: selectedTeachers.size,
      focusTeacherExtraRows,
      unsubmittedRows: finalizedRiskRows.length,
      unsubmittedSelectedRows: finalizedRiskRows.filter((row) => row.inspected).length,
      excludedRows: sourceRows.length - activeRows.length,
      excludedNoEmailRows,
      excludedNotInRosterRows,
      excludedManagementRows: excludedManagementRows.length,
      excludedManagementTeachers: excludedManagementTeachers.size,
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
    priority,
  };
}

export function normalizeInspectionNumber(value: unknown) {
  const number = Number(text(value));
  if (!Number.isFinite(number) || number < 0) return 0;
  return Math.floor(number);
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
