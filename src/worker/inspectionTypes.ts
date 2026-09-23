import type { CellValue, DataRow, SheetDefinition } from "./types";

export const INSPECTION_RULE_VERSION = "inspection-v10-unique-score-name-match";

export type InspectionPriorityMode = "coverage" | "unreported";

export interface InspectionPrioritySummary {
  mode: InspectionPriorityMode;
  focusTeacherCount: number;
  matchedFocusTeacherCount: number;
  unmatchedFocusTeacherCount: number;
  ambiguousFocusTeacherCount: number;
  scoreSourceTeacherCount: number;
  matchedScoreTeacherCount: number;
  missingScoreTeacherCount: number;
  ambiguousScoreTeacherCount: number;
}

export interface InspectionTeacherScore {
  teacherName: string;
  researchGroup: string;
  trainingLeader: string;
  priorityRank: number;
}

export interface InspectionSourceRow {
  sourceRowNumber: number;
  source: DataRow;
  teacherName: string;
  teacherEmail: string;
  studentName: string;
  studentId: string;
  courseId: string;
  lessonStart: string;
  lessonEnd: string;
  submittedValue: string;
  productGroup: string;
  campus: string;
  projectGroup: string;
  unsubmitted: boolean;
  selectionKey: string;
}

export interface RosterInfo {
  emails: Set<string>;
  roleExcludedEmails: Set<string>;
  sourceName: string;
  snapshotDate: string;
  rowCount: number;
  matchedEmailRows: number;
}

export interface DefaultRosterAsset {
  snapshotDate: string;
  sourceFile: string;
  sourceSha256: string;
  rowCount: number;
  matchedEmailRows: number;
  emails: string[];
  roleExcludedEmails?: string[];
}

export interface DefaultRosterRoleExclusionAsset {
  snapshotDate: string;
  sourceFile: string;
  emails: string[];
}

export interface InspectionCandidateRow extends InspectionSourceRow {
  active: boolean;
  excludedReason: string;
}

export interface InspectionSelectedRow extends InspectionSourceRow {
  selectionOrder: number;
  selectionReason: string;
}

export interface InspectionRiskRow extends InspectionSourceRow {
  inspected: boolean;
  inspectionOrder: number | "";
  inspectionReason: string;
}

export interface InspectionStats {
  sourceRows: number;
  eligibleRows: number;
  eligibleTeachers: number;
  selectedRows: number;
  normalSelectedRows: number;
  extraSelectedRows: number;
  selectedTeachers: number;
  focusTeacherExtraRows: number;
  unsubmittedRows: number;
  unsubmittedSelectedRows: number;
  excludedRows: number;
  excludedNoEmailRows: number;
  excludedNotInRosterRows: number;
  excludedManagementRows: number;
  excludedManagementTeachers: number;
  unknownSubmissionRows: number;
}

export interface InspectionSelection {
  selectedRows: InspectionSelectedRow[];
  riskRows: InspectionRiskRow[];
  allEligibleRows: InspectionCandidateRow[];
  stats: InspectionStats;
  sourceName: string;
  roster: RosterInfo;
  businessWeekStart: string;
  businessWeekEnd: string;
  attempt: number;
  ruleVersion: string;
  sourceSha256: string;
  rosterSha256: string;
  sourceColumns: string[];
  sampleLimit: number;
  priority: InspectionPrioritySummary;
}

export interface InspectionHistoryItem {
  position: number;
  teacherName: string;
  teacherEmail: string;
  studentName: string;
  studentId: string;
  courseId: string;
  lessonStart: string;
  lessonEnd: string;
  submittedValue: string;
  productGroup: string;
  campus: string;
  projectGroup: string;
  selectionReason: string;
}

export interface InspectionOutput {
  chunks: Uint8Array[];
  filename: string;
  summary: InspectionStats;
  priority: InspectionPrioritySummary;
}

export interface InspectionSheetDefinition extends SheetDefinition {
  rows: DataRow[];
}

export type InspectionCellValue = CellValue;
