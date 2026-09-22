import type { CellValue, DataRow, SheetDefinition } from "./types";

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

export type InspectionBatchKind = "formal" | "trial";

export interface RosterInfo {
  emails: Set<string>;
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
  selectedTeachers: number;
  unsubmittedRows: number;
  unsubmittedSelectedRows: number;
  excludedRows: number;
  excludedNoEmailRows: number;
  excludedNotInRosterRows: number;
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
  batchKind: InspectionBatchKind;
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

export interface InspectionHistoryPayload {
  batch: {
    businessWeekStart: string;
    businessWeekEnd: string;
    sourceName: string;
    sourceSha256: string;
    rosterName: string;
    rosterSha256: string;
    rosterSnapshotDate: string;
    sampleLimit: number;
    eligibleCount: number;
    selectedCount: number;
    teacherCount: number;
    unsubmittedCount: number;
    ruleVersion: string;
    attempt: number;
    batchKind: InspectionBatchKind;
  };
  items: InspectionHistoryItem[];
}

export interface InspectionOutput {
  chunks: Uint8Array[];
  filename: string;
  historyPayload: InspectionHistoryPayload;
  summary: InspectionStats;
}

export type TeachingServiceScoreStatus = "matched" | "ambiguous" | "missing_score";

export interface TeachingServiceTeacher {
  teacherName: string;
  teacherEmail: string;
  researchGroup: string;
  trainingLeader: string;
  trainingSupervisor: string;
  campus: string;
  score: number | null;
  scoreStatus: TeachingServiceScoreStatus;
  scoreTeacherName: string;
  scoreResearchGroup: string;
  scoreTrainingLeader: string;
  scoreCampus: string;
  suggestedMonthlyCount: number;
}

export interface TeachingServiceSnapshot {
  schemaVersion: number;
  snapshotDate: string;
  scorePeriod: string;
  scoreSourceFile: string;
  scoreSourceSha256: string;
  candidateSourceFile: string;
  candidateSourceSha256: string;
  rosterSourceFile: string;
  rosterSourceSha256: string;
  generatedAt: string;
  scoreRule: string;
  matchRule: string;
  matchSummary: {
    candidateTeachers: number;
    matched: number;
    ambiguous: number;
    missingScore: number;
  };
  teachers: TeachingServiceTeacher[];
}

export interface MonthlyInspectionItem extends InspectionHistoryItem {
  batchId: string;
  businessWeekStart: string;
  businessWeekEnd: string;
  attempt: number;
}

export interface MonthlyInspectionData {
  month: string;
  items: MonthlyInspectionItem[];
}

export interface MonthlyTeacherPlan extends TeachingServiceTeacher {
  recommendedCount: number;
  assignedCount: number;
  actualCount: number;
  unsubmittedCount: number;
  remainingCount: number;
  lastInspectionWeek: string;
  focusReason: string;
  items: MonthlyInspectionItem[];
}

export interface MonthlyInspectionPlan {
  month: string;
  availableSlots: number;
  candidateTeachers: number;
  assignedSlots: number;
  coverageShortfall: number;
  recommendedSlots: number;
  frequencyShortfall: number;
  extraSlots: number;
  lowScoreTeachers: number;
  missingScoreTeachers: number;
  teachers: MonthlyTeacherPlan[];
}

export interface InspectionSheetDefinition extends SheetDefinition {
  rows: DataRow[];
}

export type InspectionCellValue = CellValue;
