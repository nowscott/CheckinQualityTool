export type WeekLabel = "auto" | "第一周" | "第二周" | "第三周" | "第四周" | "第五周";

export type StatusMode = "working" | "done" | "error";

import type { InspectionPriorityMode, InspectionPrioritySummary, InspectionStats, InspectionTeacherScore } from "../worker/inspectionTypes";
import type { InspectionRosterSnapshot, RosterSnapshotMeta } from "../lib/rosterSnapshot";

export interface ProcessingStatus {
  visible: boolean;
  title: string;
  message: string;
  progress: number;
  mode: StatusMode;
}

export interface ResultSummary {
  targets: number;
  sent: number;
  unsent: number;
  exempt?: number;
  cleanChats: number;
}

export interface InspectionRequest {
  type: "process";
  mode: "inspection";
  feedbackFile: File;
  rosterFile?: File;
  defaultRoster?: InspectionRosterSnapshot;
  sampleCount: number;
  attempt: number;
  includeExplanation: boolean;
  includeReason: boolean;
  priorityMode: InspectionPriorityMode;
  focusTeacherNames: string[];
  coverageMonth?: string;
  inspectedTeachers?: { teacherEmail: string; teacherName: string }[];
  teacherScoreRows: InspectionTeacherScore[];
  /** Optional direct score map, reserved for deterministic internal fixtures. */
  teacherScoresByEmail?: Record<string, number>;
}

export interface InspectionWorkerComplete {
  type: "inspectionComplete";
  chunks: Uint8Array[];
  byteLength: number;
  filename: string;
  summary: InspectionStats;
  priority: InspectionPrioritySummary;
  rosterSnapshot: RosterSnapshotMeta;
}

export type WorkerResponse =
  | {
      type: "progress";
      title: string;
      message: string;
      progress: number;
    }
  | {
      type: "complete";
      chunks: Uint8Array[];
      byteLength: number;
      filename: string;
      summary: ResultSummary;
    }
  | {
      type: "error";
      message: string;
    }
  | InspectionWorkerComplete;
