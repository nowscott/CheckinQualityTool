export type WeekLabel = "auto" | "第一周" | "第二周" | "第三周" | "第四周" | "第五周";

export type StatusMode = "working" | "done" | "error";

import type { InspectionPriorityMode, InspectionPrioritySummary, InspectionStats, InspectionTeacherScore } from "../worker/inspectionTypes";

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
  sampleCount: number;
  attempt: number;
  includeExplanation: boolean;
  priorityMode: InspectionPriorityMode;
  focusTeacherNames: string[];
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
