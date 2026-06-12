export type WeekLabel = "auto" | "第一周" | "第二周" | "第三周" | "第四周" | "第五周";

export type StatusMode = "working" | "done" | "error";

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
  exempt: number;
  cleanChats: number;
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
      buffer: ArrayBuffer;
      filename: string;
      summary: ResultSummary;
    }
  | {
      type: "error";
      message: string;
    };
