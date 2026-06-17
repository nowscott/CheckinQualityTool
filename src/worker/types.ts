import type { WeekLabel } from "../types/worker";

export type CellValue = string | number | boolean | Date | null | undefined;
export type DataRow = Record<string, unknown>;
export type CountMap = Record<string, number>;

export interface CleanStudentName {
  original: string;
  cleaned: string;
  note: string;
}

export interface TargetRow {
  教师姓名: string;
  教师邮箱: string;
  学员姓名: string;
  原始学员姓名: string;
  姓名清洗说明: string;
  学员号: string;
  上课日期: CellValue;
  上课开始: string;
  上课结束: string;
  校区: string;
  项目组: string;
  科目: string;
  源名单行号: number;
  该周课次数: number;
}

export interface ChatRow extends DataRow {
  有效教师邮箱: string;
  邮箱来源: string;
  发送人名称: string;
  "群名/好友昵称": string;
  聊天时间: CellValue;
  聊天内容: string;
  源聊天行号: number;
}

export interface WhitelistEntry extends DataRow {
  学员号: string;
  学员姓名: string;
  匹配学员姓名: string;
  处理方式: string;
  匹配别名: string;
  匹配别名关键词: string[];
  说明: string;
}

export interface Whitelist {
  entries: WhitelistEntry[];
  byStudentId: Map<string, WhitelistEntry>;
  byOriginalStudentName: Map<string, WhitelistEntry>;
  byStudentName: Map<string, WhitelistEntry>;
}

export interface ListInfo {
  targets: TargetRow[];
  counts: CountMap;
  weekCounts: Map<number, number>;
  sheetName: string;
  weekMode?: string;
  weekDistribution?: string;
}

export interface ChatInfo {
  chats: ChatRow[];
  counts: CountMap;
  sheetName: string;
}

export interface MatchInfo {
  finalRows: DataRow[];
  detailRows: DataRow[];
  counts: CountMap;
}

export interface SourceNames {
  list: string;
  chat: string;
}

export interface ProcessRequest {
  type: "process";
  listFile: File;
  chatFile: File;
  weekLabel: WeekLabel;
  useSingle: boolean;
  whitelistCsv: string;
}

export interface SheetDefinition {
  name: string;
  rows: DataRow[];
  columns: readonly string[];
  widths: Record<string, number>;
  rowStyle?: (row: DataRow) => number;
}
