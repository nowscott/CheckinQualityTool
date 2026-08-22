import { findSheet, headerMap } from "./excelReader";
import { REMINDER_PASS_RATE } from "./reminderConfig";
import { normalizeTeacherName } from "./utils";
import type { CellValue, CountMap, DataRow } from "./types";
import { emailValue, text } from "./utils";
import type { ReminderMatchInfo } from "./reminderMatching";

const REQUIRED_HEADERS = ["姓名", "邮箱", "员工触达客户数", "员工触达群数"] as const;

export interface ReminderTouchRecord {
  姓名: string;
  规范姓名: string;
  邮箱: string;
  触达客户数: number;
  触达群聊数: number;
  汇总触达数: number;
  来源文件: string;
  源行号: number;
}

export interface ReminderTouchBucket {
  name: string;
  normalizedName: string;
  email: string;
  customerTouches: number;
  groupTouches: number;
  totalTouches: number;
  sourceFiles: Set<string>;
  records: ReminderTouchRecord[];
}

export interface ReminderTouchInfo {
  records: ReminderTouchRecord[];
  buckets: ReminderTouchBucket[];
  byEmail: Map<string, ReminderTouchBucket>;
  byName: Map<string, ReminderTouchBucket>;
  counts: CountMap;
  sheetName: string;
}

function findFirstIndex(map: Map<string, number[]>, aliases: readonly string[]) {
  for (const alias of aliases) {
    const index = (map.get(alias) || [])[0];
    if (index != null) return index;
  }
  return -1;
}

function numberValue(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) return Math.max(0, value);
  const parsed = Number(text(value).replace(/,/g, ""));
  return Number.isFinite(parsed) ? Math.max(0, parsed) : 0;
}

function emptyTouchInfo(): ReminderTouchInfo {
  return {
    records: [],
    buckets: [],
    byEmail: new Map(),
    byName: new Map(),
    counts: {
      汇总文件数: 0,
      原始汇总行数: 0,
      有效汇总行数: 0,
      汇总触达客户数: 0,
      汇总触达群聊数: 0,
      汇总触达数: 0,
    },
    sheetName: "",
  };
}

function mergeBucket(target: ReminderTouchBucket, record: ReminderTouchRecord) {
  target.customerTouches = Math.max(target.customerTouches, record.触达客户数);
  target.groupTouches = Math.max(target.groupTouches, record.触达群聊数);
  target.totalTouches = target.customerTouches + target.groupTouches;
  target.sourceFiles.add(record.来源文件);
  target.records.push(record);
}

function bucketFromRecord(record: ReminderTouchRecord): ReminderTouchBucket {
  return {
    name: record.姓名,
    normalizedName: record.规范姓名,
    email: record.邮箱,
    customerTouches: record.触达客户数,
    groupTouches: record.触达群聊数,
    totalTouches: record.汇总触达数,
    sourceFiles: new Set([record.来源文件]),
    records: [record],
  };
}

export function parseReminderTouchSummary(workbook: SheetJsWorkbook, sourceFile: string): ReminderTouchInfo {
  const found = findSheet(workbook, [...REQUIRED_HEADERS]);
  const rows = found.rows;
  const map = headerMap(rows[0]);
  const column = {
    name: findFirstIndex(map, ["姓名"]),
    email: findFirstIndex(map, ["邮箱"]),
    customer: findFirstIndex(map, ["员工触达客户数"]),
    group: findFirstIndex(map, ["员工触达群数"]),
  };
  const records: ReminderTouchRecord[] = [];
  const counts: CountMap = {
    汇总文件数: 1,
    原始汇总行数: Math.max(0, rows.length - 1),
    有效汇总行数: 0,
    跳过姓名邮箱均为空: 0,
    汇总触达客户数: 0,
    汇总触达群聊数: 0,
    汇总触达数: 0,
  };

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] as CellValue[];
    const name = text(row[column.name]);
    const email = emailValue(row[column.email]);
    const normalizedName = normalizeTeacherName(name);
    if (!email && !normalizedName) {
      counts.跳过姓名邮箱均为空 += 1;
      continue;
    }
    const customerTouches = numberValue(row[column.customer]);
    const groupTouches = numberValue(row[column.group]);
    const record: ReminderTouchRecord = {
      姓名: name,
      规范姓名: normalizedName,
      邮箱: email,
      触达客户数: customerTouches,
      触达群聊数: groupTouches,
      汇总触达数: customerTouches + groupTouches,
      来源文件: sourceFile,
      源行号: rowIndex + 1,
    };
    records.push(record);
    counts.有效汇总行数 += 1;
    counts.汇总触达客户数 += customerTouches;
    counts.汇总触达群聊数 += groupTouches;
    counts.汇总触达数 += record.汇总触达数;
  }

  return mergeReminderTouchInfos([{ ...emptyTouchInfo(), records, counts, sheetName: found.name }]);
}

export function mergeReminderTouchInfos(infos: ReminderTouchInfo[]): ReminderTouchInfo {
  const records = infos.flatMap((info) => info.records);
  const counts: CountMap = {
    汇总文件数: infos.reduce((sum, info) => sum + (info.counts.汇总文件数 || 0), 0),
    原始汇总行数: infos.reduce((sum, info) => sum + (info.counts.原始汇总行数 || 0), 0),
    有效汇总行数: infos.reduce((sum, info) => sum + (info.counts.有效汇总行数 || 0), 0),
    跳过姓名邮箱均为空: infos.reduce((sum, info) => sum + (info.counts.跳过姓名邮箱均为空 || 0), 0),
    汇总触达客户数: infos.reduce((sum, info) => sum + (info.counts.汇总触达客户数 || 0), 0),
    汇总触达群聊数: infos.reduce((sum, info) => sum + (info.counts.汇总触达群聊数 || 0), 0),
    汇总触达数: infos.reduce((sum, info) => sum + (info.counts.汇总触达数 || 0), 0),
  };
  const byEmail = new Map<string, ReminderTouchBucket>();
  const byName = new Map<string, ReminderTouchBucket>();
  const bucketsByKey = new Map<string, ReminderTouchBucket>();

  records.forEach((record) => {
    const key = record.邮箱 ? `email:${record.邮箱}` : `name:${record.规范姓名}`;
    const existed = bucketsByKey.get(key);
    if (existed) {
      mergeBucket(existed, record);
    } else {
      bucketsByKey.set(key, bucketFromRecord(record));
    }
  });

  const buckets = [...bucketsByKey.values()];
  buckets.forEach((bucket) => {
    if (bucket.email) byEmail.set(bucket.email, bucket);
    if (bucket.normalizedName) {
      const existed = byName.get(bucket.normalizedName);
      if (existed && existed !== bucket) {
        mergeBucket(existed, {
          姓名: bucket.name,
          规范姓名: bucket.normalizedName,
          邮箱: bucket.email,
          触达客户数: bucket.customerTouches,
          触达群聊数: bucket.groupTouches,
          汇总触达数: bucket.totalTouches,
          来源文件: [...bucket.sourceFiles].join("；"),
          源行号: 0,
        });
      } else {
        byName.set(bucket.normalizedName, bucket);
      }
    }
  });

  return {
    records,
    buckets,
    byEmail,
    byName,
    counts,
    sheetName: infos.map((info) => info.sheetName).filter(Boolean).join("；"),
  };
}

function rateText(value: number) {
  return `${(value * 100).toFixed(1)}%`;
}

function passText(value: number) {
  return value >= REMINDER_PASS_RATE ? "是" : "否";
}

function matchedBucket(row: DataRow, touchInfo: ReminderTouchInfo) {
  const email = emailValue(row.教师邮箱 || row.邮箱);
  if (email && touchInfo.byEmail.has(email)) {
    return { bucket: touchInfo.byEmail.get(email)!, method: "邮箱匹配" };
  }
  const normalizedName = normalizeTeacherName(row.教师姓名 || row.授课教师);
  if (normalizedName && touchInfo.byName.has(normalizedName)) {
    return { bucket: touchInfo.byName.get(normalizedName)!, method: "姓名匹配" };
  }
  return null;
}

export function applyReminderTouchSummary(
  matchInfo: ReminderMatchInfo,
  touchInfo: ReminderTouchInfo,
): ReminderMatchInfo {
  const matchedEmails = new Set<string>();
  const matchedNames = new Set<string>();
  let summaryTouches = 0;
  let effectiveTouches = 0;
  let teacherRowsWithoutSummary = 0;

  const teacherRows = matchInfo.teacherRows.map((row) => {
    const total = Number(row.应发送数) || 0;
    const matched = matchedBucket(row, touchInfo);
    if (!matched) {
      if (total > 0) teacherRowsWithoutSummary += 1;
      const existingSent = Number(row.已发送数) || 0;
      const currentRate = total ? existingSent / total : 1;
      return {
        ...row,
        汇总触达数: existingSent,
        有效触达数: existingSent,
        触达差额: Math.max(0, total - existingSent),
        触达完成率: rateText(currentRate),
        触达匹配方式: "",
        触达来源文件: "",
        是否达标: passText(currentRate),
      };
    }
    if (matched.bucket.email) matchedEmails.add(matched.bucket.email);
    if (matched.bucket.normalizedName) matchedNames.add(matched.bucket.normalizedName);
    const rawTouches = matched.bucket.totalTouches;
    const effective = Math.min(Math.max(rawTouches, Number(row.已发送数) || 0), total);
    const currentRate = total ? effective / total : 1;
    summaryTouches += rawTouches;
    effectiveTouches += effective;
    return {
      ...row,
      已发送数: effective,
      汇总触达数: rawTouches,
      有效触达数: effective,
      触达差额: Math.max(0, total - effective),
      发送差额: Math.max(0, total - effective),
      发送率: rateText(currentRate),
      触达完成率: rateText(currentRate),
      触达匹配方式: matched.method,
      触达来源文件: [...matched.bucket.sourceFiles].sort((a, b) => a.localeCompare(b, "zh-CN")).join("；"),
      是否达标: passText(currentRate),
    };
  });

  const unmatchedSummaryRecords = touchInfo.records.filter((record) => {
    if (record.邮箱 && matchedEmails.has(record.邮箱)) return false;
    return !(record.规范姓名 && matchedNames.has(record.规范姓名));
  }).length;

  return {
    ...matchInfo,
    teacherRows,
    counts: {
      ...matchInfo.counts,
      上传汇总文件数: touchInfo.counts.汇总文件数 || 0,
      上传汇总原始行数: touchInfo.counts.原始汇总行数 || 0,
      上传汇总有效行数: touchInfo.counts.有效汇总行数 || 0,
      上传汇总触达数: touchInfo.counts.汇总触达数 || 0,
      汇总触达数: summaryTouches,
      有效触达数: effectiveTouches,
      未匹配到分母老师的汇总记录数: unmatchedSummaryRecords,
      分母老师未匹配汇总记录数: teacherRowsWithoutSummary,
    },
  };
}
