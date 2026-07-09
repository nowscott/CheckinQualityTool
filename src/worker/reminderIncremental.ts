import { findSheet } from "./excelReader";
import { buildReminderOutput } from "./reminderExcelWriter";
import type { ReminderListInfo, ReminderTarget } from "./reminderListParser";
import {
  applyReminderTouchSummary,
  type ReminderTouchInfo,
} from "./reminderTouchSummary";
import {
  buildReminderGroupSummary,
  matchReminderData,
  type ReminderMatchInfo,
} from "./reminderMatching";
import type { CellValue, ChatInfo, CountMap, DataRow, SourceNames, Whitelist } from "./types";
import { normalizeMatchText, text } from "./utils";
import { findPreCleanWhitelistEntry } from "./whitelist";

const STUDENT_REQUIRED_HEADERS = ["教师姓名", "学员姓名", "匹配学员姓名", "是否发送"];
const EXCEPTION_REQUIRED_HEADERS = ["异常类型", "异常原因"];
const UPDATE_COLUMNS = [
  "是否发送",
  "匹配状态",
  "匹配方式",
  "异常原因",
  "命中位置",
  "命中关键词",
  "命中群名",
  "命中聊天时间",
  "发送人名称",
  "发送人邮箱",
  "命中质检文件",
  "源聊天行号",
  "匹配消息数",
] as const;

function rowsToObjects(rows: CellValue[][]): DataRow[] {
  const headers = rows[0].map(text);
  return rows.slice(1).map((row) => {
    const item: DataRow = {};
    headers.forEach((header, index) => {
      if (header) item[header] = row[index] ?? "";
    });
    if (item.邮箱 && !item.教师邮箱) item.教师邮箱 = item.邮箱;
    return item;
  });
}

function optionalSheetRows(workbook: SheetJsWorkbook, requiredHeaders: string[]) {
  try {
    return rowsToObjects(findSheet(workbook, requiredHeaders).rows);
  } catch {
    return [];
  }
}

function numberValue(value: unknown, fallback: number) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function normalizePersonName(value: unknown) {
  return normalizeMatchText(value).replace(/\s+/g, "").replace(/[0-9０-９]+$/u, "");
}

function assistantSummarySourceRows(rows: DataRow[]) {
  const ownGroups = new Map<string, unknown>();
  rows.forEach((row) => {
    const teacher = normalizePersonName(row.教师姓名 || row.授课教师);
    const assistant = normalizePersonName(row.助理主管);
    if (teacher && assistant && teacher === assistant && row.教研组) {
      ownGroups.set(assistant, row.教研组);
    }
  });
  return rows.map((row) => {
    const ownGroup = ownGroups.get(normalizePersonName(row.助理主管));
    return ownGroup ? { ...row, 教研组: ownGroup } : row;
  });
}

function rowToTarget(row: DataRow, index: number, whitelist?: Whitelist): ReminderTarget {
  const id = numberValue(row.质检序号, index + 1);
  const studentId = text(row.学员号);
  const rawStudent = text(row.学员姓名);
  const whitelistEntry = whitelist ? findPreCleanWhitelistEntry(studentId, rawStudent, whitelist) : null;
  return {
    id,
    授课教师: text(row.教师姓名 || row.授课教师),
    教师邮箱: text(row.教师邮箱 || row.邮箱),
    学员号: studentId,
    教研组: text(row.教研组),
    师训组长: text(row.师训组长),
    助理主管: text(row.助理主管),
    学员姓名: rawStudent,
    匹配学员姓名: text(row.匹配学员姓名 || row.学员姓名),
    匹配别名关键词: whitelistEntry?.处理方式 === "别名" ? whitelistEntry.匹配别名关键词 : [],
    白名单命中: whitelistEntry ? "是" : text(row.白名单命中),
    白名单说明: whitelistEntry?.说明 || text(row.白名单说明),
    姓名清洗说明: text(row.姓名清洗说明),
    新老生季度: text(row.新老生季度),
    年级: text(row.年级),
    学管: text(row.学管姓名 || row.学管),
    课时: text(row.课时),
    校区: text(row.校区),
    源名单行号: numberValue(row.源名单行号, index + 2),
    去重合并行号: text(row.去重合并行号),
  };
}

function shouldTryIncrementalMatch(row: DataRow) {
  const status = text(row.是否发送);
  return status !== "是" && status !== "已申诉" && status !== "已申诉通过";
}

function buildCounts(rows: DataRow[], incrementalSent: number, eligible: number): CountMap {
  const counts: CountMap = {
    应发送数: 0,
    已发送数: 0,
    未发送数: 0,
    申诉数: 0,
    待核对数: 0,
    本次可增量匹配数: eligible,
    本次新增发送数: incrementalSent,
    异常明细行数: 0,
  };
  rows.forEach((row) => {
    const status = text(row.是否发送);
    if (status === "已申诉" || status === "已申诉通过") {
      counts.申诉数 += 1;
    } else {
      counts.应发送数 += 1;
      if (status === "是") counts.已发送数 += 1;
      else counts.未发送数 += 1;
    }
    if (text(row.匹配状态).includes("核对")) counts.待核对数 += 1;
  });
  return counts;
}

function buildListInfoFromPrevious(rows: DataRow[], sheetName: string): ReminderListInfo {
  const targets = rows.map((row, index) => rowToTarget(row, index));
  return {
    targets,
    sheetName,
    counts: {
      上次结果行数: rows.length,
      本次增量候选行数: rows.filter(shouldTryIncrementalMatch).length,
    },
  };
}

export function buildIncrementalReminderOutput(
  previousWorkbook: SheetJsWorkbook,
  chatInfo: ChatInfo,
  sourceNames: SourceNames,
  includeCleanChats: boolean,
  includeResultColors = false,
  whitelist?: Whitelist,
  touchInfo?: ReminderTouchInfo,
) {
  const studentSheet = findSheet(previousWorkbook, STUDENT_REQUIRED_HEADERS);
  const previousRows = rowsToObjects(studentSheet.rows);
  const exceptionRows = optionalSheetRows(previousWorkbook, EXCEPTION_REQUIRED_HEADERS);
  const eligibleTargets = previousRows
    .map((row, index) => ({ row, target: rowToTarget(row, index) }))
    .filter(({ row, target }) => shouldTryIncrementalMatch(row) && target.授课教师 && target.匹配学员姓名);
  const listInfo: ReminderListInfo = {
    targets: eligibleTargets.map((item, index) => rowToTarget(item.row, index, whitelist)),
    counts: {
      上次结果行数: previousRows.length,
      本次增量候选行数: eligibleTargets.length,
    },
    sheetName: studentSheet.name,
  };
  const incrementalMatches = matchReminderData(listInfo, chatInfo.chats);
  const generatedById = new Map(
    incrementalMatches.studentRows.map((row) => [numberValue(row.质检序号, 0), row]),
  );
  let incrementalSent = 0;
  const finalRows = previousRows.map((row, index) => {
    if (!shouldTryIncrementalMatch(row)) return row;
    const target = rowToTarget(row, index, whitelist);
    const generated = generatedById.get(target.id);
    if (!generated || text(generated.是否发送) !== "是") return row;
    incrementalSent += 1;
    const next = { ...row };
    UPDATE_COLUMNS.forEach((column) => {
      next[column] = generated[column] ?? "";
    });
    return next;
  });
  const counts = buildCounts(finalRows, incrementalSent, eligibleTargets.length);
  const baseMatchInfo: ReminderMatchInfo = {
    studentRows: finalRows,
    teacherRows: buildReminderGroupSummary(
      finalRows,
      ["教师姓名", "教师邮箱", "教研组", "师训组长", "助理主管", "校区"],
      ["教师姓名", "教师邮箱", "教研组", "师训组长", "助理主管", "校区"],
    ),
    trainingRows: buildReminderGroupSummary(finalRows, ["教研组", "师训组长"], ["教研组", "师训组长"]),
    assistantRows: buildReminderGroupSummary(assistantSummarySourceRows(finalRows), ["助理主管"], ["教研组", "助理主管"]),
    exceptionRows: [...exceptionRows, ...incrementalMatches.exceptionRows],
    counts,
  };
  const matchInfo = touchInfo ? applyReminderTouchSummary(baseMatchInfo, touchInfo) : baseMatchInfo;
  matchInfo.counts.异常明细行数 = matchInfo.exceptionRows.length;
  return {
    output: buildReminderOutput(
      buildListInfoFromPrevious(previousRows, studentSheet.name),
      chatInfo,
      matchInfo,
      sourceNames,
      includeCleanChats,
      includeResultColors,
    ),
    summary: {
      targets: finalRows.length,
      sent: matchInfo.counts.有效触达数 || counts.已发送数,
      unsent: Math.max(0, (matchInfo.counts.应发送数 || counts.应发送数) - (matchInfo.counts.有效触达数 || counts.已发送数)),
      exceptions: matchInfo.counts.异常明细行数,
      incrementalSent,
    },
  };
}
