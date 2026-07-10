/// <reference path="./sheetjs.d.ts" />

import { preprocessChats } from "./chatCleaner";
import { readWorkbook } from "./excelReader";
import { buildOutput } from "./excelWriter";
import { buildTargets } from "./listParser";
import { matchData } from "./matching";
import { progress } from "./progress";
import { buildReminderAppeals } from "./reminderAppealParser";
import { buildReminderOutput } from "./reminderExcelWriter";
import { buildReminderTargets } from "./reminderListParser";
import { matchReminderData } from "./reminderMatching";
import {
  applyReminderTouchSummary,
  mergeReminderTouchInfos,
  parseReminderTouchSummary,
  type ReminderTouchInfo,
} from "./reminderTouchSummary";
import { ensureSheetJs } from "./sheetJsLoader";
import type { ChatInfo, WorkerRequest } from "./types";
import { inferServiceWeek } from "./utils";
import { buildWhitelist } from "./whitelist";

interface WorkerScope {
  onmessage: ((event: MessageEvent<WorkerRequest>) => void | Promise<void>) | null;
  postMessage(message: unknown, transfer?: Transferable[]): void;
}

const workerScope = self as unknown as WorkerScope;

function mergeChatInfos(infos: ChatInfo[]): ChatInfo {
  const counts: ChatInfo["counts"] = {};
  infos.forEach((info) => {
    Object.entries(info.counts).forEach(([key, value]) => {
      counts[key] = (counts[key] || 0) + value;
    });
  });
  return {
    chats: infos.flatMap((info) => info.chats),
    counts,
    sheetName: infos.map((info) => info.sheetName).join("；"),
  };
}

function emptyChatInfo(): ChatInfo {
  return {
    chats: [],
    counts: {
      原始聊天行数: 0,
      清洗后聊天行数: 0,
    },
    sheetName: "",
  };
}

async function readSummaryFiles(files: File[], stageStart: number, stageEnd: number): Promise<ReminderTouchInfo> {
  if (!files.length) throw new Error("请至少上传 1 个聊天质检汇总文件。");
  const infos: ReminderTouchInfo[] = [];
  for (let index = 0; index < files.length; index += 1) {
    const summaryFile = files[index];
    const fileStart = stageStart + Math.floor((index * (stageEnd - stageStart)) / files.length);
    const fileEnd = stageStart + Math.floor(((index + 1) * (stageEnd - stageStart)) / files.length);
    const workbook = await readWorkbook(
      summaryFile,
      fileStart,
      fileEnd,
      `聊天汇总 ${index + 1}/${files.length}`,
    );
    infos.push(parseReminderTouchSummary(workbook, summaryFile.name));
  }
  return mergeReminderTouchInfos(infos);
}

async function readChatFiles(files: File[], stageStart: number, stageEnd: number): Promise<ChatInfo> {
  if (!files.length) return emptyChatInfo();
  const chatInfos: ChatInfo[] = [];
  for (let index = 0; index < files.length; index += 1) {
    const chatFile = files[index];
    const fileStart = stageStart + Math.floor((index * (stageEnd - stageStart)) / files.length);
    const fileEnd = stageStart + Math.floor(((index + 1) * (stageEnd - stageStart)) / files.length);
    const chatWorkbook = await readWorkbook(
      chatFile,
      fileStart,
      fileEnd,
      `聊天明细 ${index + 1}/${files.length}`,
    );
    chatInfos.push(preprocessChats(chatWorkbook, chatFile.name));
  }
  return mergeChatInfos(chatInfos);
}

function localMonthDay(date = new Date()) {
  return `${date.getMonth() + 1}.${date.getDate()}`;
}

workerScope.onmessage = async ({ data }: MessageEvent<WorkerRequest>) => {
  if (data.type !== "process") return;
  try {
    await ensureSheetJs();

    if (data.mode === "reminder") {
      const whitelist = data.whitelistCsv ? buildWhitelist(data.whitelistCsv) : buildWhitelist("");
      const listWorkbook = await readWorkbook(data.denominatorFile, 3, 22, "开课提醒学员明细");
      const listInfo = buildReminderTargets(listWorkbook, whitelist);
      progress(
        "分母预处理完成",
        `原始 ${listInfo.counts.原始分母行数.toLocaleString()} 条，整行去重后 ${listInfo.targets.length.toLocaleString()} 条。`,
        28,
      );
      let appealInfo;
      if (data.appealFile) {
        const appealWorkbook = await readWorkbook(data.appealFile, 28, 34, "申诉文件");
        appealInfo = buildReminderAppeals(appealWorkbook);
        progress(
          "申诉文件读取完成",
          `申诉 ${appealInfo.counts.申诉行数.toLocaleString()} 条，已发送申诉计入分母，其余申诉公示原因并剔除。`,
          34,
        );
      }

      const touchInfo = await readSummaryFiles(data.summaryFiles, 36, 58);
      progress(
        "汇总文件读取完成",
        `${data.summaryFiles.length.toLocaleString()} 个文件，汇总触达 ${touchInfo.counts.汇总触达数.toLocaleString()} 次。`,
        60,
      );

      const chatInfo = await readChatFiles(data.chatFiles, 60, 68);
      progress(
        data.chatFiles.length ? "聊天预处理完成" : "未上传聊天明细",
        data.chatFiles.length
          ? `${data.chatFiles.length.toLocaleString()} 个参考文件，原始 ${chatInfo.counts.原始聊天行数.toLocaleString()} 条，清洗后 ${chatInfo.chats.length.toLocaleString()} 条。`
          : "将只按汇总文件计算教师及以上维度触达完成率。",
        68,
      );

      progress("正在匹配开课提醒", "聊天明细用于学员名单参考；教师汇总会按汇总文件触达数重算。", 72);
      const baseMatchInfo = matchReminderData(listInfo, chatInfo.chats, appealInfo);
      const matchInfo = applyReminderTouchSummary(baseMatchInfo, touchInfo);
      progress(
        "匹配完成",
        `应发送 ${matchInfo.counts.应发送数.toLocaleString()}，有效触达 ${matchInfo.counts.有效触达数.toLocaleString()}，异常 ${matchInfo.counts.异常明细行数.toLocaleString()}。`,
        80,
      );

      progress("正在生成 Excel", "写入公示表、学员名单及维度汇总。", 84);
      const output = buildReminderOutput(listInfo, chatInfo, matchInfo, {
        list: data.denominatorFile.name,
        chat: data.chatFiles.map((file) => file.name).join("；"),
        summary: data.summaryFiles.map((file) => file.name).join("；"),
      }, data.includeCleanChats, data.includeResultColors, data.includeExceptionSheet, data.includeExplanationSheet);
      const buffer = output.buffer as ArrayBuffer;
      workerScope.postMessage({
        type: "complete",
        buffer,
        filename: `暑期开课提醒话术发送进度（${localMonthDay()}）.xlsx`,
        summary: {
          mode: "reminder",
          targets: listInfo.targets.length,
          sent: matchInfo.counts.有效触达数 || 0,
          unsent: Math.max(0, (matchInfo.counts.应发送数 || 0) - (matchInfo.counts.有效触达数 || 0)),
          exceptions: matchInfo.counts.异常明细行数,
          summaryFiles: data.summaryFiles.length,
          chatFiles: data.chatFiles.length,
          cleanChats: chatInfo.chats.length,
        },
      }, [buffer]);
      return;
    }

    const whitelist = buildWhitelist(data.whitelistCsv);
    const listWorkbook = await readWorkbook(data.listFile, 3, 18, "课堂反馈名单");
    const listInfo = buildTargets(listWorkbook, whitelist);
    progress(
      "名单预处理完成",
      `原始 ${listInfo.counts.原始课次行数.toLocaleString()} 条，去重后 ${listInfo.targets.length.toLocaleString()} 人。`,
      24,
    );

    const chatWorkbook = await readWorkbook(data.chatFile, 25, 48, "聊天明细");
    const chatInfo = preprocessChats(chatWorkbook);
    progress(
      "聊天预处理完成",
      `原始 ${chatInfo.counts.原始聊天行数.toLocaleString()} 条，清洗后 ${chatInfo.chats.length.toLocaleString()} 条。`,
      58,
    );

    progress("正在匹配教师与学员", "按教师邮箱建立索引，再检查群名和聊天内容中的学员关键词。", 64);
    const inferredWeek = inferServiceWeek(listInfo.weekCounts);
    const selectedWeek = data.weekLabel === "auto" ? inferredWeek.label : data.weekLabel;
    listInfo.weekMode = data.weekLabel === "auto" ? "根据课次日期自动识别" : "人工指定";
    listInfo.weekDistribution = inferredWeek.distribution;
    if (!selectedWeek) {
      throw new Error("无法从课次日期识别服务周，请手动选择第一周至第五周。");
    }
    progress(
      "服务周识别完成",
      `${selectedWeek}（${listInfo.weekMode}）${inferredWeek.distribution ? `；${inferredWeek.distribution}` : ""}`,
      61,
    );

    const matchInfo = matchData(
      listInfo.targets,
      chatInfo.chats,
      data.useSingle,
      selectedWeek,
      whitelist,
    );
    progress(
      "匹配完成",
      `已发送 ${matchInfo.counts.已发送.toLocaleString()}，未发送 ${matchInfo.counts.未发送.toLocaleString()}。`,
      80,
    );

    progress("正在生成 Excel", "写入打卡结果、匹配明细、清洗后聊天和处理说明。", 84);
    const output = buildOutput(
      listInfo,
      chatInfo,
      matchInfo,
      whitelist,
      data.useSingle,
      selectedWeek,
      { list: data.listFile.name, chat: data.chatFile.name },
    );
    const stamp = new Date().toISOString().slice(0, 19).replace(/[-:T]/g, "");
    const buffer = output.buffer as ArrayBuffer;
    workerScope.postMessage({
      type: "complete",
      buffer,
      filename: `打卡质检结果_${stamp}.xlsx`,
      summary: {
        targets: listInfo.targets.length,
        sent: matchInfo.counts.已发送,
        unsent: matchInfo.counts.未发送,
        exempt: matchInfo.counts.免检,
        cleanChats: chatInfo.chats.length,
      },
    }, [buffer]);
  } catch (error) {
    const message = error instanceof Error ? error.stack || error.message : String(error);
    workerScope.postMessage({ type: "error", message });
  }
};
