/// <reference path="./sheetjs.d.ts" />

import { preprocessChats } from "./chatCleaner";
import { readWorkbook } from "./excelReader";
import { buildOutput } from "./excelWriter";
import { buildTargets } from "./listParser";
import { matchData } from "./matching";
import { progress } from "./progress";
import type { ProcessRequest } from "./types";
import { inferServiceWeek } from "./utils";
import { buildWhitelist } from "./whitelist";

declare function importScripts(...urls: string[]): void;

interface WorkerScope {
  onmessage: ((event: MessageEvent<ProcessRequest>) => void | Promise<void>) | null;
  postMessage(message: unknown, transfer?: Transferable[]): void;
}

const workerScope = self as unknown as WorkerScope;

importScripts("/vendor/xlsx.full.min.js");

workerScope.onmessage = async ({ data }: MessageEvent<ProcessRequest>) => {
  if (data.type !== "process") return;
  try {
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
