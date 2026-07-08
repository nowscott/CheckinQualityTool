/// <reference path="./sheetjs.d.ts" />

import { preprocessChats } from "./chatCleaner";
import { readWorkbook } from "./excelReader";
import { buildOutput } from "./excelWriter";
import { buildTargets } from "./listParser";
import { matchData } from "./matching";
import { progress } from "./progress";
import { buildReminderAppeals } from "./reminderAppealParser";
import { buildIncrementalReminderOutput } from "./reminderIncremental";
import { buildReminderOutput } from "./reminderExcelWriter";
import { buildReminderTargets } from "./reminderListParser";
import { matchReminderData } from "./reminderMatching";
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

function localMonthDay(date = new Date()) {
  return `${date.getMonth() + 1}.${date.getDate()}`;
}

workerScope.onmessage = async ({ data }: MessageEvent<WorkerRequest>) => {
  if (data.type !== "process") return;
  try {
    await ensureSheetJs();

    if (data.mode === "reminder") {
      if (!data.chatFiles.length) throw new Error("请至少上传 1 个企微聊天质检结果文件。");
      if (data.reminderMode === "incremental") {
        const previousWorkbook = await readWorkbook(data.previousFile, 3, 22, "上次开课提醒结果");
        progress("上次结果读取完成", "正在从“学员名单”中识别未发送行。", 26);

        const chatInfos: ChatInfo[] = [];
        for (let index = 0; index < data.chatFiles.length; index += 1) {
          const chatFile = data.chatFiles[index];
          const stageStart = 30 + Math.floor((index * 26) / data.chatFiles.length);
          const stageEnd = 30 + Math.floor(((index + 1) * 26) / data.chatFiles.length);
          const chatWorkbook = await readWorkbook(
            chatFile,
            stageStart,
            stageEnd,
            `聊天明细 ${index + 1}/${data.chatFiles.length}`,
          );
          chatInfos.push(preprocessChats(chatWorkbook, chatFile.name));
        }
        const chatInfo = mergeChatInfos(chatInfos);
        progress(
          "聊天预处理完成",
          `${data.chatFiles.length.toLocaleString()} 个文件，原始 ${chatInfo.counts.原始聊天行数.toLocaleString()} 条，清洗后 ${chatInfo.chats.length.toLocaleString()} 条。`,
          60,
        );

        progress("正在增量匹配开课提醒", "仅对上次结果中仍未发送的学员行尝试补齐命中信息。", 70);
        const result = buildIncrementalReminderOutput(previousWorkbook, chatInfo, {
          list: data.previousFile.name,
          chat: data.chatFiles.map((file) => file.name).join("；"),
        }, data.includeCleanChats, data.includeResultColors);
        progress(
          "增量匹配完成",
          `本次新增发送 ${result.summary.incrementalSent.toLocaleString()} 条，当前已发送 ${result.summary.sent.toLocaleString()} 条。`,
          82,
        );

        const buffer = result.output.buffer as ArrayBuffer;
        workerScope.postMessage({
          type: "complete",
          buffer,
          filename: `暑期课程提醒-增量（${localMonthDay()}）.xlsx`,
          summary: {
            mode: "reminder",
            reminderMode: "incremental",
            targets: result.summary.targets,
            sent: result.summary.sent,
            unsent: result.summary.unsent,
            exceptions: result.summary.exceptions,
            incrementalSent: result.summary.incrementalSent,
            chatFiles: data.chatFiles.length,
            cleanChats: chatInfo.chats.length,
          },
        }, [buffer]);
        return;
      }
      const listWorkbook = await readWorkbook(data.denominatorFile, 3, 22, "开课提醒学员明细");
      const listInfo = buildReminderTargets(listWorkbook);
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
          `通过申诉 ${appealInfo.counts.申诉通过行数.toLocaleString()} 条，将从发送率分母中剔除。`,
          34,
        );
      }

      const chatInfos: ChatInfo[] = [];
      for (let index = 0; index < data.chatFiles.length; index += 1) {
        const chatFile = data.chatFiles[index];
        const stageStart = 30 + Math.floor((index * 26) / data.chatFiles.length);
        const stageEnd = 30 + Math.floor(((index + 1) * 26) / data.chatFiles.length);
        const chatWorkbook = await readWorkbook(
          chatFile,
          stageStart,
          stageEnd,
          `聊天明细 ${index + 1}/${data.chatFiles.length}`,
        );
        chatInfos.push(preprocessChats(chatWorkbook, chatFile.name));
      }
      const chatInfo = mergeChatInfos(chatInfos);
      progress(
        "聊天预处理完成",
        `${data.chatFiles.length.toLocaleString()} 个文件，原始 ${chatInfo.counts.原始聊天行数.toLocaleString()} 条，清洗后 ${chatInfo.chats.length.toLocaleString()} 条。`,
        60,
      );

      progress("正在匹配开课提醒", "按群聊名称和聊天内容中的学员姓名执行三档优先级匹配。", 68);
      const matchInfo = matchReminderData(listInfo, chatInfo.chats, appealInfo);
      progress(
        "匹配完成",
        `应发送 ${matchInfo.counts.应发送数.toLocaleString()}，已发送 ${matchInfo.counts.已发送数.toLocaleString()}，异常 ${matchInfo.counts.异常明细行数.toLocaleString()}。`,
        80,
      );

      progress("正在生成 Excel", "写入公示表、学员名单、维度汇总、异常明细和处理说明。", 84);
      const output = buildReminderOutput(listInfo, chatInfo, matchInfo, {
        list: data.denominatorFile.name,
        chat: data.chatFiles.map((file) => file.name).join("；"),
      }, data.includeCleanChats, data.includeResultColors);
      const buffer = output.buffer as ArrayBuffer;
      workerScope.postMessage({
        type: "complete",
        buffer,
        filename: `暑期课程提醒（${localMonthDay()}）.xlsx`,
        summary: {
          mode: "reminder",
          targets: listInfo.targets.length,
          sent: matchInfo.counts.已发送数,
          unsent: matchInfo.counts.未发送数,
          exceptions: matchInfo.counts.异常明细行数,
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
