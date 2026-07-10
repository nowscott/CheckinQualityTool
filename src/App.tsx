import { useEffect, useMemo, useRef, useState, type FormEvent, type MouseEvent } from "react";
import { ChangelogDialog } from "./components/ChangelogDialog";
import { Header } from "./components/Header";
import { MatchingGuideDialog } from "./components/MatchingGuideDialog";
import { OutputGrid } from "./components/OutputGrid";
import { ReminderForm } from "./components/ReminderForm";
import { StatusCard } from "./components/StatusCard";
import { UploadForm } from "./components/UploadForm";
import { downloadResult } from "./lib/download";
import { inferWeekFromFilename } from "./lib/week";
import { useTheme } from "./hooks/useTheme";
import type { ProcessingStatus, WeekLabel, WorkerResponse } from "./types/worker";

type ActiveModal = "guide" | "changelog" | null;
type ToolMode = "checkin" | "reminder";

const INITIAL_STATUS: ProcessingStatus = {
  visible: false,
  title: "正在处理数据",
  message: "大文件需要一些时间，请不要关闭页面。",
  progress: 0,
  mode: "working",
};

let whitelistCsvPromise: Promise<string> | undefined;

function loadWhitelistCsv() {
  whitelistCsvPromise ||= fetch("/data/whitelist.csv", { cache: "no-store" }).then((response) => {
    if (!response.ok) throw new Error("内置白名单读取失败，请刷新页面后重试。");
    return response.text();
  });
  return whitelistCsvPromise;
}

function errorMessage(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}

function modeFromPath(): ToolMode {
  return window.location.pathname.replace(/\/+$/, "") === "/remind" ? "reminder" : "checkin";
}

function createProcessingWorker() {
  return new Worker(new URL("./worker/index.ts", import.meta.url), { type: "module" });
}

export default function App() {
  const [activeMode, setActiveMode] = useState<ToolMode>(modeFromPath);
  const [listFile, setListFile] = useState<File | null>(null);
  const [chatFile, setChatFile] = useState<File | null>(null);
  const [reminderListFile, setReminderListFile] = useState<File | null>(null);
  const [reminderAppealFile, setReminderAppealFile] = useState<File | null>(null);
  const [reminderSummaryFiles, setReminderSummaryFiles] = useState<File[]>([]);
  const [reminderChatFiles, setReminderChatFiles] = useState<File[]>([]);
  const [includeReminderChats, setIncludeReminderChats] = useState(false);
  const [includeReminderColors, setIncludeReminderColors] = useState(false);
  const [includeReminderExceptionSheet, setIncludeReminderExceptionSheet] = useState(false);
  const [includeReminderExplanationSheet, setIncludeReminderExplanationSheet] = useState(false);
  const [weekLabel, setWeekLabel] = useState<WeekLabel>("auto");
  const [useSingle, setUseSingle] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [status, setStatus] = useState(INITIAL_STATUS);
  const [activeModal, setActiveModal] = useState<ActiveModal>(null);
  const workerRef = useRef<Worker | null>(null);
  const { theme, usesSystemTheme, toggleTheme } = useTheme();

  useEffect(() => () => workerRef.current?.terminate(), []);

  useEffect(() => {
    const syncMode = () => setActiveMode(modeFromPath());
    window.addEventListener("popstate", syncMode);
    return () => window.removeEventListener("popstate", syncMode);
  }, []);

  const weekHint = useMemo(() => {
    if (weekLabel !== "auto") return `已手动指定为${weekLabel}`;
    const inferred = listFile ? inferWeekFromFilename(listFile.name) : "";
    return inferred
      ? `根据文件日期预计为${inferred}，生成时会再用课次日期校验`
      : "选择名单后，将根据上课日期自动判断";
  }, [listFile, weekLabel]);

  function updateStatus(
    title: string,
    message: string,
    progress = 0,
    mode: ProcessingStatus["mode"] = "working",
  ) {
    setStatus({ visible: true, title, message, progress, mode });
  }

  function finishWorker() {
    setProcessing(false);
    workerRef.current?.terminate();
    workerRef.current = null;
  }

  function changeMode(mode: ToolMode) {
    setActiveMode(mode);
    setStatus(INITIAL_STATUS);
    const path = mode === "reminder" ? "/remind" : "/";
    if (window.location.pathname !== path) window.history.pushState(null, "", path);
  }

  function handleSecondaryLinkClick(event: MouseEvent<HTMLAnchorElement>) {
    event.preventDefault();
    changeMode(activeMode === "reminder" ? "checkin" : "reminder");
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!listFile || !chatFile) return;

    workerRef.current?.terminate();
    const worker = createProcessingWorker();
    workerRef.current = worker;
    setProcessing(true);
    updateStatus("正在启动本地处理引擎", "所有文件只在当前浏览器中处理，不会上传。", 2);

    let whitelistCsv: string;
    try {
      whitelistCsv = await loadWhitelistCsv();
    } catch (error) {
      updateStatus("处理失败", errorMessage(error), 100, "error");
      finishWorker();
      return;
    }

    worker.onmessage = ({ data }: MessageEvent<WorkerResponse>) => {
      if (data.type === "progress") {
        updateStatus(data.title, data.message, data.progress);
        return;
      }
      if (data.type === "complete") {
        downloadResult(data.buffer, data.filename);
        updateStatus(
          "处理完成，结果已下载",
          `质检 ${data.summary.targets.toLocaleString()} 人：已发送 ${data.summary.sent.toLocaleString()}，未发送 ${data.summary.unsent.toLocaleString()}，免检 ${Number(data.summary.exempt || 0).toLocaleString()}；清洗后聊天 ${data.summary.cleanChats.toLocaleString()} 条。`,
          100,
          "done",
        );
        finishWorker();
        return;
      }
      updateStatus("处理失败", data.message, 100, "error");
      finishWorker();
    };

    worker.onerror = (event) => {
      updateStatus("处理失败", event.message || "浏览器工作线程发生错误。", 100, "error");
      finishWorker();
    };

    worker.postMessage({
      type: "process",
      listFile,
      chatFile,
      weekLabel,
      useSingle,
      whitelistCsv,
    });
  }

  async function handleReminderSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!reminderSummaryFiles.length) return;
    if (!reminderListFile) return;

    workerRef.current?.terminate();
    const worker = createProcessingWorker();
    workerRef.current = worker;
    setProcessing(true);
    updateStatus("正在启动本地处理引擎", "所有文件只在当前浏览器中处理，不会上传。", 2);

    worker.onmessage = ({ data }: MessageEvent<WorkerResponse>) => {
      if (data.type === "progress") {
        updateStatus(data.title, data.message, data.progress);
        return;
      }
      if (data.type === "complete") {
        downloadResult(data.buffer, data.filename);
        updateStatus(
          "处理完成，结果已下载",
          `开课提醒 ${data.summary.targets.toLocaleString()} 条：有效触达 ${data.summary.sent.toLocaleString()}，未触达 ${data.summary.unsent.toLocaleString()}，异常核对 ${Number(data.summary.exceptions || 0).toLocaleString()}；汇总文件 ${Number(data.summary.summaryFiles || 0).toLocaleString()} 个，聊天参考文件 ${Number(data.summary.chatFiles || 0).toLocaleString()} 个。`,
          100,
          "done",
        );
        finishWorker();
        return;
      }
      updateStatus("处理失败", data.message, 100, "error");
      finishWorker();
    };

    worker.onerror = (event) => {
      updateStatus("处理失败", event.message || "浏览器工作线程发生错误。", 100, "error");
      finishWorker();
    };

    let whitelistCsv: string;
    try {
      whitelistCsv = await loadWhitelistCsv();
    } catch (error) {
      updateStatus("处理失败", errorMessage(error), 100, "error");
      finishWorker();
      return;
    }

    worker.postMessage({
      type: "process",
      mode: "reminder",
      denominatorFile: reminderListFile,
      appealFile: reminderAppealFile,
      summaryFiles: reminderSummaryFiles,
      chatFiles: reminderChatFiles,
      includeCleanChats: includeReminderChats,
      includeResultColors: includeReminderColors,
      includeExceptionSheet: includeReminderExceptionSheet,
      includeExplanationSheet: includeReminderExplanationSheet,
      whitelistCsv,
    });
  }

  return (
    <>
      <main className="shell">
        <Header
          theme={theme}
          usesSystemTheme={usesSystemTheme}
          title={activeMode === "reminder" ? "开课提醒触达完成率公示" : "打卡质检数据生成"}
          subtitle={
            activeMode === "reminder"
              ? "上传开课提醒学员明细名单与聊天质检汇总文件，在浏览器本地计算教师及以上维度触达完成率。文件不会上传服务器。"
              : undefined
          }
          showGuide={activeMode === "checkin"}
          secondaryLink={
            activeMode === "reminder"
              ? { href: "/", label: "返回打卡质检" }
              : { href: "/remind", label: "开课提醒" }
          }
          onToggleTheme={toggleTheme}
          onOpenGuide={() => setActiveModal("guide")}
          onOpenChangelog={() => setActiveModal("changelog")}
          onSecondaryLinkClick={handleSecondaryLinkClick}
        />
        {activeMode === "reminder" ? (
          <ReminderForm
            denominatorFile={reminderListFile}
            appealFile={reminderAppealFile}
            summaryFiles={reminderSummaryFiles}
            chatFiles={reminderChatFiles}
            includeCleanChats={includeReminderChats}
            includeResultColors={includeReminderColors}
            includeExceptionSheet={includeReminderExceptionSheet}
            includeExplanationSheet={includeReminderExplanationSheet}
            processing={processing}
            onDenominatorFileChange={setReminderListFile}
            onAppealFileChange={setReminderAppealFile}
            onSummaryFilesChange={setReminderSummaryFiles}
            onChatFilesChange={setReminderChatFiles}
            onIncludeCleanChatsChange={setIncludeReminderChats}
            onIncludeResultColorsChange={setIncludeReminderColors}
            onIncludeExceptionSheetChange={setIncludeReminderExceptionSheet}
            onIncludeExplanationSheetChange={setIncludeReminderExplanationSheet}
            onSubmit={handleReminderSubmit}
          />
        ) : (
          <UploadForm
            listFile={listFile}
            chatFile={chatFile}
            weekLabel={weekLabel}
            weekHint={weekHint}
            useSingle={useSingle}
            processing={processing}
            onListFileChange={setListFile}
            onChatFileChange={setChatFile}
            onWeekLabelChange={setWeekLabel}
            onUseSingleChange={setUseSingle}
            onSubmit={handleSubmit}
          />
        )}
        <StatusCard status={status} />
        <OutputGrid mode={activeMode} />
      </main>

      <ChangelogDialog
        open={activeModal === "changelog"}
        onClose={() => setActiveModal(null)}
      />
      <MatchingGuideDialog
        open={activeModal === "guide"}
        onClose={() => setActiveModal(null)}
      />
    </>
  );
}
