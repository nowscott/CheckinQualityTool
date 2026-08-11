import { useEffect, useMemo, useRef, useState, type FormEvent, type MouseEvent } from "react";
import { ChangelogDialog } from "./components/ChangelogDialog";
import { Header } from "./components/Header";
import { MatchingGuideDialog } from "./components/MatchingGuideDialog";
import { OutputGrid } from "./components/OutputGrid";
import { ReminderForm } from "./components/ReminderForm";
import { StageReportBeautifyForm } from "./components/StageReportBeautifyForm";
import { StageReportForm } from "./components/StageReportForm";
import { StatusCard } from "./components/StatusCard";
import { UploadForm } from "./components/UploadForm";
import { downloadResult } from "./lib/download";
import { inferWeekFromFilename } from "./lib/week";
import { useTheme } from "./hooks/useTheme";
import type { ProcessingStatus, WeekLabel, WorkerResponse } from "./types/worker";

type ActiveModal = "guide" | "changelog" | null;
type ToolMode = "checkin" | "reminder" | "stageReport";
type StageReportTab = "check" | "publish";
type StatusKey = ToolMode | "stageReportPublish";

interface RouteState {
  mode: ToolMode;
  stageReportTab: StageReportTab;
}

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

function routeFromLocation(): RouteState {
  const path = window.location.pathname.replace(/\/+$/, "") || "/";
  if (path === "/remind") return { mode: "reminder", stageReportTab: "check" };
  if (path === "/report") {
    return {
      mode: "stageReport",
      stageReportTab: window.location.hash === "#publish" ? "publish" : "check",
    };
  }
  if (path === "/stage-report/beautify") {
    return { mode: "stageReport", stageReportTab: "publish" };
  }
  if (path === "/stage-report") return { mode: "stageReport", stageReportTab: "check" };
  return { mode: "checkin", stageReportTab: "check" };
}

function routeUrl(route: RouteState) {
  if (route.mode === "reminder") return "/remind";
  if (route.mode === "stageReport") {
    return route.stageReportTab === "publish" ? "/report#publish" : "/report";
  }
  return "/";
}

function initialRoute() {
  const route = routeFromLocation();
  const canonicalUrl = routeUrl(route);
  if (`${window.location.pathname}${window.location.hash}` !== canonicalUrl) {
    window.history.replaceState(null, "", canonicalUrl);
  }
  return route;
}

function createProcessingWorker() {
  return new Worker(new URL("./worker/index.ts", import.meta.url), { type: "module" });
}

export default function App() {
  const [route, setRoute] = useState<RouteState>(initialRoute);
  const [listFile, setListFile] = useState<File | null>(null);
  const [chatFile, setChatFile] = useState<File | null>(null);
  const [reminderListFile, setReminderListFile] = useState<File | null>(null);
  const [reminderAppealFile, setReminderAppealFile] = useState<File | null>(null);
  const [reminderSummaryFiles, setReminderSummaryFiles] = useState<File[]>([]);
  const [reminderChatFiles, setReminderChatFiles] = useState<File[]>([]);
  const [stageReportDenominatorFile, setStageReportDenominatorFile] = useState<File | null>(null);
  const [stageReportChatFiles, setStageReportChatFiles] = useState<File[]>([]);
  const [stageReportBeautifyFile, setStageReportBeautifyFile] = useState<File | null>(null);
  const [includeReminderChats, setIncludeReminderChats] = useState(false);
  const [includeReminderColors, setIncludeReminderColors] = useState(false);
  const [includeReminderExceptionSheet, setIncludeReminderExceptionSheet] = useState(false);
  const [includeReminderExplanationSheet, setIncludeReminderExplanationSheet] = useState(false);
  const [weekLabel, setWeekLabel] = useState<WeekLabel>("auto");
  const [useSingle, setUseSingle] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [statuses, setStatuses] = useState<Record<StatusKey, ProcessingStatus>>({
    checkin: INITIAL_STATUS,
    reminder: INITIAL_STATUS,
    stageReport: INITIAL_STATUS,
    stageReportPublish: INITIAL_STATUS,
  });
  const [activeModal, setActiveModal] = useState<ActiveModal>(null);
  const workerRef = useRef<Worker | null>(null);
  const { theme, usesSystemTheme, toggleTheme } = useTheme();

  useEffect(() => () => workerRef.current?.terminate(), []);

  useEffect(() => {
    const syncRoute = () => setRoute(routeFromLocation());
    window.addEventListener("popstate", syncRoute);
    window.addEventListener("hashchange", syncRoute);
    return () => {
      window.removeEventListener("popstate", syncRoute);
      window.removeEventListener("hashchange", syncRoute);
    };
  }, []);

  const activeMode = route.mode;
  const stageReportTab = route.stageReportTab;
  const activeStatusKey: StatusKey = activeMode === "stageReport" && stageReportTab === "publish"
    ? "stageReportPublish"
    : activeMode;

  useEffect(() => {
    document.title = activeMode === "reminder"
      ? "开课提醒触达完成率公示"
      : activeMode === "stageReport"
        ? "阶段性报告"
        : "打卡质检数据生成";
  }, [activeMode]);

  const weekHint = useMemo(() => {
    if (weekLabel !== "auto") return `已手动指定为${weekLabel}`;
    const inferred = listFile ? inferWeekFromFilename(listFile.name) : "";
    return inferred
      ? `根据文件日期预计为${inferred}，生成时会再用课次日期校验`
      : "选择名单后，将根据上课日期自动判断";
  }, [listFile, weekLabel]);

  function updateStatus(
    key: StatusKey,
    title: string,
    message: string,
    progress = 0,
    mode: ProcessingStatus["mode"] = "working",
  ) {
    setStatuses((current) => ({
      ...current,
      [key]: { visible: true, title, message, progress, mode },
    }));
  }

  function finishWorker() {
    setProcessing(false);
    workerRef.current?.terminate();
    workerRef.current = null;
  }

  function handleNavigationClick(event: MouseEvent<HTMLAnchorElement>) {
    event.preventDefault();
    const href = event.currentTarget.getAttribute("href");
    if (!href) return;
    const target = new URL(href, window.location.origin);
    const nextRoute: RouteState = target.pathname === "/remind"
      ? { mode: "reminder", stageReportTab: "check" }
      : target.pathname === "/report"
        ? {
            mode: "stageReport",
            stageReportTab: target.hash === "#publish" ? "publish" : "check",
          }
        : { mode: "checkin", stageReportTab: "check" };
    const nextUrl = routeUrl(nextRoute);
    setRoute(nextRoute);
    if (`${window.location.pathname}${window.location.hash}` !== nextUrl) {
      window.history.pushState(null, "", nextUrl);
    }
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!listFile || !chatFile) return;

    workerRef.current?.terminate();
    const worker = createProcessingWorker();
    workerRef.current = worker;
    setProcessing(true);
    updateStatus("checkin", "正在启动本地处理引擎", "所有文件只在当前浏览器中处理，不会上传。", 2);

    let whitelistCsv: string;
    try {
      whitelistCsv = await loadWhitelistCsv();
    } catch (error) {
      updateStatus("checkin", "处理失败", errorMessage(error), 100, "error");
      finishWorker();
      return;
    }

    worker.onmessage = ({ data }: MessageEvent<WorkerResponse>) => {
      if (data.type === "progress") {
        updateStatus("checkin", data.title, data.message, data.progress);
        return;
      }
      if (data.type === "complete") {
        downloadResult(data.buffer, data.filename);
        updateStatus(
          "checkin",
          "处理完成，结果已下载",
          `质检 ${data.summary.targets.toLocaleString()} 人：已发送 ${data.summary.sent.toLocaleString()}，未发送 ${data.summary.unsent.toLocaleString()}，免检 ${Number(data.summary.exempt || 0).toLocaleString()}；清洗后聊天 ${data.summary.cleanChats.toLocaleString()} 条。`,
          100,
          "done",
        );
        finishWorker();
        return;
      }
      updateStatus("checkin", "处理失败", data.message, 100, "error");
      finishWorker();
    };

    worker.onerror = (event) => {
      updateStatus("checkin", "处理失败", event.message || "浏览器工作线程发生错误。", 100, "error");
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
    updateStatus("reminder", "正在启动本地处理引擎", "所有文件只在当前浏览器中处理，不会上传。", 2);

    worker.onmessage = ({ data }: MessageEvent<WorkerResponse>) => {
      if (data.type === "progress") {
        updateStatus("reminder", data.title, data.message, data.progress);
        return;
      }
      if (data.type === "complete") {
        downloadResult(data.buffer, data.filename);
        updateStatus(
          "reminder",
          "处理完成，结果已下载",
          `开课提醒 ${data.summary.targets.toLocaleString()} 条：有效触达 ${data.summary.sent.toLocaleString()}，未触达 ${data.summary.unsent.toLocaleString()}，异常核对 ${Number(data.summary.exceptions || 0).toLocaleString()}；汇总文件 ${Number(data.summary.summaryFiles || 0).toLocaleString()} 个，聊天参考文件 ${Number(data.summary.chatFiles || 0).toLocaleString()} 个。`,
          100,
          "done",
        );
        finishWorker();
        return;
      }
      updateStatus("reminder", "处理失败", data.message, 100, "error");
      finishWorker();
    };

    worker.onerror = (event) => {
      updateStatus("reminder", "处理失败", event.message || "浏览器工作线程发生错误。", 100, "error");
      finishWorker();
    };

    let whitelistCsv: string;
    try {
      whitelistCsv = await loadWhitelistCsv();
    } catch (error) {
      updateStatus("reminder", "处理失败", errorMessage(error), 100, "error");
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

  async function handleStageReportSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!stageReportDenominatorFile || !stageReportChatFiles.length) return;
    workerRef.current?.terminate();
    const worker = createProcessingWorker();
    workerRef.current = worker;
    setProcessing(true);
    updateStatus("stageReport", "正在启动本地处理引擎", "所有文件只在当前浏览器中处理，不会上传。", 2);
    worker.onmessage = ({ data }: MessageEvent<WorkerResponse>) => {
      if (data.type === "progress") {
        updateStatus("stageReport", data.title, data.message, data.progress);
        return;
      }
      if (data.type === "complete") {
        downloadResult(data.buffer, data.filename);
        updateStatus(
          "stageReport",
          "处理完成，结果已下载",
          `阶段性报告 ${data.summary.targets.toLocaleString()} 条：已发送 ${data.summary.sent.toLocaleString()}，未发送 ${data.summary.unsent.toLocaleString()}；清洗后聊天 ${data.summary.cleanChats.toLocaleString()} 条。`,
          100,
          "done",
        );
        finishWorker();
        return;
      }
      updateStatus("stageReport", "处理失败", data.message, 100, "error");
      finishWorker();
    };
    worker.onerror = (event) => {
      updateStatus("stageReport", "处理失败", event.message || "浏览器工作线程发生错误。", 100, "error");
      finishWorker();
    };
    worker.postMessage({
      type: "process",
      mode: "stageReport",
      denominatorFile: stageReportDenominatorFile,
      chatFiles: stageReportChatFiles,
    });
  }

  async function handleStageReportBeautifySubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!stageReportBeautifyFile) return;
    workerRef.current?.terminate();
    const worker = createProcessingWorker();
    workerRef.current = worker;
    setProcessing(true);
    updateStatus("stageReportPublish", "正在启动本地处理引擎", "所有文件只在当前浏览器中处理，不会上传。", 2);
    worker.onmessage = ({ data }: MessageEvent<WorkerResponse>) => {
      if (data.type === "progress") {
        updateStatus("stageReportPublish", data.title, data.message, data.progress);
        return;
      }
      if (data.type === "complete") {
        downloadResult(data.buffer, data.filename);
        updateStatus(
          "stageReportPublish",
          "处理完成，结果已下载",
          `已整理 ${Number(data.summary.stageRows || 0).toLocaleString()} 条阶段性报告明细、${Number(data.summary.windowRows || 0).toLocaleString()} 条窗口期明细，${Number(data.summary.teacherRows || 0).toLocaleString()} 条教师汇总；输出 ${Number(data.summary.sheets || 0).toLocaleString()} 个 Sheet。`,
          100,
          "done",
        );
        finishWorker();
        return;
      }
      updateStatus("stageReportPublish", "处理失败", data.message, 100, "error");
      finishWorker();
    };
    worker.onerror = (event) => {
      updateStatus("stageReportPublish", "处理失败", event.message || "浏览器工作线程发生错误。", 100, "error");
      finishWorker();
    };
    worker.postMessage({
      type: "process",
      mode: "stageReportBeautify",
      sourceFile: stageReportBeautifyFile,
    });
  }

  const navigationLinks = [
    { href: "/", label: "打卡质检", active: activeMode === "checkin" },
    { href: "/remind", label: "开课提醒", active: activeMode === "reminder" },
    { href: "/report", label: "阶段性报告", active: activeMode === "stageReport" },
  ];

  return (
    <>
      <main className="shell">
        <Header
          theme={theme}
          usesSystemTheme={usesSystemTheme}
          title={activeMode === "reminder" ? "开课提醒触达完成率公示" : activeMode === "stageReport" ? "阶段性报告" : "打卡质检数据生成"}
          subtitle={
            activeMode === "reminder"
              ? "上传开课提醒学员明细名单与聊天质检汇总文件，在浏览器本地计算教师及以上维度触达完成率。文件不会上传服务器。"
              : activeMode === "stageReport"
                ? "核验阶段性报告发送情况，或将窗口期与非窗口期原始表单整理为统一公示版。文件不会上传服务器。"
                : undefined
          }
          showGuide={activeMode === "checkin"}
          navigationLinks={navigationLinks}
          onToggleTheme={toggleTheme}
          onOpenGuide={() => setActiveModal("guide")}
          onOpenChangelog={() => setActiveModal("changelog")}
          onNavigationClick={handleNavigationClick}
        />
        {activeMode === "stageReport" ? (
          <nav className="stage-report-tabs" aria-label="阶段性报告功能">
            <a
              className={stageReportTab === "check" ? "active" : undefined}
              href="/report"
              aria-current={stageReportTab === "check" ? "page" : undefined}
              onClick={handleNavigationClick}
            >
              <strong>发送检查</strong>
              <small>分母＋聊天记录</small>
            </a>
            <a
              className={stageReportTab === "publish" ? "active" : undefined}
              href="/report#publish"
              aria-current={stageReportTab === "publish" ? "page" : undefined}
              onClick={handleNavigationClick}
            >
              <strong>公示版整理</strong>
              <small>窗口期＋非窗口期原始表单</small>
            </a>
          </nav>
        ) : null}
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
        ) : activeMode === "stageReport" && stageReportTab === "check" ? (
          <StageReportForm
            denominatorFile={stageReportDenominatorFile}
            chatFiles={stageReportChatFiles}
            processing={processing}
            onDenominatorFileChange={setStageReportDenominatorFile}
            onChatFilesChange={setStageReportChatFiles}
            onSubmit={handleStageReportSubmit}
          />
        ) : activeMode === "stageReport" ? (
          <StageReportBeautifyForm
            sourceFile={stageReportBeautifyFile}
            processing={processing}
            onSourceFileChange={setStageReportBeautifyFile}
            onSubmit={handleStageReportBeautifySubmit}
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
        <StatusCard status={statuses[activeStatusKey]} />
        <OutputGrid mode={activeMode === "stageReport" && stageReportTab === "publish" ? "stageReportBeautify" : activeMode} />
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
