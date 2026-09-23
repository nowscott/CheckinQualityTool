import { useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import { ChangelogDialog } from "./components/ChangelogDialog";
import { Header } from "./components/Header";
import { MatchingGuideDialog } from "./components/MatchingGuideDialog";
import { OutputGrid } from "./components/OutputGrid";
import { StatusCard } from "./components/StatusCard";
import { UploadForm } from "./components/UploadForm";
import { InspectionPage } from "./components/InspectionPage";
import { downloadResult } from "./lib/download";
import { inferWeekFromFilename } from "./lib/week";
import { useTheme } from "./hooks/useTheme";
import type { ProcessingStatus, WeekLabel, WorkerResponse } from "./types/worker";

type ActiveModal = "guide" | "changelog" | null;
type ActivePage = "checkin" | "inspection";

const INITIAL_STATUS: ProcessingStatus = {
  visible: false,
  title: "正在处理数据",
  message: "大文件需要一些时间，请不要关闭页面。",
  progress: 0,
  mode: "working",
};

function loadWhitelistCsv() {
  return fetch(`/api/whitelist?ts=${Date.now()}`, { cache: "no-store" }).then(async (response) => {
    if (!response.ok) {
      let message = "在线白名单读取失败，请稍后重试。";
      try {
        const body = (await response.json()) as { error?: string };
        if (body.error) message = body.error;
      } catch {
        // Keep the stable user-facing error when the upstream response is not JSON.
      }
      throw new Error(message);
    }
    return response.text();
  });
}

function errorMessage(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}

function createProcessingWorker() {
  return new Worker(new URL("./worker/index.ts", import.meta.url), { type: "module" });
}

export default function App() {
  const [listFile, setListFile] = useState<File | null>(null);
  const [chatFile, setChatFile] = useState<File | null>(null);
  const [weekLabel, setWeekLabel] = useState<WeekLabel>("auto");
  const [useSingle, setUseSingle] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [status, setStatus] = useState<ProcessingStatus>(INITIAL_STATUS);
  const [activeModal, setActiveModal] = useState<ActiveModal>(null);
  const [activePage, setActivePage] = useState<ActivePage>("checkin");
  const workerRef = useRef<Worker | null>(null);
  const { theme, usesSystemTheme, toggleTheme } = useTheme();

  useEffect(() => {
    document.title = activePage === "inspection" ? "课堂反馈抽检" : "打卡质检数据生成";
    return () => workerRef.current?.terminate();
  }, [activePage]);

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

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!listFile || !chatFile) return;

    workerRef.current?.terminate();
    const worker = createProcessingWorker();
    workerRef.current = worker;
    setProcessing(true);
    updateStatus("正在读取在线白名单", "读取腾讯文档后，Excel 文件仍只在当前浏览器中处理。", 2);

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
        downloadResult(data.chunks, data.filename);
        updateStatus(
          "处理完成，结果已下载",
          `质检 ${data.summary.targets.toLocaleString()} 人：已发送 ${data.summary.sent.toLocaleString()}，未发送 ${data.summary.unsent.toLocaleString()}，免检 ${Number(data.summary.exempt || 0).toLocaleString()}；清洗后聊天 ${data.summary.cleanChats.toLocaleString()} 条。`,
          100,
          "done",
        );
        finishWorker();
        return;
      }
      if (data.type !== "error") return;
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

  return (
    <>
      <main className="shell">
        <Header
          title={activePage === "inspection" ? "课堂反馈抽检" : undefined}
          subtitle={activePage === "inspection" ? "主管、经理和岗位纠偏名单人员排除；支持教师覆盖优先，或按腾讯文档中的本月未反馈教师优先导出。" : undefined}
          showGuide={activePage === "checkin"}
          theme={theme}
          usesSystemTheme={usesSystemTheme}
          onToggleTheme={toggleTheme}
          onOpenGuide={() => setActiveModal("guide")}
          onOpenChangelog={() => setActiveModal("changelog")}
        />
        <nav className="mode-tabs" aria-label="数据工具页面">
          <button type="button" className={activePage === "checkin" ? "active" : ""} onClick={() => setActivePage("checkin")}>打卡质检</button>
          <button type="button" className={activePage === "inspection" ? "active" : ""} onClick={() => setActivePage("inspection")}>课堂反馈抽检</button>
        </nav>
        {activePage === "inspection" ? (
          <InspectionPage />
        ) : (
          <>
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
            <StatusCard status={status} />
            <OutputGrid />
          </>
        )}
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
