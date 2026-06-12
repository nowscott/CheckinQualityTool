import { useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import { ChangelogDialog } from "./components/ChangelogDialog";
import { Header } from "./components/Header";
import { MatchingGuideDialog } from "./components/MatchingGuideDialog";
import { OutputGrid } from "./components/OutputGrid";
import { StatusCard } from "./components/StatusCard";
import { UploadForm } from "./components/UploadForm";
import { downloadResult } from "./lib/download";
import { inferWeekFromFilename } from "./lib/week";
import type { ProcessingStatus, WeekLabel, WorkerResponse } from "./types/worker";

type ActiveModal = "guide" | "changelog" | null;

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

export default function App() {
  const [listFile, setListFile] = useState<File | null>(null);
  const [chatFile, setChatFile] = useState<File | null>(null);
  const [weekLabel, setWeekLabel] = useState<WeekLabel>("auto");
  const [useSingle, setUseSingle] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [status, setStatus] = useState(INITIAL_STATUS);
  const [activeModal, setActiveModal] = useState<ActiveModal>(null);
  const workerRef = useRef<Worker | null>(null);

  useEffect(() => () => workerRef.current?.terminate(), []);

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
    const worker = new Worker("/worker.js");
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
          `质检 ${data.summary.targets.toLocaleString()} 人：已发送 ${data.summary.sent.toLocaleString()}，未发送 ${data.summary.unsent.toLocaleString()}，免检 ${data.summary.exempt.toLocaleString()}；清洗后聊天 ${data.summary.cleanChats.toLocaleString()} 条。`,
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

  return (
    <>
      <main className="shell">
        <Header
          onOpenGuide={() => setActiveModal("guide")}
          onOpenChangelog={() => setActiveModal("changelog")}
        />
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
