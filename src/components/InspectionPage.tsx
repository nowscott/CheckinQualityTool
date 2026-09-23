import { useEffect, useRef, useState } from "react";
import { downloadResult } from "../lib/download";
import { UploadCard } from "./UploadCard";
import { StatusCard } from "./StatusCard";
import type { InspectionPriorityMode, InspectionTeacherScore } from "../worker/inspectionTypes";
import type { InspectionWorkerComplete, ProcessingStatus, WorkerResponse } from "../types/worker";

interface FeedbackFocusResponse {
  error?: string;
  rowCount?: number;
  unreportedRowCount?: number;
  unreportedTeacherNames?: string[];
  missingNameRows?: number;
}

interface TeacherScoresResponse {
  error?: string;
  rowCount?: number;
  missingNameRows?: number;
  missingScoreRows?: number;
  rows?: InspectionTeacherScore[];
}

const INITIAL_STATUS: ProcessingStatus = {
  visible: false,
  title: "正在生成抽检名单",
  message: "Excel 文件在当前浏览器中处理，请不要关闭页面。",
  progress: 0,
  mode: "working",
};

function createProcessingWorker() {
  return new Worker(new URL("../worker/index.ts", import.meta.url), { type: "module" });
}

function errorMessage(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}

async function loadUnreportedTeacherNames() {
  const response = await fetch(`/api/feedback-focus?ts=${Date.now()}`, { cache: "no-store" });
  const body = (await response.json().catch(() => ({}))) as FeedbackFocusResponse;
  if (!response.ok) throw new Error(body.error || "本月未反馈重点关注表读取失败。");
  if (!Array.isArray(body.unreportedTeacherNames)) throw new Error("重点关注表返回格式无效。");
  return {
    names: body.unreportedTeacherNames,
    rowCount: Number(body.rowCount || 0),
    unreportedRowCount: Number(body.unreportedRowCount || 0),
    missingNameRows: Number(body.missingNameRows || 0),
  };
}

async function loadTeacherScores() {
  const response = await fetch(`/api/teacher-scores?ts=${Date.now()}`, { cache: "no-store" });
  const body = (await response.json().catch(() => ({}))) as TeacherScoresResponse;
  if (!response.ok) throw new Error(body.error || "教学服务评分页读取失败。");
  if (!Array.isArray(body.rows) || !body.rows.length) throw new Error("教学服务评分页没有可用的评分记录。");
  return {
    rows: body.rows,
    rowCount: Number(body.rowCount || 0),
    missingNameRows: Number(body.missingNameRows || 0),
    missingScoreRows: Number(body.missingScoreRows || 0),
  };
}

export function InspectionPage() {
  const [feedbackFile, setFeedbackFile] = useState<File | null>(null);
  const [rosterFile, setRosterFile] = useState<File | null>(null);
  const [sampleCount, setSampleCount] = useState("1000");
  const [includeExplanation, setIncludeExplanation] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [status, setStatus] = useState<ProcessingStatus>(INITIAL_STATUS);
  const workerRef = useRef<Worker | null>(null);

  useEffect(() => () => workerRef.current?.terminate(), []);

  function updateStatus(title: string, message: string, progress = 0, mode: ProcessingStatus["mode"] = "working") {
    setStatus({ visible: true, title, message, progress, mode });
  }

  function finishWorker() {
    setProcessing(false);
    workerRef.current?.terminate();
    workerRef.current = null;
  }

  async function exportSelection(priorityMode: InspectionPriorityMode) {
    if (!feedbackFile) return;
    const count = Number(sampleCount);
    if (!Number.isInteger(count) || count < 1) {
      updateStatus("抽检数无效", "请输入大于 0 的整数。", 100, "error");
      return;
    }

    workerRef.current?.terminate();
    workerRef.current = null;
    setProcessing(true);
    const priorityLabel = priorityMode === "unreported"
      ? "优先覆盖本月未反馈教师"
      : "优先全覆盖教师，剩余名额为未反馈教师加频";
    updateStatus(priorityLabel, "正在准备抽检数据。", 2);

    let focusTeacherNames: string[] = [];
    let teacherScoreRows: InspectionTeacherScore[];
    let focusSourceSummary = "";
    try {
      updateStatus(
        "正在读取腾讯文档",
        priorityMode === "coverage"
          ? "读取未反馈教师名单，以便在全覆盖后给未反馈教师加频。"
          : "读取“反馈抽检重点关注”页的本月未反馈教师。",
        5,
      );
      const focus = await loadUnreportedTeacherNames();
      focusTeacherNames = focus.names;
      focusSourceSummary = `重点关注表 ${focus.rowCount} 行，其中未反馈 ${focus.unreportedRowCount} 行、${focus.names.length} 位教师。`;
      if (priorityMode === "coverage") focusSourceSummary += " 覆盖教师后，剩余名额优先用于未反馈教师加频。";
      if (focus.missingNameRows) focusSourceSummary += ` ${focus.missingNameRows} 行缺少教师姓名，已跳过。`;
    } catch (error) {
      if (priorityMode === "unreported") {
        updateStatus("读取未反馈名单失败", errorMessage(error), 100, "error");
        setProcessing(false);
        return;
      }
      focusSourceSummary = `重点关注表读取失败，本次覆盖优先导出将按稳定排序补足剩余名额，不进行未反馈教师加频。${errorMessage(error)}`;
    }

    try {
      updateStatus("正在读取评分来源", "按教师姓名及组织信息匹配评分排序；只取优先顺序，不读取原始分值。经理岗位继续排除。", 9);
      const scores = await loadTeacherScores();
      teacherScoreRows = scores.rows;
      focusSourceSummary += ` 评分源 ${scores.rowCount} 行，读取到 ${scores.rows.length} 位教师赋分。`;
      if (scores.missingNameRows) focusSourceSummary += ` ${scores.missingNameRows} 行缺少教师姓名，已跳过。`;
      if (scores.missingScoreRows) focusSourceSummary += ` ${scores.missingScoreRows} 行缺少有效赋分，已跳过。`;
    } catch (error) {
      updateStatus("读取评分失败", errorMessage(error), 100, "error");
      setProcessing(false);
      return;
    }

    const worker = createProcessingWorker();
    workerRef.current = worker;
    worker.onmessage = ({ data }: MessageEvent<WorkerResponse>) => {
      if (data.type === "progress") {
        updateStatus(data.title, data.message, data.progress);
        return;
      }
      if (data.type === "inspectionComplete") {
        const complete = data as InspectionWorkerComplete;
        downloadResult(complete.chunks, complete.filename);
        const prioritySummary = complete.priority;
        const focusMessage = prioritySummary.focusTeacherCount
          ? `；未反馈教师 ${prioritySummary.focusTeacherCount} 位，匹配 ${prioritySummary.matchedFocusTeacherCount} 位，未匹配 ${prioritySummary.unmatchedFocusTeacherCount} 位，同名冲突 ${prioritySummary.ambiguousFocusTeacherCount} 位`
          : "";
        const focusExtraMessage = complete.summary.focusTeacherExtraRows
          ? `；未反馈教师加频 ${complete.summary.focusTeacherExtraRows} 条`
          : "";
        const scoreMessage = `；评分 ${prioritySummary.matchedScoreTeacherCount}/${prioritySummary.scoreSourceTeacherCount} 位教师匹配，缺分 ${prioritySummary.missingScoreTeacherCount} 位，身份歧义 ${prioritySummary.ambiguousScoreTeacherCount} 位`;
        updateStatus(
          "处理完成，结果已下载",
          `候选 ${complete.summary.eligibleTeachers.toLocaleString()} 位教师、${complete.summary.eligibleRows.toLocaleString()} 条课程；排除管理岗位 ${complete.summary.excludedManagementTeachers.toLocaleString()} 位教师、${complete.summary.excludedManagementRows.toLocaleString()} 条课程；抽检 ${complete.summary.selectedRows.toLocaleString()} 条，命中 ${complete.summary.selectedTeachers.toLocaleString()} 位教师${focusMessage}${focusExtraMessage}${scoreMessage}。${focusSourceSummary}`,
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
      mode: "inspection",
      feedbackFile,
      rosterFile: rosterFile || undefined,
      sampleCount: count,
      attempt: 1,
      includeExplanation,
      priorityMode,
      focusTeacherNames,
      teacherScoreRows,
    });
  }

  return (
    <>
      <section className="card inspection-card">
        <div className="grid">
          <UploadCard
            id="inspection-feedback-file"
            name="inspection_feedback_file"
            step="01 / 必需"
            title="上传本周课堂反馈明细"
            description="读取课次、教师、学员与报告生成状态。"
            file={feedbackFile}
            onChange={setFeedbackFile}
          />
          <UploadCard
            id="inspection-roster-file"
            name="inspection_roster_file"
            step="02 / 可选"
            title="上传最新在职教师明细"
            description="不上传时使用 2026-09-22 内置名单；上传表需含岗位描述字段。经理岗位排除，主管仍参与抽检。"
            file={rosterFile}
            required={false}
            onChange={setRosterFile}
          />
        </div>
        <label className="inspection-count-field" htmlFor="inspection-sample-count">
          抽检课程条数
          <input
            id="inspection-sample-count"
            className="number-input"
            type="number"
            min="1"
            step="1"
            value={sampleCount}
            onChange={(event) => setSampleCount(event.target.value)}
          />
        </label>
        <label className="inspection-explanation-option">
          <input
            type="checkbox"
            checked={includeExplanation}
            onChange={(event) => setIncludeExplanation(event.target.checked)}
          />
          同时导出处理说明页
        </label>
        <div className="inspection-export-actions">
          <button type="button" disabled={!feedbackFile || processing} onClick={() => void exportSelection("coverage")}>
            {processing ? "正在生成…" : "导出：优先全覆盖教师"}
          </button>
          <button type="button" className="secondary" disabled={!feedbackFile || processing} onClick={() => void exportSelection("unreported")}>
            {processing ? "正在生成…" : "导出：优先本月未反馈教师"}
          </button>
        </div>
        <p className="inspection-local-note">
          仅排除岗位描述含“经理”的教师，主管仍参与抽检。两种优先方式均从腾讯文档读取本月未反馈名单和教学服务赋分排序；按教师姓名匹配，重名时使用教研组、师训组长消歧。课程与教师文件只在当前浏览器处理，不上传到服务器。
        </p>
      </section>
      <StatusCard status={status} />
    </>
  );
}
