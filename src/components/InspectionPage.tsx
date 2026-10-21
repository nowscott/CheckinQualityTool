import { useEffect, useRef, useState } from "react";
import { downloadResult } from "../lib/download";
import { loadInspectionWorkerUrl } from "../lib/inspectionWorker";
import { loadLatestInspectionRoster, type InspectionRosterSnapshot, type RosterSnapshotMeta } from "../lib/rosterSnapshot";
import inspectionWorkerUrl from "../worker/inspectionEntry.ts?worker&url";
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

async function createProcessingWorker() {
  const url = await loadInspectionWorkerUrl(import.meta.env.DEV ? inspectionWorkerUrl : undefined);
  return new Worker(url, { type: "module" });
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

async function fetchCoverage(month: string, signal?: AbortSignal) {
  const response = await fetch(`/api/inspection-coverage?month=${encodeURIComponent(month)}&ts=${Date.now()}`, { cache: "no-store", signal });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(body.error || "抽检登记读取失败。");
  if (body.month !== month || !Array.isArray(body.teachers) || typeof body.sourceUrl !== "string"
    || !body.teachers.every((teacher: { teacherName?: unknown; teacherEmail?: unknown }) => teacher && typeof teacher.teacherName === "string" && typeof teacher.teacherEmail === "string")) throw new Error("抽检登记返回格式无效。");
  return {
    teachers: body.teachers as { teacherEmail: string; teacherName: string }[],
    sourceUrl: body.sourceUrl,
    summary: `${body.month} 已登记 ${body.rowCount} 条课程，去重后 ${body.teachers.length} 位教师。${body.undatedRowCount ? ` ${body.undatedRowCount} 行未填抽检日期，暂不计入覆盖，请补填。` : ""}${body.invalidDateRowCount ? ` ${body.invalidDateRowCount} 行抽检日期无效（行号：${body.invalidDateRows.join("、")}），暂不计入覆盖，请修正为 YYYY-MM-DD。` : ""}`,
  };
}

export function InspectionPage() {
  const [feedbackFile, setFeedbackFile] = useState<File | null>(null);
  const [rosterFile, setRosterFile] = useState<File | null>(null);
  const [coverageMonth, setCoverageMonth] = useState(() => new Intl.DateTimeFormat("sv-SE", { timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit" }).format(new Date()));
  const [coverageUrl, setCoverageUrl] = useState("https://docs.qq.com/sheet/DUGFUT3dOckJVeXFm?tab=ih60c3");
  const [coverageSummary, setCoverageSummary] = useState("");
  const [sampleCount, setSampleCount] = useState("1000");
  const [includeExplanation, setIncludeExplanation] = useState(false);
  const [includeReason, setIncludeReason] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [usedRoster, setUsedRoster] = useState<RosterSnapshotMeta | null>(null);
  const [status, setStatus] = useState<ProcessingStatus>(INITIAL_STATUS);
  const workerRef = useRef<Worker | null>(null);

  useEffect(() => () => workerRef.current?.terminate(), []);

  useEffect(() => {
    const controller = new AbortController();
    setCoverageUrl("https://docs.qq.com/sheet/DUGFUT3dOckJVeXFm?tab=ih60c3");
    setCoverageSummary("正在按抽检日期读取所选月份的登记…");
    void fetchCoverage(coverageMonth, controller.signal).then((body) => {
      setCoverageUrl(body.sourceUrl);
      setCoverageSummary(body.summary);
    }).catch((error) => {
      if (!controller.signal.aborted) setCoverageSummary(errorMessage(error));
    });
    return () => controller.abort();
  }, [coverageMonth]);

  function updateStatus(title: string, message: string, progress = 0, mode: ProcessingStatus["mode"] = "working") {
    setStatus({ visible: true, title, message, progress, mode });
  }

  function finishWorker() {
    setProcessing(false);
    workerRef.current?.terminate();
    workerRef.current = null;
  }

  async function loadCoverage() {
    const body = await fetchCoverage(coverageMonth);
    setCoverageUrl(body.sourceUrl);
    setCoverageSummary(body.summary);
    return body.teachers;
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
    const priorityLabel = priorityMode === "monthly" ? "优先本月未覆盖教师" : priorityMode === "unreported"
      ? "优先覆盖本月未反馈教师"
      : "优先全覆盖教师，剩余名额为未反馈教师加频";
    updateStatus(priorityLabel, "正在准备抽检数据。", 2);

    // Capture the user-selected file and service version before any other asynchronous work.
    const selectedRosterFile = rosterFile;
    let defaultRoster: InspectionRosterSnapshot | undefined;
    setUsedRoster(null);
    if (!selectedRosterFile) {
      try {
        updateStatus("正在读取最新在职名单", "本次处理固定使用读取到的快照；读取失败会停止。", 3);
        defaultRoster = await loadLatestInspectionRoster();
        setUsedRoster(defaultRoster);
      } catch (error) {
        updateStatus("读取最新在职名单失败", errorMessage(error), 100, "error");
        setProcessing(false);
        return;
      }
    }

    let coverageSourceSummary = "";
    let inspectedTeachers: { teacherEmail: string; teacherName: string }[] = [];
    if (priorityMode === "monthly") {
      try {
        const coverage = await fetchCoverage(coverageMonth);
        inspectedTeachers = coverage.teachers;
        coverageSourceSummary = coverage.summary;
        setCoverageUrl(coverage.sourceUrl);
        setCoverageSummary(coverage.summary);
      } catch (error) {
        updateStatus("读取抽检登记失败", errorMessage(error), 100, "error");
        setProcessing(false);
        return;
      }
    }
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

    let worker: Worker;
    try {
      worker = await createProcessingWorker();
    } catch (error) {
      updateStatus("加载抽检程序失败", errorMessage(error), 100, "error");
      setProcessing(false);
      return;
    }
    workerRef.current = worker;
    worker.onmessage = ({ data }: MessageEvent<WorkerResponse>) => {
      if (data.type === "progress") {
        updateStatus(data.title, data.message, data.progress);
        return;
      }
      if (data.type === "inspectionComplete") {
        const complete = data as InspectionWorkerComplete;
        setUsedRoster(complete.rosterSnapshot);
        downloadResult(complete.chunks, complete.filename);
        const prioritySummary = complete.priority;
        const coverageMessage = prioritySummary.mode === "monthly"
          ? `；${prioritySummary.coverageMonth} 台账登记 ${prioritySummary.registeredTeacherCount} 位，本次候选已覆盖 ${prioritySummary.matchedCoveredTeacherCount} 位、未覆盖 ${prioritySummary.uncoveredTeacherCount} 位`
          : "";
        const focusMessage = prioritySummary.focusTeacherCount
          ? `；未反馈教师 ${prioritySummary.focusTeacherCount} 位，匹配 ${prioritySummary.matchedFocusTeacherCount} 位，未匹配 ${prioritySummary.unmatchedFocusTeacherCount} 位，同名冲突 ${prioritySummary.ambiguousFocusTeacherCount} 位`
          : "";
        const focusExtraMessage = complete.summary.focusTeacherExtraRows
          ? `；未反馈教师加频 ${complete.summary.focusTeacherExtraRows} 条`
          : "";
        const supervisorUnsubmittedMessage = complete.summary.excludedSupervisorUnsubmittedRows
          ? `；主管及助理主管未生成报告未加抽 ${complete.summary.excludedSupervisorUnsubmittedRows} 条`
          : "";
        const scoreMessage = `；评分 ${prioritySummary.matchedScoreTeacherCount}/${prioritySummary.scoreSourceTeacherCount} 位教师匹配，缺分 ${prioritySummary.missingScoreTeacherCount} 位，身份歧义 ${prioritySummary.ambiguousScoreTeacherCount} 位`;
        updateStatus(
          "处理完成，结果已下载",
          `在职快照 ${complete.rosterSnapshot.snapshotDate || "上传文件未识别日期"}；候选 ${complete.summary.eligibleTeachers.toLocaleString()} 位教师、${complete.summary.eligibleRows.toLocaleString()} 条课程；排除经理 ${complete.summary.excludedManagementTeachers.toLocaleString()} 位教师、${complete.summary.excludedManagementRows.toLocaleString()} 条课程；抽检 ${complete.summary.selectedRows.toLocaleString()} 条，命中 ${complete.summary.selectedTeachers.toLocaleString()} 位教师${coverageMessage}${focusMessage}${focusExtraMessage}${supervisorUnsubmittedMessage}${scoreMessage}。${coverageSourceSummary}${focusSourceSummary}`,
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
      rosterFile: selectedRosterFile || undefined,
      defaultRoster,
      sampleCount: count,
      attempt: 1,
      includeExplanation,
      includeReason,
      priorityMode,
      focusTeacherNames,
      inspectedTeachers,
      coverageMonth,
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
            title="上传指定版本在职教师明细"
            description="不上传时，每次处理读取统一接口最新名单；历史报告可上传对应日期的明细。上传表需含岗位描述字段。经理完全排除；主管、助理主管可参加普通抽检，但未生成报告不加抽。"
            file={rosterFile}
            required={false}
            onChange={setRosterFile}
          />
        </div>
        <div className="inspection-local-note">
          <h3>本月抽检覆盖</h3>
          <p>完成抽检后，由负责老师在 <a href={coverageUrl} target="_blank" rel="noreferrer">腾讯文档抽检登记表</a> 从抽检文件直接复制前四列追加粘贴，可带表头，勾选入选原因时第五列也可一起粘贴。粘贴后在第六列「抽检日期」批量填写实际完成日期（如 2026-10-21）。系统按抽检日期所属月份统计，同一教师多条课程自动去重；日期未填或无效的记录暂不计入覆盖并提示补填。所有月份使用同一张表，历史记录无需清空。</p>
          <label>抽检月份 <input type="month" value={coverageMonth} disabled={processing} onChange={(event) => { setCoverageMonth(event.target.value); setCoverageSummary(""); setCoverageUrl("https://docs.qq.com/sheet/DUGFUT3dOckJVeXFm?tab=ih60c3"); }} /></label>
          <button type="button" className="secondary" disabled={processing || !coverageMonth} onClick={() => { void loadCoverage().catch((error) => setCoverageSummary(errorMessage(error))); }}>读取已抽检人数</button>
          {coverageSummary ? <p role="status">{coverageSummary}</p> : null}
          <p>未覆盖优先：先抽当月未登记的教师，再覆盖已登记教师，剩余名额按评分排序补足。仅针对本次课程明细中的在职教师；生成名单后仍需人工登记实际完成情况。未生成报告的容量外加抽规则继续保留。</p>
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
            checked={includeReason}
            onChange={(event) => setIncludeReason(event.target.checked)}
          />
          附带入选原因列
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
          <button type="button" disabled={!feedbackFile || processing || !coverageMonth} onClick={() => void exportSelection("monthly")}>
            {processing ? "正在生成…" : "导出：优先本月未覆盖教师"}
          </button>
          <button type="button" disabled={!feedbackFile || processing} onClick={() => void exportSelection("coverage")}>
            {processing ? "正在生成…" : "导出：优先全覆盖教师"}
          </button>
          <button type="button" className="secondary" disabled={!feedbackFile || processing} onClick={() => void exportSelection("unreported")}>
            {processing ? "正在生成…" : "导出：优先本月未反馈教师"}
          </button>
        </div>
        <p className="inspection-local-note">
          名单默认包含教师姓名、学员姓名、课次日期和课次时间；“入选原因”可选附带。经理完全排除；主管、助理主管可参加普通抽检，但其未生成报告课程不做容量外加抽。三种优先方式均读取教学服务赋分排序及未反馈名单；月度覆盖模式另读取所选月份的抽检登记。按教师姓名匹配，重名时使用教研组、师训组长消歧。课程与教师文件只在当前浏览器处理，不上传到服务器。
        </p>
      </section>
      {usedRoster ? <p className="inspection-local-note" title={`源文件 SHA-256：${usedRoster.sourceSha256}`}>
        本次在职快照：{usedRoster.snapshotDate || "上传文件未识别日期"} · {usedRoster.rowCount.toLocaleString()} 行 · {usedRoster.sourceFile}
      </p> : null}
      <StatusCard status={status} />
    </>
  );
}
