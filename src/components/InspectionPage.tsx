import { useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import { downloadResult } from "../lib/download";
import {
  clearInspectionAuthCookie,
  INSPECTION_AUTH_REVALIDATION_MS,
  isInspectionAuthCookieStale,
  readInspectionAuthCookie,
  type InspectionAuthCookie,
  type InspectionAuthUserCookie,
} from "../lib/inspectionAuthCookie";
import { UploadCard } from "./UploadCard";
import { StatusCard } from "./StatusCard";
import { buildMonthlyInspectionPlan, monthWeekCount } from "../worker/monthlyPlanner";
import type {
  InspectionBatchKind,
  InspectionHistoryPayload,
  InspectionStats,
  MonthlyInspectionData,
  MonthlyInspectionPlan,
  TeachingServiceSnapshot,
} from "../worker/inspectionTypes";
import type { InspectionWorkerComplete, ProcessingStatus, WorkerResponse } from "../types/worker";

type SaveAction = "create" | "reuse" | "replace";
type HistoryView = "teachers" | "batches" | "monthly" | "users";
type HistoryFilterKind = "formal" | "trial" | "all";
type HistoryFilterStatus = "active" | "voided" | "all";

interface HistoryFiltersForRequest {
  batchKind: HistoryFilterKind;
  status: HistoryFilterStatus;
  from: string;
  to: string;
}

interface BatchSummary {
  id: string;
  businessWeekStart: string;
  businessWeekEnd: string;
  sourceName: string;
  rosterName: string;
  sourceSha256: string;
  attempt: number;
  sampleLimit: number;
  selectedCount: number;
  teacherCount: number;
  unsubmittedCount: number;
  createdAt: string;
  status: string;
  voidedAt?: string | null;
  batchKind: InspectionBatchKind;
}

interface HistoryItem {
  position: number;
  teacherName: string;
  teacherEmail: string;
  studentName: string;
  studentId: string;
  courseId: string;
  lessonStart: string;
  lessonEnd: string;
  submittedValue: string;
  productGroup: string;
  campus: string;
  projectGroup: string;
  selectionReason: string;
}

interface HistoryDetail extends BatchSummary {
  items: HistoryItem[];
  total: number;
  page: number;
  pageSize: number;
}

interface TeacherSummary {
  teacherKey: string;
  teacherName: string;
  teacherEmail: string;
  courseCount: number;
  batchCount: number;
  weekCount: number;
  unsubmittedCount: number;
  formalCourseCount: number;
  trialCourseCount: number;
  lastWeek: string | null;
  projectGroups: string;
  productGroups: string;
}

interface TeacherDetailItem extends HistoryItem {
  batchId: string;
  businessWeekStart: string;
  businessWeekEnd: string;
  attempt: number;
  status: string;
  batchKind: InspectionBatchKind;
}

interface TeacherWeek {
  batchId: string;
  businessWeekStart: string;
  businessWeekEnd: string;
  attempt: number;
  batchKind: InspectionBatchKind;
  status: string;
  courseCount: number;
  unsubmittedCount: number;
}

interface TeacherDetailPayload {
  teacher: {
    teacherKey: string;
    teacherName: string;
    teacherEmail: string;
    courseCount: number;
    batchCount: number;
    weekCount: number;
    unsubmittedCount: number;
  };
  weeks: TeacherWeek[];
  items: TeacherDetailItem[];
  total: number;
  page: number;
  pageSize: number;
}

type AuthUser = InspectionAuthUserCookie;

interface ManagedUser extends AuthUser {
  isActive: boolean;
  createdAt: string;
}

interface AuditEntry {
  id: string;
  action: string;
  metadata: Record<string, unknown>;
  createdAt: string;
  actorUsername: string;
  targetUsername: string;
}

interface SaveResponse {
  reused?: boolean;
  conflict?: boolean;
  batch: BatchSummary;
}

const INITIAL_STATUS: ProcessingStatus = {
  visible: false,
  title: "正在生成抽检名单",
  message: "大文件需要一些时间，请不要关闭页面。",
  progress: 0,
  mode: "working",
};

function currentMonthLabel() {
  const date = new Date();
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function monthEndLabel(month: string) {
  const [year, monthNumber] = month.split("-").map(Number);
  const date = new Date(Date.UTC(year, monthNumber, 0));
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function createProcessingWorker() {
  return new Worker(new URL("../worker/index.ts", import.meta.url), { type: "module" });
}

function errorMessage(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}

async function responseJson(response: Response) {
  try {
    return (await response.json()) as Record<string, unknown>;
  } catch {
    return {};
  }
}

export function InspectionPage() {
  const [initialAuthCookie] = useState<InspectionAuthCookie | null>(() => readInspectionAuthCookie());
  const [feedbackFile, setFeedbackFile] = useState<File | null>(null);
  const [rosterFile, setRosterFile] = useState<File | null>(null);
  const [sampleCount, setSampleCount] = useState("200");
  const [includeExplanation, setIncludeExplanation] = useState(false);
  const [batchKind, setBatchKind] = useState<InspectionBatchKind>("formal");
  const [authUser, setAuthUser] = useState<AuthUser | null>(() => initialAuthCookie?.user || null);
  const [authChecked, setAuthChecked] = useState(() => Boolean(initialAuthCookie));
  const [loginUsername, setLoginUsername] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginBusy, setLoginBusy] = useState(false);
  const [authenticated, setAuthenticated] = useState(() => Boolean(initialAuthCookie));
  const [processing, setProcessing] = useState(false);
  const [status, setStatus] = useState<ProcessingStatus>(INITIAL_STATUS);
  const [summary, setSummary] = useState<InspectionStats | null>(null);
  const [activeBatch, setActiveBatch] = useState<BatchSummary | null>(null);
  const [pendingReplacement, setPendingReplacement] = useState<BatchSummary | null>(null);
  const [history, setHistory] = useState<BatchSummary[]>([]);
  const [historyView, setHistoryView] = useState<HistoryView>("teachers");
  const [historyQuery, setHistoryQuery] = useState("");
  const [historyPage, setHistoryPage] = useState(1);
  const [historyTotal, setHistoryTotal] = useState(0);
  const [historyKindFilter, setHistoryKindFilter] = useState<HistoryFilterKind>("formal");
  const [historyStatusFilter, setHistoryStatusFilter] = useState<HistoryFilterStatus>("active");
  const [historyFrom, setHistoryFrom] = useState("");
  const [historyTo, setHistoryTo] = useState("");
  const [historyDetail, setHistoryDetail] = useState<HistoryDetail | null>(null);
  const [historyDetailPage, setHistoryDetailPage] = useState(1);
  const [teachers, setTeachers] = useState<TeacherSummary[]>([]);
  const [teacherTotal, setTeacherTotal] = useState(0);
  const [teacherPage, setTeacherPage] = useState(1);
  const [teacherQuery, setTeacherQuery] = useState("");
  const [teacherSort, setTeacherSort] = useState<"courses" | "weeks" | "recent">("courses");
  const [teacherDetail, setTeacherDetail] = useState<TeacherDetailPayload | null>(null);
  const [teacherDetailPage, setTeacherDetailPage] = useState(1);
  const [teacherLoading, setTeacherLoading] = useState(false);
  const [managedUsers, setManagedUsers] = useState<ManagedUser[]>([]);
  const [usersLoading, setUsersLoading] = useState(false);
  const [newUserUsername, setNewUserUsername] = useState("");
  const [newUserDisplayName, setNewUserDisplayName] = useState("");
  const [newUserPassword, setNewUserPassword] = useState("");
  const [newUserRole, setNewUserRole] = useState<AuthUser["role"]>("viewer");
  const [resetPasswords, setResetPasswords] = useState<Record<string, string>>({});
  const [auditEntries, setAuditEntries] = useState<AuditEntry[]>([]);
  const [monthlyMonth, setMonthlyMonth] = useState(currentMonthLabel);
  const [weeklySlots, setWeeklySlots] = useState("900");
  const [monthlyQuery, setMonthlyQuery] = useState("");
  const [monthlyLowOnly, setMonthlyLowOnly] = useState(false);
  const [monthlyPendingOnly, setMonthlyPendingOnly] = useState(false);
  const [teachingServiceSnapshot, setTeachingServiceSnapshot] = useState<TeachingServiceSnapshot | null>(null);
  const [monthlyData, setMonthlyData] = useState<MonthlyInspectionData | null>(null);
  const [monthlyError, setMonthlyError] = useState("");
  const workerRef = useRef<Worker | null>(null);
  const authCookieRef = useRef<InspectionAuthCookie | null>(initialAuthCookie);
  const authRefreshPromiseRef = useRef<Promise<void> | null>(null);
  const authRefreshTimerRef = useRef<number | null>(null);

  useEffect(() => () => {
    workerRef.current?.terminate();
    if (authRefreshTimerRef.current !== null) window.clearTimeout(authRefreshTimerRef.current);
  }, []);

  function updateStatus(title: string, message: string, progress = 0, mode: ProcessingStatus["mode"] = "working") {
    setStatus({ visible: true, title, message, progress, mode });
  }

  function finishWorker() {
    setProcessing(false);
    workerRef.current?.terminate();
    workerRef.current = null;
  }

  async function loadAuthenticatedData() {
    await Promise.all([loadHistory(1), loadTeachers(1)]);
  }

  function scheduleAuthRefresh(cookie: InspectionAuthCookie) {
    if (authRefreshTimerRef.current !== null) window.clearTimeout(authRefreshTimerRef.current);
    const elapsed = Math.max(0, Date.now() - cookie.issuedAt);
    const delay = Math.max(1000, INSPECTION_AUTH_REVALIDATION_MS - elapsed);
    authRefreshTimerRef.current = window.setTimeout(() => {
      authRefreshTimerRef.current = null;
      void loadAuthState(true);
    }, delay);
  }

  function cacheAuthUser(user: AuthUser) {
    const cookie: InspectionAuthCookie = {
      user,
      issuedAt: Date.now(),
    };
    authCookieRef.current = cookie;
    scheduleAuthRefresh(cookie);
  }

  function clearCachedAuth() {
    authCookieRef.current = null;
    clearInspectionAuthCookie();
    if (authRefreshTimerRef.current !== null) {
      window.clearTimeout(authRefreshTimerRef.current);
      authRefreshTimerRef.current = null;
    }
  }

  function loadInitialAuthenticatedData() {
    // The visible-view effects load the current table. Load batches in the
    // background only when that view will not already request them.
    if (historyView !== "batches") void loadHistory(1);
  }

  async function loadAuthState(background = false) {
    if (authRefreshPromiseRef.current) return authRefreshPromiseRef.current;
    const refresh = (async () => {
      try {
        const response = await fetch("/api/inspection/auth/me", { credentials: "same-origin", cache: "no-store" });
        const body = await responseJson(response);
        if (response.ok && body.user) {
          const user = body.user as AuthUser;
          cacheAuthUser(user);
          setAuthUser(user);
          setAuthenticated(true);
          setAuthChecked(true);
          if (!authCookieRef.current || !background) loadInitialAuthenticatedData();
          return;
        }

        clearCachedAuth();
        setAuthUser(null);
        setAuthenticated(false);
        setAuthChecked(true);
      } catch (error) {
        if (!background) {
          updateStatus("登录状态读取失败", errorMessage(error), 100, "error");
          setAuthChecked(true);
        }
        // A cached page stays usable after a transient background failure.
      }
    })();
    authRefreshPromiseRef.current = refresh;
    try {
      await refresh;
    } finally {
      if (authRefreshPromiseRef.current === refresh) authRefreshPromiseRef.current = null;
    }
  }

  async function login(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!loginUsername.trim() || !loginPassword) return;
    setLoginBusy(true);
    try {
      const response = await fetch("/api/inspection/auth/login", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username: loginUsername, password: loginPassword }),
      });
      const body = await responseJson(response);
      if (!response.ok || !body.user) throw new Error(String(body.error || "登录失败，请检查用户名和密码。"));
      const user = body.user as AuthUser;
      cacheAuthUser(user);
      setAuthUser(user);
      setAuthenticated(true);
      setAuthChecked(true);
      setLoginPassword("");
      loadInitialAuthenticatedData();
      updateStatus("登录成功", "抽检数据正在后台加载。", 100, "done");
    } catch (error) {
      clearCachedAuth();
      updateStatus("登录失败", errorMessage(error), 100, "error");
    } finally {
      setLoginBusy(false);
    }
  }

  async function logout() {
    try {
      await fetch("/api/inspection/auth/logout", { method: "POST", credentials: "same-origin", headers: { "Content-Type": "application/json" }, body: "{}" });
    } finally {
      clearCachedAuth();
      setAuthUser(null);
      setAuthenticated(false);
      setAuthChecked(true);
      setHistory([]);
      setTeachers([]);
      setTeacherDetail(null);
      setHistoryDetail(null);
      setHistoryView("teachers");
    }
  }

  async function loadManagedUsers() {
    setUsersLoading(true);
    try {
      const response = await fetch("/api/inspection/users", { credentials: "same-origin", cache: "no-store" });
      const body = await responseJson(response);
      if (!response.ok) throw new Error(String(body.error || "用户列表读取失败。"));
      setManagedUsers(Array.isArray(body.users) ? body.users as ManagedUser[] : []);
    } catch (error) {
      updateStatus("用户列表读取失败", errorMessage(error), 100, "error");
    } finally {
      setUsersLoading(false);
    }
  }

  async function loadAuditEntries() {
    try {
      const response = await fetch("/api/inspection/audit", { credentials: "same-origin", cache: "no-store" });
      const body = await responseJson(response);
      if (!response.ok) throw new Error(String(body.error || "审计记录读取失败。"));
      setAuditEntries(Array.isArray(body.entries) ? body.entries as AuditEntry[] : []);
    } catch (error) {
      updateStatus("审计记录读取失败", errorMessage(error), 100, "error");
    }
  }

  async function createManagedUser(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    try {
      const response = await fetch("/api/inspection/users", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username: newUserUsername, displayName: newUserDisplayName, password: newUserPassword, role: newUserRole }),
      });
      const body = await responseJson(response);
      if (!response.ok) throw new Error(String(body.error || "用户创建失败。"));
      setNewUserUsername("");
      setNewUserDisplayName("");
      setNewUserPassword("");
      await loadManagedUsers();
      updateStatus("用户已创建", "新用户可以使用账号密码登录抽检系统。", 100, "done");
    } catch (error) {
      updateStatus("用户创建失败", errorMessage(error), 100, "error");
    }
  }

  async function updateManagedUser(id: string, patch: Record<string, unknown>) {
    try {
      const response = await fetch(`/api/inspection/users/${encodeURIComponent(id)}`, {
        method: "PATCH",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(patch),
      });
      const body = await responseJson(response);
      if (!response.ok) throw new Error(String(body.error || "用户更新失败。"));
      await loadManagedUsers();
    } catch (error) {
      updateStatus("用户更新失败", errorMessage(error), 100, "error");
    }
  }

  async function resetManagedPassword(id: string) {
    const passwordValue = resetPasswords[id] || "";
    try {
      const response = await fetch(`/api/inspection/users/${encodeURIComponent(id)}/reset-password`, {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ password: passwordValue }),
      });
      const body = await responseJson(response);
      if (!response.ok) throw new Error(String(body.error || "密码重置失败。"));
      setResetPasswords((current) => ({ ...current, [id]: "" }));
      updateStatus("密码已重置", "该用户的旧会话已失效。", 100, "done");
    } catch (error) {
      updateStatus("密码重置失败", errorMessage(error), 100, "error");
    }
  }

  useEffect(() => {
    const refreshIfStale = () => {
      const cookie = authCookieRef.current;
      if (cookie && isInspectionAuthCookieStale(cookie)) void loadAuthState(true);
    };
    const handleVisibilityChange = () => {
      if (document.visibilityState === "visible") refreshIfStale();
    };
    window.addEventListener("focus", refreshIfStale);
    document.addEventListener("visibilitychange", handleVisibilityChange);

    if (initialAuthCookie) {
      // Render from the cookie first; this database check never blocks the page.
      void loadAuthState(true);
    } else {
      void loadAuthState(false);
    }

    return () => {
      window.removeEventListener("focus", refreshIfStale);
      document.removeEventListener("visibilitychange", handleVisibilityChange);
    };
  }, []);

  async function loadHistory(page = historyPage) {
    const params = new URLSearchParams({
      page: String(page),
      pageSize: "30",
      q: historyQuery,
      batchKind: historyKindFilter,
      status: historyStatusFilter,
      from: historyFrom,
      to: historyTo,
    });
    const response = await fetch(`/api/inspection/history?${params.toString()}`, { credentials: "same-origin", cache: "no-store" });
    const body = await responseJson(response);
    if (!response.ok) throw new Error(String(body.error || "抽检历史读取失败。"));
    const batches = Array.isArray(body.batches) ? body.batches as BatchSummary[] : [];
    setHistory(batches);
    setHistoryPage(page);
    setHistoryTotal(Number(body.total || 0));
    setActiveBatch(batches.find((batch) => batch.status === "active") || null);
  }

  async function loadHistoryDetail(id: string, page = 1) {
    try {
      const params = new URLSearchParams({ page: String(page), pageSize: "100" });
      const response = await fetch(`/api/inspection/history/${encodeURIComponent(id)}?${params.toString()}`, { credentials: "same-origin", cache: "no-store" });
      const body = await responseJson(response);
      if (!response.ok) throw new Error(String(body.error || "抽检批次读取失败。"));
      setHistoryDetail(body.batch as HistoryDetail);
      setHistoryDetailPage(page);
    } catch (error) {
      updateStatus("历史读取失败", errorMessage(error), 100, "error");
    }
  }

  async function loadTeachers(page = teacherPage) {
    setTeacherLoading(true);
    try {
      const params = new URLSearchParams({
        page: String(page),
        pageSize: "30",
        q: teacherQuery,
        batchKind: historyKindFilter,
        status: historyStatusFilter,
        from: historyFrom,
        to: historyTo,
        sort: teacherSort,
      });
      const response = await fetch(`/api/inspection/teachers?${params.toString()}`, { credentials: "same-origin", cache: "no-store" });
      const body = await responseJson(response);
      if (!response.ok) throw new Error(String(body.error || "教师抽检统计读取失败。"));
      setTeachers(Array.isArray(body.teachers) ? body.teachers as TeacherSummary[] : []);
      setTeacherTotal(Number(body.total || 0));
      setTeacherPage(page);
    } catch (error) {
      updateStatus("教师统计读取失败", errorMessage(error), 100, "error");
    } finally {
      setTeacherLoading(false);
    }
  }

  async function loadTeacherDetail(teacherKey: string, page = 1, overrides: Partial<Pick<HistoryFiltersForRequest, "batchKind" | "status" | "from" | "to">> = {}) {
    try {
      const params = new URLSearchParams({
        page: String(page),
        pageSize: "50",
        batchKind: overrides.batchKind || historyKindFilter,
        status: overrides.status || historyStatusFilter,
        from: overrides.from || historyFrom,
        to: overrides.to || historyTo,
      });
      const response = await fetch(`/api/inspection/teachers/${encodeURIComponent(teacherKey)}?${params.toString()}`, { credentials: "same-origin", cache: "no-store" });
      const body = await responseJson(response);
      if (!response.ok) throw new Error(String(body.error || "教师抽检明细读取失败。"));
      setTeacherDetail(body as unknown as TeacherDetailPayload);
      setTeacherDetailPage(page);
    } catch (error) {
      updateStatus("教师明细读取失败", errorMessage(error), 100, "error");
    }
  }

  async function loadMonthlyPanelData(month = monthlyMonth) {
    try {
      const [scoreResponse, monthlyResponse] = await Promise.all([
        fetch("/api/inspection/teaching-service", { credentials: "same-origin", cache: "no-store" }),
        fetch(`/api/inspection/monthly?month=${encodeURIComponent(month)}`, { credentials: "same-origin", cache: "no-store" }),
      ]);
      const scoreBody = await responseJson(scoreResponse);
      const monthlyBody = await responseJson(monthlyResponse);
      if (!scoreResponse.ok) throw new Error(String(scoreBody.error || "教学服务评分读取失败。"));
      if (!monthlyResponse.ok) throw new Error(String(monthlyBody.error || "月度抽检记录读取失败。"));
      setTeachingServiceSnapshot(scoreBody as unknown as TeachingServiceSnapshot);
      setMonthlyData(monthlyBody as unknown as MonthlyInspectionData);
      setMonthlyError("");
    } catch (error) {
      setMonthlyError(errorMessage(error));
    }
  }

  useEffect(() => {
    if (authenticated && historyView === "monthly") void loadMonthlyPanelData(monthlyMonth);
  }, [authenticated, historyView, monthlyMonth]);

  useEffect(() => {
    if (!authenticated || historyView !== "teachers") return;
    const timer = window.setTimeout(() => void loadTeachers(1), 220);
    return () => window.clearTimeout(timer);
  }, [authenticated, historyView, teacherQuery, teacherSort, historyKindFilter, historyStatusFilter, historyFrom, historyTo]);

  useEffect(() => {
    if (!authenticated || historyView !== "batches") return;
    const timer = window.setTimeout(() => void loadHistory(1), 220);
    return () => window.clearTimeout(timer);
  }, [authenticated, historyView, historyQuery, historyKindFilter, historyStatusFilter, historyFrom, historyTo]);

  useEffect(() => {
    if (authenticated && authUser?.role === "admin" && historyView === "users") {
      void Promise.all([loadManagedUsers(), loadAuditEntries()]);
    }
  }, [authenticated, authUser?.role, historyView]);

  async function handleUnlockHistory() {
    try {
      updateStatus("正在读取抽检历史", "只读取历史记录，不会重新抽取课程。", 30);
      if (authenticated) await loadAuthenticatedData();
      else throw new Error("请先登录抽检系统。");
      updateStatus("抽检历史已加载", "默认按教师汇总，可按教师、邮箱、项目组或业务周检索，再打开课程明细。", 100, "done");
    } catch (error) {
      updateStatus("历史读取失败", errorMessage(error), 100, "error");
    }
  }

  async function savePayload(payload: InspectionHistoryPayload, action: SaveAction, batchId = "") {
    const endpoint = action === "replace" ? `/api/inspection/batches/${encodeURIComponent(batchId)}/replace` : "/api/inspection/batches";
    const response = await fetch(endpoint, {
      method: "POST",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const body = await responseJson(response);
    if (response.status === 409 && body.batch) {
      return { conflict: true, reused: false, batch: body.batch as BatchSummary };
    }
    if (!response.ok) {
      const error = new Error(String(body.error || "抽检历史保存失败。")) as Error & { conflict?: BatchSummary };
      if (response.status === 409 && body.batch) error.conflict = body.batch as BatchSummary;
      throw error;
    }
    return body as unknown as SaveResponse;
  }

  async function runInspection(attempt: number, action: SaveAction, replaceBatchId = "", saveHistory = authenticated) {
    if (!feedbackFile) return;
    const worker = createProcessingWorker();
    workerRef.current?.terminate();
    workerRef.current = worker;
    setProcessing(true);
    updateStatus("正在生成抽检名单", `读取课程反馈和在职明细，第 ${attempt} 次抽取。`, 2);

    worker.onmessage = async ({ data }: MessageEvent<WorkerResponse>) => {
      if (data.type === "progress") {
        updateStatus(data.title, data.message, data.progress);
        return;
      }
      if (data.type === "error") {
        updateStatus("处理失败", data.message, 100, "error");
        finishWorker();
        return;
      }
      if (data.type !== "inspectionComplete") return;
      const complete = data as InspectionWorkerComplete;
      finishWorker();
      setSummary(complete.summary);
      try {
        let saved: SaveResponse | null = null;
        if (saveHistory) saved = await savePayload(complete.historyPayload, action, replaceBatchId);
        if (saved?.conflict && saved.batch) {
          setPendingReplacement(saved.batch);
          updateStatus("本周已有抽检批次", `已有第 ${saved.batch.attempt} 次抽检记录。确认替换后才会生成新名单。`, 100, "error");
          return;
        }
        if (saved?.reused && saved.batch.attempt !== attempt && action === "create") {
          updateStatus("正在还原现有批次", `本周已有第 ${saved.batch.attempt} 次抽检记录，重新生成同一份名单。`, 35);
          await runInspection(saved.batch.attempt, "reuse", "", false);
          return;
        }
        if (saved?.batch) {
          setActiveBatch(saved.batch);
          setPendingReplacement(null);
          void Promise.all([loadHistory(1), loadTeachers(1)]).catch(() => undefined);
          updateStatus(
            saved.reused ? "已还原本周抽检名单" : "抽检名单已保存并下载",
            `抽检 ${complete.summary.selectedRows.toLocaleString()} 条，覆盖 ${complete.summary.selectedTeachers.toLocaleString()} 位教师；未生成报告风险 ${complete.summary.unsubmittedRows.toLocaleString()} 条。`,
            100,
            "done",
          );
        } else {
          updateStatus("结果已下载，历史未保存", "当前页面未完成历史解锁或数据库尚未配置，已保留本地结果。", 100, "error");
        }
        downloadResult(complete.chunks, complete.filename);
      } catch (error) {
        const message = errorMessage(error);
        updateStatus("结果已下载，历史保存失败", message, 100, "error");
        downloadResult(complete.chunks, complete.filename);
      }
    };
    worker.onerror = (event) => {
      updateStatus("处理失败", event.message || "浏览器工作线程发生错误。", 100, "error");
      finishWorker();
    };
    worker.postMessage({
      type: "process",
      mode: "inspection",
      feedbackFile,
      rosterFile,
      sampleCount: Number(sampleCount),
      attempt,
      includeExplanation,
      batchKind,
    });
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!feedbackFile) return;
    try {
      const count = Number(sampleCount);
      if (!Number.isInteger(count) || count < 1) throw new Error("抽检数必须是大于 0 的整数。");
      let canSaveHistory = authenticated;
      if (!authenticated) {
        throw new Error("请先登录抽检系统。");
      }
      await runInspection(1, "create", "", canSaveHistory);
    } catch (error) {
      updateStatus("操作失败", errorMessage(error), 100, "error");
    }
  }

  async function replaceCurrentBatch() {
    if (!pendingReplacement) return;
    await runInspection(pendingReplacement.attempt + 1, "replace", pendingReplacement.id, true);
  }

  const monthlyPlan = useMemo<MonthlyInspectionPlan | null>(() => {
    if (!teachingServiceSnapshot || !monthlyData) return null;
    return buildMonthlyInspectionPlan(teachingServiceSnapshot, monthlyData, Number(weeklySlots) * monthWeekCount(monthlyMonth));
  }, [teachingServiceSnapshot, monthlyData, weeklySlots, monthlyMonth]);

  const visibleMonthlyTeachers = useMemo(() => {
    if (!monthlyPlan) return [];
    const query = monthlyQuery.trim().toLocaleLowerCase();
    return monthlyPlan.teachers.filter((teacher) => {
      const matchesQuery = !query || `${teacher.teacherName} ${teacher.teacherEmail} ${teacher.researchGroup} ${teacher.trainingLeader} ${teacher.campus}`.toLocaleLowerCase().includes(query);
      const matchesLow = !monthlyLowOnly || (teacher.scoreStatus === "matched" && (teacher.score ?? 100) < 80);
      const matchesPending = !monthlyPendingOnly || teacher.remainingCount > 0;
      return matchesQuery && matchesLow && matchesPending;
    });
  }, [monthlyPlan, monthlyQuery, monthlyLowOnly, monthlyPendingOnly]);

  const canOperate = authUser?.role === "admin" || authUser?.role === "operator";
  const isAdmin = authUser?.role === "admin";

  if (!authChecked) {
    return <><section className="card inspection-login-card"><div className="inspection-login-copy"><span className="history-kicker">INSPECTION ACCESS</span><strong>正在检查登录状态</strong><p>请稍候，系统正在读取抽检权限。</p></div></section><StatusCard status={status} /></>;
  }

  if (!authUser) {
    return (
      <>
        <section className="card inspection-login-card">
          <div className="inspection-login-copy">
            <span className="history-kicker">INSPECTION ACCESS</span>
            <strong>登录课堂反馈抽检</strong>
            <p>登录后才能查看抽检历史、教师统计和课程明细。账号由管理员统一配置，打卡质检仍可独立使用。</p>
          </div>
          <form className="inspection-login-form" onSubmit={login}>
            <label><span>用户名</span><input className="text-input" type="text" autoComplete="username" value={loginUsername} onChange={(event) => setLoginUsername(event.target.value)} placeholder="输入用户名" /></label>
            <label><span>密码</span><input className="text-input" type="password" autoComplete="current-password" value={loginPassword} onChange={(event) => setLoginPassword(event.target.value)} placeholder="输入密码" /></label>
            <button type="submit" disabled={loginBusy}><span>{loginBusy ? "正在登录…" : "登录"}</span><small>仅用于课堂反馈抽检数据</small></button>
          </form>
        </section>
        <StatusCard status={status} />
      </>
    );
  }

  return (
    <>
      {canOperate ? <section className="card inspection-card">
        <form onSubmit={handleSubmit}>
          <div className="grid">
            <UploadCard
              id="inspection-feedback-file"
              name="inspection_feedback_file"
              step="01"
              title="周度课程反馈明细"
              description="系统导出的本周全部课程记录"
              file={feedbackFile}
              onChange={setFeedbackFile}
            />
            <UploadCard
              id="inspection-roster-file"
              name="inspection_roster_file"
              step="02"
              title="最新在职明细"
              description="不上传时使用项目内置的 2026-09-15 快照"
              file={rosterFile}
              required={false}
              onChange={setRosterFile}
            />
          </div>

          <div className="options inspection-options">
            <label>
              <span className="option-title">抽检课程条数</span>
              <input
                className="number-input"
                type="number"
                min="1"
                step="1"
                value={sampleCount}
                onChange={(event) => setSampleCount(event.target.value)}
              />
              <small>先覆盖不同教师，再优先抽取“是否生成报告=否”的课程。</small>
            </label>
          </div>

          <label className="switch-row inspection-explanation">
            <span>
              <strong>同时导出处理说明</strong>
              <small>默认只导出抽检名单和风险表，需要复核过滤口径时再打开。</small>
            </span>
            <input type="checkbox" checked={includeExplanation} onChange={(event) => setIncludeExplanation(event.target.checked)} />
          </label>

          <label className="switch-row inspection-explanation">
            <span>
              <strong>本次标记为试运行</strong>
              <small>试运行批次会进入历史查询，但不占用正式周次，也不参与正式频次统计。</small>
            </span>
            <input type="checkbox" checked={batchKind === "trial"} onChange={(event) => setBatchKind(event.target.checked ? "trial" : "formal")} />
          </label>

          <div className="rules">
            <span>抽检规则</span>
            <b>在职邮箱过滤</b>
            <b>未生成报告优先</b>
            <b>尽量覆盖教师</b>
            <b>同周幂等</b>
          </div>

          <button type="submit" disabled={processing}>
            <span>生成课堂反馈抽检名单</span>
            <small>输出抽检名单、未生成报告风险表，并保存审计记录</small>
          </button>
        </form>
      </section> : null}

      {canOperate && pendingReplacement ? (
        <section className="inspection-alert">
          <div>
            <strong>本周已有第 {pendingReplacement.attempt} 次抽检记录</strong>
            <p>如果当前文件是修正版，可以废弃现有批次并生成第 {pendingReplacement.attempt + 1} 次记录。旧批次会保留为历史。</p>
          </div>
          <button type="button" onClick={replaceCurrentBatch} disabled={processing}>废弃并重新抽取</button>
        </section>
      ) : null}

      <StatusCard status={status} />

      {summary ? (
        <section className="inspection-summary">
          <article><span>候选课程</span><strong>{summary.eligibleRows.toLocaleString()}</strong><small>已通过在职邮箱过滤</small></article>
          <article><span>抽检课程</span><strong>{summary.selectedRows.toLocaleString()}</strong><small>不超过负责人填写的上限</small></article>
          <article><span>覆盖教师</span><strong>{summary.selectedTeachers.toLocaleString()}</strong><small>优先保证教师覆盖</small></article>
          <article><span>未生成报告</span><strong>{summary.unsubmittedRows.toLocaleString()}</strong><small>风险表会完整列出</small></article>
        </section>
      ) : null}

      {activeBatch ? (
        <section className="inspection-history-card">
          <div>
            <span className="history-kicker">CURRENT BATCH</span>
            <strong>{activeBatch.businessWeekStart}~{activeBatch.businessWeekEnd} · 第 {activeBatch.attempt} 次</strong>
            <p>已保存 {activeBatch.selectedCount.toLocaleString()} 条抽检课程，覆盖 {activeBatch.teacherCount.toLocaleString()} 位教师。</p>
          </div>
          <span className="history-status">已记录</span>
        </section>
      ) : null}

      {authenticated ? (
        <section className="inspection-history-workspace">
          <div className="inspection-history-heading history-workspace-heading">
            <div>
              <span className="history-kicker">INSPECTION DATABASE</span>
              <strong>抽检数据中心</strong>
              <p>默认按教师汇总。课程明细只在选择教师或批次后加载，避免历史越积越长。</p>
            </div>
            <div className="history-user-tools">
              <span className="history-user-badge">{authUser.displayName} · {authUser.role === "admin" ? "管理员" : authUser.role === "operator" ? "操作员" : "只读"}</span>
              <button className="history-action" type="button" onClick={() => void logout()}>退出登录</button>
            </div>
          </div>
          <div className="history-view-tabs" role="tablist" aria-label="抽检数据视图">
            <button type="button" className={historyView === "teachers" ? "active" : ""} onClick={() => setHistoryView("teachers")}>教师统计</button>
            <button type="button" className={historyView === "batches" ? "active" : ""} onClick={() => setHistoryView("batches")}>批次历史</button>
            <button type="button" className={historyView === "monthly" ? "active" : ""} onClick={() => setHistoryView("monthly")}>月度计划</button>
            {isAdmin ? <button type="button" className={historyView === "users" ? "active" : ""} onClick={() => setHistoryView("users")}>用户管理</button> : null}
            <button type="button" className="history-refresh-button" onClick={() => void handleUnlockHistory()} disabled={processing}>刷新数据</button>
          </div>

          {historyView !== "users" ? <div className="history-filter-bar">
            <input
              className="history-search"
              type="search"
              value={historyView === "teachers" ? teacherQuery : historyView === "batches" ? historyQuery : monthlyQuery}
              onChange={(event) => historyView === "teachers" ? setTeacherQuery(event.target.value) : historyView === "batches" ? setHistoryQuery(event.target.value) : setMonthlyQuery(event.target.value)}
              placeholder={historyView === "teachers" ? "搜索教师、邮箱、教研组、项目组" : historyView === "batches" ? "搜索周次或来源文件" : "按教师、教研组或师训组长查询"}
            />
            <select className="history-filter-select" value={historyKindFilter} disabled={historyView === "monthly"} onChange={(event) => setHistoryKindFilter(event.target.value as HistoryFilterKind)}>
              <option value="formal">正式批次</option>
              <option value="trial">试运行</option>
              <option value="all">全部类型</option>
            </select>
            <select className="history-filter-select" value={historyStatusFilter} disabled={historyView === "monthly"} onChange={(event) => setHistoryStatusFilter(event.target.value as HistoryFilterStatus)}>
              <option value="active">当前生效</option>
              <option value="voided">已废弃</option>
              <option value="all">全部状态</option>
            </select>
            <label className="history-date-filter">从<input type="date" disabled={historyView === "monthly"} value={historyFrom} onChange={(event) => setHistoryFrom(event.target.value)} /></label>
            <label className="history-date-filter">到<input type="date" disabled={historyView === "monthly"} value={historyTo} onChange={(event) => setHistoryTo(event.target.value)} /></label>
            {historyView === "teachers" ? (
              <select className="history-filter-select" value={teacherSort} onChange={(event) => setTeacherSort(event.target.value as "courses" | "weeks" | "recent")}>
                <option value="courses">按课程数</option>
                <option value="weeks">按业务周数</option>
                <option value="recent">按最近抽检</option>
              </select>
            ) : null}
          </div> : null}

          {historyView === "teachers" ? (
            <>
              <div className="inspection-data-summary">
                <article><span>教师数</span><strong>{teacherTotal.toLocaleString()}</strong><small>符合当前筛选条件</small></article>
                <article><span>当前页</span><strong>{teachers.length.toLocaleString()}</strong><small>每页 30 位教师</small></article>
                <article><span>统计口径</span><strong>{historyKindFilter === "formal" && historyStatusFilter === "active" ? "正式·生效" : "已筛选"}</strong><small>已废弃记录默认不混入</small></article>
              </div>
              <div className="history-table-wrap">
                <table className="history-table teacher-summary-table">
                  <thead><tr><th>教师</th><th>抽检课程</th><th>抽检批次</th><th>业务周</th><th>正式/试运行</th><th>未生成报告</th><th>最近抽检</th><th>项目组</th><th>操作</th></tr></thead>
                  <tbody>{teachers.map((teacher) => (
                    <tr key={teacher.teacherKey}>
                      <td><strong>{teacher.teacherName}</strong><small className="table-secondary">{teacher.teacherEmail || "无邮箱"}</small></td>
                      <td>{teacher.courseCount.toLocaleString()}</td>
                      <td>{teacher.batchCount.toLocaleString()}</td>
                      <td>{teacher.weekCount.toLocaleString()}</td>
                      <td>{teacher.formalCourseCount.toLocaleString()} / {teacher.trialCourseCount.toLocaleString()}</td>
                      <td>{teacher.unsubmittedCount.toLocaleString()}</td>
                      <td>{teacher.lastWeek || "—"}</td>
                      <td title={teacher.projectGroups}>{teacher.projectGroups || "—"}</td>
                      <td><button className="history-action" type="button" onClick={() => void loadTeacherDetail(teacher.teacherKey)}>查看明细</button></td>
                    </tr>
                  ))}</tbody>
                </table>
                {teacherLoading ? <p className="history-empty">正在读取教师统计……</p> : null}
                {!teacherLoading && !teachers.length ? <p className="history-empty">没有符合条件的教师。</p> : null}
              </div>
              <div className="history-pagination">
                <span>共 {teacherTotal.toLocaleString()} 位教师</span>
                <button type="button" disabled={teacherPage <= 1 || teacherLoading} onClick={() => void loadTeachers(teacherPage - 1)}>上一页</button>
                <span>第 {teacherPage} / {Math.max(1, Math.ceil(teacherTotal / 30))} 页</span>
                <button type="button" disabled={teacherPage >= Math.ceil(teacherTotal / 30) || teacherLoading} onClick={() => void loadTeachers(teacherPage + 1)}>下一页</button>
              </div>
            </>
          ) : null}

          {historyView === "batches" ? (
            <>
              <div className="history-table-wrap">
                <table className="history-table">
                  <thead><tr><th>业务周</th><th>尝试</th><th>来源</th><th>抽检</th><th>风险</th><th>状态</th><th>操作</th></tr></thead>
                  <tbody>{history.map((batch) => (
                    <tr key={batch.id}>
                      <td>{batch.businessWeekStart}~{batch.businessWeekEnd}</td>
                      <td>第 {batch.attempt} 次</td>
                      <td title={`${batch.sourceName}；${batch.rosterName}`}>{batch.sourceName}</td>
                      <td>{batch.selectedCount.toLocaleString()} / {batch.teacherCount.toLocaleString()}人</td>
                      <td>{batch.unsubmittedCount.toLocaleString()}</td>
                      <td><span className={`history-badge ${batch.batchKind}`}>{batch.batchKind === "trial" ? "试运行" : "正式"}</span> <span className={`history-badge ${batch.status}`}>{batch.status === "active" ? "生效" : "已废弃"}</span></td>
                      <td><button className="history-action" type="button" onClick={() => void loadHistoryDetail(batch.id)}>查看课程</button></td>
                    </tr>
                  ))}</tbody>
                </table>
                {!history.length ? <p className="history-empty">暂时没有符合条件的抽检批次。</p> : null}
              </div>
              <div className="history-pagination">
                <span>共 {historyTotal.toLocaleString()} 个批次</span>
                <button type="button" disabled={historyPage <= 1} onClick={() => void loadHistory(historyPage - 1)}>上一页</button>
                <span>第 {historyPage} / {Math.max(1, Math.ceil(historyTotal / 30))} 页</span>
                <button type="button" disabled={historyPage >= Math.ceil(historyTotal / 30)} onClick={() => void loadHistory(historyPage + 1)}>下一页</button>
              </div>
            </>
          ) : null}

          {historyView === "users" && isAdmin ? (
            <div className="admin-users-panel">
              <form className="admin-user-create-form" onSubmit={createManagedUser}>
                <div><label>用户名<input className="text-input" value={newUserUsername} onChange={(event) => setNewUserUsername(event.target.value)} placeholder="例如 zhangsan" /></label></div>
                <div><label>显示名称<input className="text-input" value={newUserDisplayName} onChange={(event) => setNewUserDisplayName(event.target.value)} placeholder="例如 张三" /></label></div>
                <div><label>初始密码<input className="text-input" type="password" value={newUserPassword} onChange={(event) => setNewUserPassword(event.target.value)} placeholder="不能为空，最多 200 位" /></label></div>
                <div><label>角色<select className="history-filter-select" value={newUserRole} onChange={(event) => setNewUserRole(event.target.value as AuthUser["role"])}><option value="viewer">只读</option><option value="operator">操作员</option><option value="admin">管理员</option></select></label></div>
                <button type="submit">新增用户</button>
              </form>
              <div className="history-table-wrap">
                <table className="history-table admin-users-table">
                  <thead><tr><th>用户</th><th>角色</th><th>状态</th><th>最后登录</th><th>修改角色</th><th>停用/恢复</th><th>重置密码</th></tr></thead>
                  <tbody>{managedUsers.map((user) => (
                    <tr key={user.id}>
                      <td><strong>{user.displayName}</strong><small className="table-secondary">{user.username}</small></td>
                      <td><select className="history-filter-select" value={user.role} disabled={user.id === authUser.id} onChange={(event) => void updateManagedUser(user.id, { role: event.target.value })}><option value="viewer">只读</option><option value="operator">操作员</option><option value="admin">管理员</option></select></td>
                      <td><span className={`history-badge ${user.isActive ? "active" : "voided"}`}>{user.isActive ? "启用" : "已停用"}</span></td>
                      <td>{user.lastLoginAt ? new Date(user.lastLoginAt).toLocaleString("zh-CN") : "未登录"}</td>
                      <td><button className="history-action" type="button" disabled={user.id === authUser.id} onClick={() => void updateManagedUser(user.id, { role: user.role === "admin" ? "viewer" : "admin" })}>{user.role === "admin" ? "降为只读" : "升为管理员"}</button></td>
                      <td><button className="history-action" type="button" disabled={user.id === authUser.id} onClick={() => void updateManagedUser(user.id, { isActive: !user.isActive })}>{user.isActive ? "停用" : "恢复"}</button></td>
                      <td><div className="admin-reset-control"><input className="text-input" type="password" value={resetPasswords[user.id] || ""} onChange={(event) => setResetPasswords((current) => ({ ...current, [user.id]: event.target.value }))} placeholder="新密码" /><button className="history-action" type="button" onClick={() => void resetManagedPassword(user.id)}>重置</button></div></td>
                    </tr>
                  ))}</tbody>
                </table>
                {usersLoading ? <p className="history-empty">正在读取用户……</p> : null}
                {!usersLoading && !managedUsers.length ? <p className="history-empty">暂时没有用户。</p> : null}
              </div>
              <div className="admin-audit-heading"><strong>最近操作审计</strong><button className="history-action" type="button" onClick={() => void loadAuditEntries()}>刷新审计</button></div>
              <div className="history-table-wrap">
                <table className="history-table admin-audit-table">
                  <thead><tr><th>时间</th><th>操作人</th><th>对象</th><th>动作</th><th>附加信息</th></tr></thead>
                  <tbody>{auditEntries.map((entry) => <tr key={entry.id}><td>{new Date(entry.createdAt).toLocaleString("zh-CN")}</td><td>{entry.actorUsername}</td><td>{entry.targetUsername}</td><td>{entry.action}</td><td>{JSON.stringify(entry.metadata || {})}</td></tr>)}</tbody>
                </table>
              </div>
            </div>
          ) : null}
        </section>
      ) : null}

      {teacherDetail ? (
        <section className="inspection-history-detail teacher-detail-panel">
          <div className="inspection-history-heading">
            <div>
              <span className="history-kicker">TEACHER DETAIL</span>
              <strong>{teacherDetail.teacher.teacherName} · {teacherDetail.teacher.courseCount.toLocaleString()} 节抽检课程</strong>
              <p>{teacherDetail.teacher.teacherEmail || "无邮箱"} · {teacherDetail.teacher.batchCount} 个批次 · {teacherDetail.teacher.weekCount} 个业务周 · 未生成报告 {teacherDetail.teacher.unsubmittedCount} 节</p>
            </div>
            <button className="history-action" type="button" onClick={() => setTeacherDetail(null)}>关闭明细</button>
          </div>
          <div className="teacher-week-list">
            {teacherDetail.weeks.map((week) => <div className="teacher-week-card" key={week.batchId}><strong>{week.businessWeekStart}~{week.businessWeekEnd}</strong><span>第 {week.attempt} 次 · {week.courseCount} 节 · 未生成 {week.unsubmittedCount}</span><em className={`history-badge ${week.status}`}>{week.status === "active" ? "生效" : "已废弃"}</em></div>)}
          </div>
          <div className="history-detail-table-wrap">
            <table className="history-table history-detail-table">
              <thead><tr><th>业务周</th><th>批次</th><th>学员</th><th>课程ID</th><th>上课时间</th><th>是否生成报告</th><th>入选原因</th></tr></thead>
              <tbody>{teacherDetail.items.map((item) => <tr key={`${item.batchId}-${item.position}`}><td>{item.businessWeekStart}~{item.businessWeekEnd}</td><td>第 {item.attempt} 次</td><td>{item.studentName}</td><td>{item.courseId}</td><td>{item.lessonStart}~{item.lessonEnd.slice(11)}</td><td>{item.submittedValue}</td><td>{item.selectionReason}</td></tr>)}</tbody>
            </table>
            {!teacherDetail.items.length ? <p className="history-empty">没有符合条件的课程明细。</p> : null}
          </div>
          <div className="history-pagination">
            <span>共 {teacherDetail.total.toLocaleString()} 节课程</span>
            <button type="button" disabled={teacherDetailPage <= 1} onClick={() => void loadTeacherDetail(teacherDetail.teacher.teacherKey, teacherDetailPage - 1)}>上一页</button>
            <span>第 {teacherDetailPage} / {Math.max(1, Math.ceil(teacherDetail.total / 50))} 页</span>
            <button type="button" disabled={teacherDetailPage >= Math.ceil(teacherDetail.total / 50)} onClick={() => void loadTeacherDetail(teacherDetail.teacher.teacherKey, teacherDetailPage + 1)}>下一页</button>
          </div>
        </section>
      ) : null}

      {authenticated && historyView === "monthly" ? (
        <section className="monthly-inspection-panel">
          <div className="inspection-history-heading">
            <div>
              <span className="history-kicker">TEACHING SERVICE MONTHLY REVIEW</span>
              <strong>教学服务月度抽检</strong>
              <p>高分每月一次，中高分每两周一次，其余每周至少一次；空余次数优先给普检未发送，再给低分教师加抽。当前按每周 800～1,000 次估算，默认使用 900 次。</p>
            </div>
            <div className="monthly-controls">
              <label>月份<input className="monthly-input" type="month" value={monthlyMonth} onChange={(event) => setMonthlyMonth(event.target.value)} /></label>
              <label>每周可用抽检课程数<input className="monthly-input" type="number" min="0" step="1" value={weeklySlots} onChange={(event) => setWeeklySlots(event.target.value)} /></label>
            </div>
          </div>
          {monthlyError ? <p className="monthly-error">{monthlyError}</p> : null}
          {monthlyPlan ? (
            <>
              <div className="monthly-summary">
                <article><span>候选教师</span><strong>{monthlyPlan.candidateTeachers.toLocaleString()}</strong></article>
                <article><span>模拟分配次数</span><strong>{monthlyPlan.assignedSlots.toLocaleString()}</strong></article>
                <article><span>低分教师</span><strong>{monthlyPlan.lowScoreTeachers.toLocaleString()}</strong></article>
                <article><span>覆盖缺口</span><strong>{monthlyPlan.coverageShortfall.toLocaleString()}</strong></article>
                <article><span>频次缺口</span><strong>{monthlyPlan.frequencyShortfall.toLocaleString()}</strong></article>
                <article><span>季度无分数</span><strong>{monthlyPlan.missingScoreTeachers.toLocaleString()}</strong></article>
              </div>
              <div className="monthly-toolbar">
                <label><input type="checkbox" checked={monthlyLowOnly} onChange={(event) => setMonthlyLowOnly(event.target.checked)} />只看低分教师</label>
                <label><input type="checkbox" checked={monthlyPendingOnly} onChange={(event) => setMonthlyPendingOnly(event.target.checked)} />只看待补足</label>
              </div>
              <div className="history-table-wrap">
                <table className="history-table monthly-table">
                  <thead><tr><th>教师</th><th>教研组</th><th>师训组长</th><th>季度赋分</th><th>建议次数</th><th>模拟分配</th><th>已抽次数</th><th>普检未发送</th><th>剩余</th><th>最近抽检周</th><th>重点原因</th><th>操作</th></tr></thead>
                  <tbody>{visibleMonthlyTeachers.map((teacher) => (
                    <tr key={teacher.teacherEmail}>
                      <td title={teacher.teacherEmail}>{teacher.teacherName}</td>
                      <td>{teacher.researchGroup || "—"}</td>
                      <td>{teacher.trainingLeader || "—"}</td>
                      <td>{teacher.scoreStatus === "matched" && teacher.score != null ? teacher.score.toFixed(1) : "季度无分数"}</td>
                      <td>{teacher.recommendedCount}</td>
                      <td>{teacher.assignedCount}</td>
                      <td>{teacher.actualCount}</td>
                      <td>{teacher.unsubmittedCount}</td>
                      <td>{teacher.remainingCount}</td>
                      <td>{teacher.lastInspectionWeek || "—"}</td>
                      <td>{teacher.focusReason}</td>
                      <td><button className="history-action" type="button" onClick={() => void loadTeacherDetail(teacher.teacherEmail ? `email:${teacher.teacherEmail}` : `name:${teacher.teacherName.trim().toLocaleLowerCase()}`, 1, { batchKind: "formal", status: "active", from: `${monthlyMonth}-01`, to: monthEndLabel(monthlyMonth) })}>查看课程</button></td>
                    </tr>
                  ))}</tbody>
                </table>
                {!visibleMonthlyTeachers.length ? <p className="history-empty">没有符合条件的教师。</p> : null}
              </div>
            </>
          ) : <p className="history-empty">正在读取季度评分和正式抽检记录。</p>}
        </section>
      ) : null}

      {historyDetail ? (
        <section className="inspection-history-detail">
          <div className="inspection-history-heading">
            <div>
              <span className="history-kicker">BATCH DETAIL</span>
              <strong>{historyDetail.businessWeekStart}~{historyDetail.businessWeekEnd} · 第 {historyDetail.attempt} 次</strong>
              <p>{historyDetail.sourceName} · {historyDetail.total.toLocaleString()} 条入选课程</p>
            </div>
            <button className="history-action" type="button" onClick={() => setHistoryDetail(null)}>关闭明细</button>
          </div>
          <div className="history-detail-table-wrap">
            <table className="history-table history-detail-table">
              <thead><tr><th>序号</th><th>教师</th><th>教师邮箱</th><th>学员</th><th>课程ID</th><th>上课时间</th><th>是否生成报告</th><th>入选原因</th></tr></thead>
              <tbody>{historyDetail.items.map((item) => <tr key={`${historyDetail.id}-${item.position}`}><td>{item.position}</td><td>{item.teacherName}</td><td>{item.teacherEmail}</td><td>{item.studentName}</td><td>{item.courseId}</td><td>{item.lessonStart}~{item.lessonEnd.slice(11)}</td><td>{item.submittedValue}</td><td>{item.selectionReason}</td></tr>)}</tbody>
            </table>
          </div>
          <div className="history-pagination">
            <span>共 {historyDetail.total.toLocaleString()} 节课程</span>
            <button type="button" disabled={historyDetailPage <= 1} onClick={() => void loadHistoryDetail(historyDetail.id, historyDetailPage - 1)}>上一页</button>
            <span>第 {historyDetailPage} / {Math.max(1, Math.ceil(historyDetail.total / 100))} 页</span>
            <button type="button" disabled={historyDetailPage >= Math.ceil(historyDetail.total / 100)} onClick={() => void loadHistoryDetail(historyDetail.id, historyDetailPage + 1)}>下一页</button>
          </div>
        </section>
      ) : null}

    </>
  );
}
