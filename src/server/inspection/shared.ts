/// <reference types="node" />

import { createHash, createHmac, randomBytes, scrypt, timingSafeEqual } from "node:crypto";
import { promisify } from "node:util";
import { neon } from "@neondatabase/serverless";

declare const process: { env: Record<string, string | undefined> };

export interface ApiRequest {
  method?: string;
  body?: unknown;
  headers?: Record<string, string | string[] | undefined>;
  query?: Record<string, string | string[] | undefined>;
}

export interface ApiResponse {
  status(code: number): ApiResponse;
  setHeader(name: string, value: string | string[]): void;
  json(body: unknown): void;
  send(body: string): void;
}

const SESSION_MAX_AGE_SECONDS = 8 * 60 * 60;
const RULE_VERSION = "inspection-v1";
const AUTH_COOKIE = "inspection_auth";
const PASSWORD_MIN_LENGTH = 10;
const PASSWORD_HASH_PREFIX = "scrypt$v1";
const scryptAsync = promisify(scrypt) as unknown as (password: string | Buffer, salt: string | Buffer, keylen: number, options?: { N?: number; r?: number; p?: number; maxmem?: number }) => Promise<Buffer>;

export type UserRole = "admin" | "operator" | "viewer";
const ROLE_RANK: Record<UserRole, number> = { viewer: 1, operator: 2, admin: 3 };

export interface AuthUser {
  id: string;
  username: string;
  displayName: string;
  role: UserRole;
  isActive: boolean;
  lastLoginAt: string | null;
}

let schemaPromise: Promise<void> | null = null;

function env(name: string) {
  return process.env[name] || "";
}

function sessionSecret() {
  return env("INSPECTION_SESSION_SECRET") || env("INSPECTION_PASSWORD");
}

function authMode() {
  const mode = env("INSPECTION_AUTH_MODE").toLowerCase();
  return mode === "login" || mode === "legacy" || mode === "dual" ? mode : "dual";
}

function signature(value: string) {
  return createHmac("sha256", sessionSecret()).update(value).digest("hex");
}

function cookieHeader(value: string, maxAge: number) {
  return `inspection_session=${value}; Max-Age=${maxAge}; Path=/; HttpOnly; SameSite=Lax${env("VERCEL_ENV") === "production" ? "; Secure" : ""}`;
}

export function authCookieHeader(value: string, maxAge: number) {
  return `${AUTH_COOKIE}=${value}; Max-Age=${maxAge}; Path=/; HttpOnly; SameSite=Lax${env("VERCEL_ENV") === "production" ? "; Secure" : ""}`;
}

export function issueSession() {
  const issuedAt = Math.floor(Date.now() / 1000);
  const nonce = `${issuedAt}.${randomBytes(16).toString("hex")}`;
  return `${nonce}.${signature(nonce)}`;
}

function requestCookie(request: ApiRequest) {
  const header = request.headers?.cookie;
  const cookie = Array.isArray(header) ? header.join(";") : header || "";
  return cookie.split(";").map((item) => item.trim()).find((item) => item.startsWith("inspection_session="))?.slice("inspection_session=".length) || "";
}

function requestCookieValue(request: ApiRequest, name: string) {
  const header = request.headers?.cookie;
  const cookie = Array.isArray(header) ? header.join(";") : header || "";
  return cookie.split(";").map((item) => item.trim()).find((item) => item.startsWith(`${name}=`))?.slice(name.length + 1) || "";
}

export function hasValidSession(request: ApiRequest) {
  const token = requestCookie(request);
  const parts = token.split(".");
  if (parts.length !== 3 || !sessionSecret()) return false;
  const nonce = `${parts[0]}.${parts[1]}`;
  const issuedAt = Number(parts[0]);
  if (!Number.isFinite(issuedAt) || Math.floor(Date.now() / 1000) - issuedAt > SESSION_MAX_AGE_SECONDS) return false;
  const actual = Buffer.from(parts[2]);
  const expected = Buffer.from(signature(nonce));
  return actual.length === expected.length && timingSafeEqual(actual, expected);
}

export async function requireSession(request: ApiRequest, response: ApiResponse, minimumRole: UserRole = "viewer") {
  const user = await currentAuthUser(request);
  if (user) {
    if (!canRole(user.role, minimumRole)) {
      response.status(403).json({ error: "当前账号没有执行此操作的权限。" });
      return false;
    }
    return true;
  }
  if (authMode() !== "login" && env("INSPECTION_PASSWORD") && hasValidSession(request)) return true;
  response.status(401).json({ error: "请先登录抽检系统。" });
  return false;
}

function normalizeUsername(value: unknown) {
  return safeText(value, 120).toLocaleLowerCase();
}

function validateUsername(value: unknown) {
  const username = normalizeUsername(value);
  if (!/^[a-z0-9][a-z0-9._+@-]{1,118}$/u.test(username)) throw new Error("用户名应为 2～120 位字母、数字或常用符号。");
  return username;
}

export function validatePassword(value: unknown) {
  const password = String(value ?? "");
  const length = Array.from(password).length;
  if (length < PASSWORD_MIN_LENGTH || length > 200) throw new Error(`密码长度应为 ${PASSWORD_MIN_LENGTH}～200 位。`);
  return password;
}

export async function hashPassword(value: unknown) {
  const password = validatePassword(value);
  const salt = randomBytes(16);
  const derived = await scryptAsync(password, salt, 64, { N: 32768, r: 8, p: 1, maxmem: 64 * 1024 * 1024 }) as Buffer;
  return `${PASSWORD_HASH_PREFIX}$32768$8$1$${salt.toString("base64url")}$${derived.toString("base64url")}`;
}

export async function verifyPassword(value: unknown, encoded: string) {
  const password = String(value ?? "");
  const parts = String(encoded || "").split("$");
  if (parts.length !== 7 || `${parts[0]}$${parts[1]}` !== PASSWORD_HASH_PREFIX) return false;
  const n = Number(parts[2]);
  const r = Number(parts[3]);
  const p = Number(parts[4]);
  if (!Number.isInteger(n) || !Number.isInteger(r) || !Number.isInteger(p)) return false;
  try {
    const salt = Buffer.from(parts[5], "base64url");
    const expected = Buffer.from(parts[6], "base64url");
    const actual = await scryptAsync(password, salt, expected.length, { N: n, r, p, maxmem: 64 * 1024 * 1024 }) as Buffer;
    return actual.length === expected.length && timingSafeEqual(actual, expected);
  } catch {
    return false;
  }
}

function tokenHash(value: string) {
  return createHash("sha256").update(value).digest("hex");
}

function mapUser(row: Record<string, any>): AuthUser {
  return {
    id: String(row.id),
    username: String(row.username),
    displayName: String(row.display_name || row.username),
    role: row.role === "admin" || row.role === "operator" ? row.role : "viewer",
    isActive: row.is_active !== false,
    lastLoginAt: row.last_login_at ? String(row.last_login_at) : null,
  };
}

function canRole(role: UserRole, minimum: UserRole) {
  return ROLE_RANK[role] >= ROLE_RANK[minimum];
}

function headerValue(request: ApiRequest, name: string) {
  const value = request.headers?.[name] ?? request.headers?.[name.toLowerCase()];
  return Array.isArray(value) ? value[0] || "" : value || "";
}

export function assertSameOrigin(request: ApiRequest) {
  const origin = headerValue(request, "origin");
  if (!origin) return;
  const host = headerValue(request, "x-forwarded-host") || headerValue(request, "host");
  if (!host) return;
  try {
    if (new URL(origin).host !== host.split(",")[0].trim()) throw new Error("请求来源不受信任。");
  } catch (error) {
    if (error instanceof Error && error.message === "请求来源不受信任。") throw error;
    throw new Error("请求来源不受信任。");
  }
}

export function parseBody(request: ApiRequest) {
  if (request.body && typeof request.body === "object") return request.body as Record<string, unknown>;
  if (typeof request.body === "string" && request.body.trim()) return JSON.parse(request.body) as Record<string, unknown>;
  return {};
}

export function safeText(value: unknown, max = 500) {
  return String(value ?? "").trim().slice(0, max);
}

export function safeInt(value: unknown, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.floor(number) : fallback;
}

export function safeSha(value: unknown) {
  const result = safeText(value, 64).toLowerCase();
  if (!/^[a-f0-9]{64}$/u.test(result)) throw new Error("输入文件摘要格式不正确。");
  return result;
}

function databaseDate(value: unknown) {
  if (!value) return null;
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/u.test(value)) return value;
  const date = value instanceof Date ? value : new Date(String(value));
  if (!Number.isFinite(date.valueOf())) return safeText(value, 20);
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

export function assertHistoryPayload(value: unknown) {
  const body = (value || {}) as { batch?: Record<string, unknown>; items?: unknown };
  const batch = body.batch || {};
  const rawItems = Array.isArray(body.items) ? body.items : [];
  const items = rawItems.slice(0, 10000).map((item, index) => {
    const row = (item || {}) as Record<string, unknown>;
    return {
      position: safeInt(row.position, index + 1),
      teacherName: safeText(row.teacherName, 80),
      teacherEmail: safeText(row.teacherEmail, 160).toLowerCase(),
      studentName: safeText(row.studentName, 80),
      studentId: safeText(row.studentId, 120),
      courseId: safeText(row.courseId, 120),
      lessonStart: safeText(row.lessonStart, 40),
      lessonEnd: safeText(row.lessonEnd, 40),
      submittedValue: safeText(row.submittedValue, 20),
      productGroup: safeText(row.productGroup, 120),
      campus: safeText(row.campus, 120),
      projectGroup: safeText(row.projectGroup, 120),
      selectionReason: safeText(row.selectionReason, 200),
    };
  });
  const normalized = {
    businessWeekStart: safeText(batch.businessWeekStart, 10),
    businessWeekEnd: safeText(batch.businessWeekEnd, 10),
    sourceName: safeText(batch.sourceName, 240),
    sourceSha256: safeSha(batch.sourceSha256),
    rosterName: safeText(batch.rosterName, 240),
    rosterSha256: safeSha(batch.rosterSha256),
    rosterSnapshotDate: safeText(batch.rosterSnapshotDate, 10),
    sampleLimit: safeInt(batch.sampleLimit),
    eligibleCount: safeInt(batch.eligibleCount),
    selectedCount: safeInt(batch.selectedCount),
    teacherCount: safeInt(batch.teacherCount),
    unsubmittedCount: safeInt(batch.unsubmittedCount),
    ruleVersion: safeText(batch.ruleVersion, 80) || RULE_VERSION,
    attempt: Math.max(1, safeInt(batch.attempt, 1)),
    batchKind: batch.batchKind === "trial" ? "trial" : "formal",
  };
  if (!/^\d{4}-\d{2}-\d{2}$/u.test(normalized.businessWeekStart) || !/^\d{4}-\d{2}-\d{2}$/u.test(normalized.businessWeekEnd)) throw new Error("业务周日期格式不正确。");
  if (!normalized.sourceName || !normalized.rosterName || !items.length) throw new Error("抽检历史记录缺少来源文件或抽检课程。");
  return { batch: normalized, items };
}

export async function database() {
  const url = env("DATABASE_URL");
  if (!url) throw new Error("历史数据库尚未配置 DATABASE_URL。");
  const sql = neon(url);
  if (!schemaPromise) {
    schemaPromise = (async () => {
      await sql`CREATE TABLE IF NOT EXISTS inspection_batches (
        id text PRIMARY KEY,
        business_week_start date NOT NULL,
        business_week_end date NOT NULL,
        source_name text NOT NULL,
        source_sha256 text NOT NULL,
        roster_name text NOT NULL,
        roster_sha256 text NOT NULL,
        roster_snapshot_date date,
        sample_limit integer NOT NULL,
        eligible_count integer NOT NULL,
        selected_count integer NOT NULL,
        teacher_count integer NOT NULL,
        unsubmitted_count integer NOT NULL,
        rule_version text NOT NULL,
        attempt integer NOT NULL,
        batch_kind text NOT NULL DEFAULT 'formal',
        status text NOT NULL DEFAULT 'active',
        created_at timestamptz NOT NULL DEFAULT now(),
        voided_at timestamptz
      )`;
      await sql`ALTER TABLE inspection_batches ADD COLUMN IF NOT EXISTS batch_kind text NOT NULL DEFAULT 'formal'`;
      await sql`DROP INDEX IF EXISTS inspection_active_week_idx`;
      await sql`CREATE UNIQUE INDEX IF NOT EXISTS inspection_active_week_kind_idx ON inspection_batches (business_week_start, batch_kind) WHERE status = 'active'`;
      await sql`CREATE INDEX IF NOT EXISTS inspection_batches_filter_idx ON inspection_batches (status, batch_kind, business_week_start DESC, created_at DESC)`;
      await sql`CREATE TABLE IF NOT EXISTS inspection_items (
        batch_id text NOT NULL REFERENCES inspection_batches(id),
        position integer NOT NULL,
        teacher_name text NOT NULL,
        teacher_email text NOT NULL,
        student_name text NOT NULL,
        student_id text NOT NULL,
        course_id text NOT NULL,
        lesson_start text NOT NULL,
        lesson_end text NOT NULL,
        submitted_value text NOT NULL,
        product_group text NOT NULL,
        campus text NOT NULL,
        project_group text NOT NULL,
        selection_reason text NOT NULL,
        PRIMARY KEY (batch_id, position)
      )`;
      await sql`CREATE INDEX IF NOT EXISTS inspection_items_teacher_idx ON inspection_items (LOWER(TRIM(teacher_email)), batch_id)`;
      await sql`CREATE TABLE IF NOT EXISTS inspection_users (
        id text PRIMARY KEY,
        username text NOT NULL UNIQUE,
        display_name text NOT NULL,
        password_hash text NOT NULL,
        role text NOT NULL DEFAULT 'viewer' CHECK (role IN ('admin', 'operator', 'viewer')),
        is_active boolean NOT NULL DEFAULT true,
        session_version integer NOT NULL DEFAULT 1,
        last_login_at timestamptz,
        created_at timestamptz NOT NULL DEFAULT now(),
        updated_at timestamptz NOT NULL DEFAULT now()
      )`;
      await sql`CREATE TABLE IF NOT EXISTS inspection_sessions (
        token_hash text PRIMARY KEY,
        user_id text NOT NULL REFERENCES inspection_users(id) ON DELETE CASCADE,
        created_at timestamptz NOT NULL DEFAULT now(),
        expires_at timestamptz NOT NULL,
        revoked_at timestamptz
      )`;
      await sql`CREATE INDEX IF NOT EXISTS inspection_sessions_user_idx ON inspection_sessions (user_id, expires_at DESC) WHERE revoked_at IS NULL`;
      await sql`CREATE TABLE IF NOT EXISTS inspection_login_attempts (
        username text PRIMARY KEY,
        window_started_at timestamptz NOT NULL DEFAULT now(),
        failed_count integer NOT NULL DEFAULT 0,
        locked_until timestamptz
      )`;
      await sql`CREATE TABLE IF NOT EXISTS inspection_audit_log (
        id text PRIMARY KEY,
        actor_user_id text REFERENCES inspection_users(id) ON DELETE SET NULL,
        target_user_id text REFERENCES inspection_users(id) ON DELETE SET NULL,
        action text NOT NULL,
        metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
        created_at timestamptz NOT NULL DEFAULT now()
      )`;
      await sql`CREATE INDEX IF NOT EXISTS inspection_audit_created_idx ON inspection_audit_log (created_at DESC)`;
    })();
  }
  await schemaPromise;
  return sql;
}

export function authModeInfo() {
  return { mode: authMode(), legacyAvailable: authMode() !== "login" && Boolean(env("INSPECTION_PASSWORD")) };
}

export function publicUser(user: AuthUser) {
  return {
    id: user.id,
    username: user.username,
    displayName: user.displayName,
    role: user.role,
    lastLoginAt: user.lastLoginAt,
  };
}

export async function currentAuthUser(request: ApiRequest): Promise<AuthUser | null> {
  const rawToken = requestCookieValue(request, AUTH_COOKIE);
  if (!rawToken) return null;
  const sql = await database();
  const rows = await sql`
    SELECT u.id, u.username, u.display_name, u.role, u.is_active, u.last_login_at
    FROM inspection_sessions s
    JOIN inspection_users u ON u.id = s.user_id
    WHERE s.token_hash = ${tokenHash(rawToken)}
      AND s.revoked_at IS NULL
      AND s.expires_at > now()
      AND u.is_active = true
    LIMIT 1
  `;
  return rows.length ? mapUser(rows[0]) : null;
}

export async function createAuthSession(userId: string) {
  const sql = await database();
  const rawToken = randomBytes(32).toString("base64url");
  await sql`INSERT INTO inspection_sessions (token_hash, user_id, expires_at) VALUES (${tokenHash(rawToken)}, ${userId}, now() + (${SESSION_MAX_AGE_SECONDS} * interval '1 second'))`;
  return rawToken;
}

export async function revokeAuthSession(request: ApiRequest) {
  const rawToken = requestCookieValue(request, AUTH_COOKIE);
  if (!rawToken) return;
  const sql = await database();
  await sql`UPDATE inspection_sessions SET revoked_at = now() WHERE token_hash = ${tokenHash(rawToken)} AND revoked_at IS NULL`;
}

export async function revokeUserSessions(userId: string) {
  const sql = await database();
  await sql`UPDATE inspection_sessions SET revoked_at = now() WHERE user_id = ${userId} AND revoked_at IS NULL`;
}

export async function audit(action: string, actorUserId: string | null, targetUserId: string | null, metadata: Record<string, unknown> = {}) {
  const sql = await database();
  await sql`INSERT INTO inspection_audit_log (id, actor_user_id, target_user_id, action, metadata) VALUES (${crypto.randomUUID()}, ${actorUserId}, ${targetUserId}, ${action}, ${JSON.stringify(metadata)})`;
}

async function loginAttemptStatus(username: string) {
  const sql = await database();
  const rows = await sql`SELECT failed_count, window_started_at, locked_until FROM inspection_login_attempts WHERE username = ${username} LIMIT 1`;
  if (!rows.length) return { locked: false, failedCount: 0 };
  const row = rows[0] as Record<string, any>;
  const locked = row.locked_until && new Date(String(row.locked_until)).valueOf() > Date.now();
  return { locked, failedCount: safeInt(row.failed_count) };
}

async function recordLoginFailure(username: string) {
  const sql = await database();
  const existing = await sql`SELECT failed_count, window_started_at, locked_until FROM inspection_login_attempts WHERE username = ${username} LIMIT 1`;
  const row = existing[0] as Record<string, any> | undefined;
  const windowAge = row?.window_started_at ? Date.now() - new Date(String(row.window_started_at)).valueOf() : Number.POSITIVE_INFINITY;
  if (!row || windowAge > 15 * 60 * 1000) {
    await sql`INSERT INTO inspection_login_attempts (username, window_started_at, failed_count, locked_until) VALUES (${username}, now(), 1, NULL) ON CONFLICT (username) DO UPDATE SET window_started_at = now(), failed_count = 1, locked_until = NULL`;
    return;
  }
  const nextCount = safeInt(row.failed_count) + 1;
  const lockedUntil = nextCount >= 5 ? new Date(Date.now() + 15 * 60 * 1000).toISOString() : null;
  await sql`UPDATE inspection_login_attempts SET failed_count = ${nextCount}, locked_until = ${lockedUntil} WHERE username = ${username}`;
}

async function clearLoginFailures(username: string) {
  const sql = await database();
  await sql`DELETE FROM inspection_login_attempts WHERE username = ${username}`;
}

export async function authenticateUser(rawUsername: unknown, rawPassword: unknown) {
  const username = normalizeUsername(rawUsername);
  if (!username) return null;
  const status = await loginAttemptStatus(username);
  if (status.locked) return null;
  const sql = await database();
  const rows = await sql`SELECT id, username, display_name, password_hash, role, is_active, last_login_at FROM inspection_users WHERE username = ${username} LIMIT 1`;
  const row = rows[0] as Record<string, any> | undefined;
  const valid = Boolean(row?.is_active) && await verifyPassword(rawPassword, String(row?.password_hash || ""));
  if (!valid) {
    await recordLoginFailure(username);
    return null;
  }
  await clearLoginFailures(username);
  const userId = String(row?.id || "");
  await sql`UPDATE inspection_users SET last_login_at = now(), updated_at = now() WHERE id = ${userId}`;
  const user = mapUser({ ...row, last_login_at: new Date().toISOString() });
  await audit("login", user.id, user.id, {});
  return { user, token: await createAuthSession(user.id) };
}

export async function registerUser(raw: { username?: unknown; displayName?: unknown; password?: unknown; confirmation?: unknown }) {
  const username = validateUsername(raw.username);
  const displayName = safeText(raw.displayName, 120) || username;
  const password = validatePassword(raw.password);
  if (raw.confirmation !== undefined && String(raw.confirmation ?? "") !== password) throw new Error("两次密码不一致。");
  const passwordHash = await hashPassword(password);
  const sql = await database();
  const id = crypto.randomUUID();
  try {
    await sql`INSERT INTO inspection_users (id, username, display_name, password_hash, role) VALUES (${id}, ${username}, ${displayName}, ${passwordHash}, 'viewer')`;
  } catch (error) {
    if ((error as { code?: string })?.code === "23505") throw new Error("该用户名已经存在。");
    throw error;
  }
  const user = mapUser({ id, username, display_name: displayName, role: "viewer", is_active: true, last_login_at: new Date().toISOString() });
  await sql`UPDATE inspection_users SET last_login_at = now(), updated_at = now() WHERE id = ${id}`;
  await audit("user_registered", null, id, { username, role: "viewer" });
  return { user, token: await createAuthSession(id) };
}

export async function listUsers() {
  const sql = await database();
  const rows = await sql`SELECT id, username, display_name, role, is_active, last_login_at, created_at FROM inspection_users ORDER BY is_active DESC, username ASC`;
  return rows.map((row: Record<string, any>) => ({ ...publicUser(mapUser(row)), isActive: row.is_active !== false, createdAt: row.created_at }));
}

function normalizeRole(value: unknown): UserRole {
  if (value === "admin" || value === "operator" || value === "viewer") return value;
  throw new Error("用户角色不正确。");
}

export async function createUser(actor: AuthUser, raw: { username?: unknown; displayName?: unknown; password?: unknown; role?: unknown }) {
  const username = validateUsername(raw.username);
  const displayName = safeText(raw.displayName, 120) || username;
  const passwordHash = await hashPassword(raw.password);
  const role = normalizeRole(raw.role);
  const sql = await database();
  try {
    const id = crypto.randomUUID();
    await sql`INSERT INTO inspection_users (id, username, display_name, password_hash, role) VALUES (${id}, ${username}, ${displayName}, ${passwordHash}, ${role})`;
    await audit("user_created", actor.id, id, { username, role });
    const rows = await sql`SELECT id, username, display_name, role, is_active, last_login_at, created_at FROM inspection_users WHERE id = ${id}`;
    return { ...publicUser(mapUser(rows[0])), isActive: true, createdAt: rows[0].created_at };
  } catch (error) {
    if ((error as { code?: string })?.code === "23505") throw new Error("该用户名已经存在。");
    throw error;
  }
}

export async function updateUser(actor: AuthUser, id: string, raw: { displayName?: unknown; role?: unknown; isActive?: unknown }) {
  const sql = await database();
  const rows = await sql`SELECT id, username, display_name, role, is_active, last_login_at, created_at FROM inspection_users WHERE id = ${id} LIMIT 1`;
  if (!rows.length) throw new Error("用户不存在。");
  const current = rows[0] as Record<string, any>;
  const nextRole = raw.role === undefined ? normalizeRole(current.role) : normalizeRole(raw.role);
  const nextActive = raw.isActive === undefined ? current.is_active !== false : Boolean(raw.isActive);
  if (actor.id === id && (!nextActive || nextRole !== "admin")) throw new Error("不能停用或降级当前管理员账号。");
  if (current.role === "admin" && (nextRole !== "admin" || !nextActive)) {
    const countRows = await sql`SELECT COUNT(*)::int AS count FROM inspection_users WHERE role = 'admin' AND is_active = true`;
    if (safeInt(countRows[0]?.count) <= 1) throw new Error("系统至少需要保留一个启用中的管理员。");
  }
  const displayName = raw.displayName === undefined ? String(current.display_name) : safeText(raw.displayName, 120) || String(current.username);
  await sql`UPDATE inspection_users SET display_name = ${displayName}, role = ${nextRole}, is_active = ${nextActive}, session_version = session_version + 1, updated_at = now() WHERE id = ${id}`;
  if (!nextActive) await revokeUserSessions(id);
  await audit(nextActive ? "user_updated" : "user_disabled", actor.id, id, { role: nextRole, isActive: nextActive });
  const updated = await sql`SELECT id, username, display_name, role, is_active, last_login_at, created_at FROM inspection_users WHERE id = ${id}`;
  return { ...publicUser(mapUser(updated[0])), isActive: updated[0].is_active !== false, createdAt: updated[0].created_at };
}

export async function resetUserPassword(actor: AuthUser, id: string, rawPassword: unknown) {
  const passwordHash = await hashPassword(rawPassword);
  const sql = await database();
  const rows = await sql`SELECT id FROM inspection_users WHERE id = ${id} LIMIT 1`;
  if (!rows.length) throw new Error("用户不存在。");
  await sql`UPDATE inspection_users SET password_hash = ${passwordHash}, session_version = session_version + 1, updated_at = now() WHERE id = ${id}`;
  await revokeUserSessions(id);
  await audit("password_reset", actor.id, id, {});
}

export async function auditEntries(limit = 100) {
  const sql = await database();
  const safeLimit = Math.min(200, Math.max(1, Math.floor(limit)));
  const rows = await sql`
    SELECT a.id, a.action, a.metadata, a.created_at,
      actor.username AS actor_username, target.username AS target_username
    FROM inspection_audit_log a
    LEFT JOIN inspection_users actor ON actor.id = a.actor_user_id
    LEFT JOIN inspection_users target ON target.id = a.target_user_id
    ORDER BY a.created_at DESC
    LIMIT ${safeLimit}
  `;
  return rows.map((row: Record<string, any>) => ({
    id: row.id,
    action: row.action,
    metadata: row.metadata,
    createdAt: row.created_at,
    actorUsername: row.actor_username || "系统",
    targetUsername: row.target_username || "—",
  }));
}

function mapBatch(row: Record<string, unknown>) {
  return {
    id: row.id,
    businessWeekStart: databaseDate(row.business_week_start),
    businessWeekEnd: databaseDate(row.business_week_end),
    sourceName: row.source_name,
    sourceSha256: row.source_sha256,
    rosterName: row.roster_name,
    rosterSha256: row.roster_sha256,
    rosterSnapshotDate: databaseDate(row.roster_snapshot_date),
    sampleLimit: row.sample_limit,
    eligibleCount: row.eligible_count,
    selectedCount: row.selected_count,
    teacherCount: row.teacher_count,
    unsubmittedCount: row.unsubmitted_count,
    ruleVersion: row.rule_version,
    attempt: row.attempt,
    batchKind: row.batch_kind === "trial" ? "trial" : "formal",
    status: row.status,
    createdAt: row.created_at,
    voidedAt: row.voided_at,
  };
}

async function itemsFor(sql: any, batchId: string) {
  const rows: Array<Record<string, any>> = await sql`SELECT position, teacher_name, teacher_email, student_name, student_id, course_id, lesson_start, lesson_end, submitted_value, product_group, campus, project_group, selection_reason FROM inspection_items WHERE batch_id = ${batchId} ORDER BY position`;
  return rows.map((row) => ({
    position: row.position,
    teacherName: row.teacher_name,
    teacherEmail: row.teacher_email,
    studentName: row.student_name,
    studentId: row.student_id,
    courseId: row.course_id,
    lessonStart: row.lesson_start,
    lessonEnd: row.lesson_end,
    submittedValue: row.submitted_value,
    productGroup: row.product_group,
    campus: row.campus,
    projectGroup: row.project_group,
    selectionReason: row.selection_reason,
  }));
}

export interface HistoryFilters {
  query: string;
  from: string;
  to: string;
  batchKind: "formal" | "trial" | "all";
  status: "active" | "voided" | "all";
  page: number;
  pageSize: number;
  sort: "courses" | "weeks" | "recent";
}

function queryText(request: ApiRequest, key: string) {
  const value = request.query?.[key];
  return safeText(Array.isArray(value) ? value[0] : value, 120);
}

function validDateQuery(value: string) {
  return /^\d{4}-\d{2}-\d{2}$/u.test(value) ? value : "";
}

export function parseHistoryFilters(request: ApiRequest, defaults: Partial<HistoryFilters> = {}): HistoryFilters {
  const rawKind = queryText(request, "batchKind");
  const rawStatus = queryText(request, "status");
  const rawSort = queryText(request, "sort");
  const page = Math.max(1, Math.min(100000, safeInt(queryText(request, "page") || String(defaults.page || 1), defaults.page || 1)));
  const pageSize = Math.max(1, Math.min(100, safeInt(queryText(request, "pageSize") || String(defaults.pageSize || 50), defaults.pageSize || 50)));
  return {
    query: queryText(request, "q") || defaults.query || "",
    from: validDateQuery(queryText(request, "from")) || defaults.from || "",
    to: validDateQuery(queryText(request, "to")) || defaults.to || "",
    batchKind: rawKind === "trial" || rawKind === "formal" || rawKind === "all" ? rawKind : defaults.batchKind || "all",
    status: rawStatus === "active" || rawStatus === "voided" || rawStatus === "all" ? rawStatus : defaults.status || "all",
    page,
    pageSize,
    sort: rawSort === "weeks" || rawSort === "recent" || rawSort === "courses" ? rawSort : defaults.sort || "recent",
  };
}

export async function currentBatch(weekStart?: string, batchKind = "formal") {
  const sql = await database();
  const rows = weekStart
    ? await sql`SELECT * FROM inspection_batches WHERE status = 'active' AND business_week_start = ${weekStart} AND batch_kind = ${batchKind} LIMIT 1`
    : await sql`SELECT * FROM inspection_batches WHERE status = 'active' AND batch_kind = ${batchKind} ORDER BY business_week_start DESC LIMIT 1`;
  if (!rows.length) return null;
  return { ...mapBatch(rows[0]), items: await itemsFor(sql, String(rows[0].id)) };
}

export async function historyBatches(filters: HistoryFilters = {
  query: "",
  from: "",
  to: "",
  batchKind: "all",
  status: "all",
  page: 1,
  pageSize: 50,
  sort: "recent",
}) {
  const sql = await database();
  const like = `%${filters.query.toLocaleLowerCase()}%`;
  const kindAll = filters.batchKind === "all";
  const statusAll = filters.status === "all";
  const fromDate = filters.from || null;
  const toDate = filters.to || null;
  const fromAll = !fromDate;
  const toAll = !toDate;
  const offset = (filters.page - 1) * filters.pageSize;
  const rows = await sql`
    SELECT *, COUNT(*) OVER() AS total_count
    FROM inspection_batches
    WHERE (${kindAll} OR batch_kind = ${filters.batchKind})
      AND (${statusAll} OR status = ${filters.status})
      AND (${fromAll} OR business_week_start >= ${fromDate}::date)
      AND (${toAll} OR business_week_start <= ${toDate}::date)
      AND (
        ${!filters.query}
        OR LOWER(source_name) LIKE ${like}
        OR LOWER(roster_name) LIKE ${like}
        OR TO_CHAR(business_week_start, 'YYYY-MM-DD') LIKE ${like}
        OR TO_CHAR(business_week_end, 'YYYY-MM-DD') LIKE ${like}
      )
    ORDER BY created_at DESC
    LIMIT ${filters.pageSize} OFFSET ${offset}
  `;
  return {
    batches: rows.map((row) => mapBatch(row)),
    total: rows.length ? safeInt(rows[0].total_count) : 0,
  };
}

function teacherKeyExpression() {
  return `COALESCE(NULLIF(LOWER(TRIM(i.teacher_email)), ''), 'name:' || LOWER(TRIM(i.teacher_name)))`;
}

function teacherKeyValue(key: string) {
  const value = safeText(key, 240).toLocaleLowerCase();
  if (value.startsWith("email:")) return { kind: "email", value: value.slice(6) };
  if (value.startsWith("name:")) return { kind: "name", value: value.slice(5) };
  return { kind: "email", value };
}

function mapInspectionItem(row: Record<string, any>) {
  return {
    batchId: row.batch_id || row.id,
    businessWeekStart: databaseDate(row.business_week_start),
    businessWeekEnd: databaseDate(row.business_week_end),
    attempt: row.attempt,
    status: row.status,
    batchKind: row.batch_kind === "trial" ? "trial" : "formal",
    position: row.position,
    teacherName: row.teacher_name,
    teacherEmail: row.teacher_email,
    studentName: row.student_name,
    studentId: row.student_id,
    courseId: row.course_id,
    lessonStart: row.lesson_start,
    lessonEnd: row.lesson_end,
    submittedValue: row.submitted_value,
    productGroup: row.product_group,
    campus: row.campus,
    projectGroup: row.project_group,
    selectionReason: row.selection_reason,
  };
}

export async function teacherSummaries(filters: HistoryFilters) {
  const sql = await database();
  const like = `%${filters.query.toLocaleLowerCase()}%`;
  const kindAll = filters.batchKind === "all";
  const statusAll = filters.status === "all";
  const fromDate = filters.from || null;
  const toDate = filters.to || null;
  const fromAll = !fromDate;
  const toAll = !toDate;
  const offset = (filters.page - 1) * filters.pageSize;
  const keyExpression = teacherKeyExpression();
  const rows = await sql`
    WITH base AS (
      SELECT ${sql.unsafe(keyExpression)} AS teacher_key,
        i.teacher_name, i.teacher_email, i.student_name, i.course_id,
        i.submitted_value, i.product_group, i.campus, i.project_group,
        b.id AS batch_id, b.business_week_start, b.business_week_end,
        b.batch_kind, b.status, b.created_at
      FROM inspection_items i
      JOIN inspection_batches b ON b.id = i.batch_id
      WHERE (${kindAll} OR b.batch_kind = ${filters.batchKind})
        AND (${statusAll} OR b.status = ${filters.status})
        AND (${fromAll} OR b.business_week_start >= ${fromDate}::date)
        AND (${toAll} OR b.business_week_start <= ${toDate}::date)
    ), matched_keys AS (
      SELECT DISTINCT teacher_key
      FROM base
      WHERE ${!filters.query}
        OR LOWER(COALESCE(teacher_name, '')) LIKE ${like}
        OR LOWER(COALESCE(teacher_email, '')) LIKE ${like}
        OR LOWER(COALESCE(product_group, '')) LIKE ${like}
        OR LOWER(COALESCE(campus, '')) LIKE ${like}
        OR LOWER(COALESCE(project_group, '')) LIKE ${like}
    ), teacher_stats AS (
      SELECT b.teacher_key,
        (ARRAY_AGG(b.teacher_name ORDER BY b.business_week_start DESC, b.created_at DESC))[1] AS teacher_name,
        (ARRAY_AGG(b.teacher_email ORDER BY b.business_week_start DESC, b.created_at DESC))[1] AS teacher_email,
        COUNT(*)::int AS course_count,
        COUNT(DISTINCT b.batch_id)::int AS batch_count,
        COUNT(DISTINCT b.business_week_start)::int AS week_count,
        COUNT(*) FILTER (WHERE LOWER(TRIM(b.submitted_value)) IN ('否', 'no', '未生成'))::int AS unsubmitted_count,
        COUNT(*) FILTER (WHERE b.batch_kind = 'formal')::int AS formal_course_count,
        COUNT(*) FILTER (WHERE b.batch_kind = 'trial')::int AS trial_course_count,
        MAX(b.business_week_start) AS last_week,
        STRING_AGG(DISTINCT NULLIF(b.project_group, ''), '、') AS project_groups,
        STRING_AGG(DISTINCT NULLIF(b.product_group, ''), '、') AS product_groups
      FROM base b
      JOIN matched_keys m ON m.teacher_key = b.teacher_key
      GROUP BY b.teacher_key
    )
    SELECT *, COUNT(*) OVER() AS total_count
    FROM teacher_stats
    ORDER BY
      CASE WHEN ${filters.sort === "weeks"} THEN week_count END DESC NULLS LAST,
      CASE WHEN ${filters.sort === "recent"} THEN last_week END DESC NULLS LAST,
      course_count DESC, teacher_name ASC
    LIMIT ${filters.pageSize} OFFSET ${offset}
  `;
  return {
    teachers: rows.map((row: Record<string, any>) => ({
      teacherKey: row.teacher_key,
      teacherName: row.teacher_name || "未命名教师",
      teacherEmail: row.teacher_email || "",
      courseCount: row.course_count,
      batchCount: row.batch_count,
      weekCount: row.week_count,
      unsubmittedCount: row.unsubmitted_count,
      formalCourseCount: row.formal_course_count,
      trialCourseCount: row.trial_course_count,
      lastWeek: databaseDate(row.last_week),
      projectGroups: row.project_groups || "",
      productGroups: row.product_groups || "",
    })),
    total: rows.length ? safeInt(rows[0].total_count) : 0,
  };
}

export async function teacherDetail(rawKey: string, filters: HistoryFilters) {
  const sql = await database();
  const key = teacherKeyValue(rawKey);
  const kindAll = filters.batchKind === "all";
  const statusAll = filters.status === "all";
  const fromDate = filters.from || null;
  const toDate = filters.to || null;
  const fromAll = !fromDate;
  const toAll = !toDate;
  const offset = (filters.page - 1) * filters.pageSize;
  const isEmail = key.kind === "email";
  const rows = await sql`
    SELECT i.*, b.business_week_start, b.business_week_end, b.attempt, b.batch_kind, b.status, b.created_at
    FROM inspection_items i
    JOIN inspection_batches b ON b.id = i.batch_id
    WHERE ((${isEmail} AND LOWER(TRIM(i.teacher_email)) = ${key.value})
      OR (${!isEmail} AND LOWER(TRIM(i.teacher_email)) = '' AND LOWER(TRIM(i.teacher_name)) = ${key.value}))
      AND (${kindAll} OR b.batch_kind = ${filters.batchKind})
      AND (${statusAll} OR b.status = ${filters.status})
      AND (${fromAll} OR b.business_week_start >= ${fromDate}::date)
      AND (${toAll} OR b.business_week_start <= ${toDate}::date)
    ORDER BY b.business_week_start DESC, b.attempt DESC, i.position
    LIMIT ${filters.pageSize} OFFSET ${offset}
  `;
  const weekRows = await sql`
    SELECT b.id, b.business_week_start, b.business_week_end, b.attempt, b.batch_kind, b.status,
      COUNT(*)::int AS course_count,
      COUNT(*) FILTER (WHERE LOWER(TRIM(i.submitted_value)) IN ('否', 'no', '未生成'))::int AS unsubmitted_count
    FROM inspection_items i
    JOIN inspection_batches b ON b.id = i.batch_id
    WHERE ((${isEmail} AND LOWER(TRIM(i.teacher_email)) = ${key.value})
      OR (${!isEmail} AND LOWER(TRIM(i.teacher_email)) = '' AND LOWER(TRIM(i.teacher_name)) = ${key.value}))
      AND (${kindAll} OR b.batch_kind = ${filters.batchKind})
      AND (${statusAll} OR b.status = ${filters.status})
      AND (${fromAll} OR b.business_week_start >= ${fromDate}::date)
      AND (${toAll} OR b.business_week_start <= ${toDate}::date)
    GROUP BY b.id
    ORDER BY b.business_week_start DESC, b.attempt DESC
  `;
  const summaryRows = await sql`
    SELECT
      MIN(i.teacher_name) AS teacher_name,
      MIN(i.teacher_email) AS teacher_email,
      COUNT(*)::int AS course_count,
      COUNT(DISTINCT b.id)::int AS batch_count,
      COUNT(DISTINCT b.business_week_start)::int AS week_count,
      COUNT(*) FILTER (WHERE LOWER(TRIM(i.submitted_value)) IN ('否', 'no', '未生成'))::int AS unsubmitted_count
    FROM inspection_items i
    JOIN inspection_batches b ON b.id = i.batch_id
    WHERE ((${isEmail} AND LOWER(TRIM(i.teacher_email)) = ${key.value})
      OR (${!isEmail} AND LOWER(TRIM(i.teacher_email)) = '' AND LOWER(TRIM(i.teacher_name)) = ${key.value}))
      AND (${kindAll} OR b.batch_kind = ${filters.batchKind})
      AND (${statusAll} OR b.status = ${filters.status})
      AND (${fromAll} OR b.business_week_start >= ${fromDate}::date)
      AND (${toAll} OR b.business_week_start <= ${toDate}::date)
  `;
  const first = summaryRows[0] as Record<string, any> | undefined;
  return {
    teacher: first && first.course_count ? {
      teacherKey: rawKey,
      teacherName: first.teacher_name || "未命名教师",
      teacherEmail: first.teacher_email || "",
      courseCount: safeInt(first.course_count),
      batchCount: safeInt(first.batch_count),
      weekCount: safeInt(first.week_count),
      unsubmittedCount: safeInt(first.unsubmitted_count),
    } : null,
    weeks: weekRows.map((row: Record<string, any>) => ({
      batchId: row.id,
      businessWeekStart: databaseDate(row.business_week_start),
      businessWeekEnd: databaseDate(row.business_week_end),
      attempt: row.attempt,
      batchKind: row.batch_kind === "trial" ? "trial" : "formal",
      status: row.status,
      courseCount: row.course_count,
      unsubmittedCount: row.unsubmitted_count,
    })),
    items: rows.map(mapInspectionItem),
    total: safeInt(first?.course_count),
    page: filters.page,
    pageSize: filters.pageSize,
  };
}

export async function historyBatch(id: string, page = 1, pageSize = 100) {
  const sql = await database();
  const rows = await sql`SELECT * FROM inspection_batches WHERE id = ${id} LIMIT 1`;
  if (!rows.length) return null;
  const safePageNumber = Math.max(1, page);
  const safePageSize = Math.min(100, Math.max(1, pageSize));
  const offset = (safePageNumber - 1) * safePageSize;
  const [itemRows, countRows] = await Promise.all([
    sql`SELECT position, teacher_name, teacher_email, student_name, student_id, course_id, lesson_start, lesson_end, submitted_value, product_group, campus, project_group, selection_reason FROM inspection_items WHERE batch_id = ${id} ORDER BY position LIMIT ${safePageSize} OFFSET ${offset}`,
    sql`SELECT COUNT(*)::int AS total_count FROM inspection_items WHERE batch_id = ${id}`,
  ]);
  return {
    ...mapBatch(rows[0]),
    items: itemRows.map((row: Record<string, any>) => ({
      position: row.position,
      teacherName: row.teacher_name,
      teacherEmail: row.teacher_email,
      studentName: row.student_name,
      studentId: row.student_id,
      courseId: row.course_id,
      lessonStart: row.lesson_start,
      lessonEnd: row.lesson_end,
      submittedValue: row.submitted_value,
      productGroup: row.product_group,
      campus: row.campus,
      projectGroup: row.project_group,
      selectionReason: row.selection_reason,
    })),
    total: safeInt(countRows[0]?.total_count),
    page: safePageNumber,
    pageSize: safePageSize,
  };
}

export async function monthlyInspection(month: string) {
  const sql = await database();
  const start = new Date(`${month}-01T00:00:00Z`);
  const end = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth() + 1, 1));
  const nextMonth = `${end.getUTCFullYear()}-${String(end.getUTCMonth() + 1).padStart(2, "0")}-01`;
  const rows = await sql`
    SELECT b.id, b.business_week_start, b.business_week_end, b.attempt,
      i.position, i.teacher_name, i.teacher_email, i.student_name, i.student_id,
      i.course_id, i.lesson_start, i.lesson_end, i.submitted_value,
      i.product_group, i.campus, i.project_group, i.selection_reason
    FROM inspection_batches b
    JOIN inspection_items i ON i.batch_id = b.id
    WHERE b.status = 'active' AND b.batch_kind = 'formal'
      AND b.business_week_start >= ${month + "-01"}
      AND b.business_week_start < ${nextMonth}
    ORDER BY b.business_week_start, b.attempt, i.position
  `;
  return {
    month,
    items: rows.map((row: Record<string, unknown>) => ({
      batchId: row.id,
      businessWeekStart: databaseDate(row.business_week_start),
      businessWeekEnd: databaseDate(row.business_week_end),
      attempt: row.attempt,
      position: row.position,
      teacherName: row.teacher_name,
      teacherEmail: row.teacher_email,
      studentName: row.student_name,
      studentId: row.student_id,
      courseId: row.course_id,
      lessonStart: row.lesson_start,
      lessonEnd: row.lesson_end,
      submittedValue: row.submitted_value,
      productGroup: row.product_group,
      campus: row.campus,
      projectGroup: row.project_group,
      selectionReason: row.selection_reason,
    })),
  };
}

async function insertBatch(sql: any, payload: ReturnType<typeof assertHistoryPayload>, id: string) {
  const { batch, items } = payload;
  await sql`INSERT INTO inspection_batches (id, business_week_start, business_week_end, source_name, source_sha256, roster_name, roster_sha256, roster_snapshot_date, sample_limit, eligible_count, selected_count, teacher_count, unsubmitted_count, rule_version, attempt, batch_kind, status) VALUES (${id}, ${batch.businessWeekStart}, ${batch.businessWeekEnd}, ${batch.sourceName}, ${batch.sourceSha256}, ${batch.rosterName}, ${batch.rosterSha256}, ${batch.rosterSnapshotDate || null}, ${batch.sampleLimit}, ${batch.eligibleCount}, ${batch.selectedCount}, ${batch.teacherCount}, ${batch.unsubmittedCount}, ${batch.ruleVersion}, ${batch.attempt}, ${batch.batchKind}, 'active')`;
  for (const item of items) {
    await sql`INSERT INTO inspection_items (batch_id, position, teacher_name, teacher_email, student_name, student_id, course_id, lesson_start, lesson_end, submitted_value, product_group, campus, project_group, selection_reason) VALUES (${id}, ${item.position}, ${item.teacherName}, ${item.teacherEmail}, ${item.studentName}, ${item.studentId}, ${item.courseId}, ${item.lessonStart}, ${item.lessonEnd}, ${item.submittedValue}, ${item.productGroup}, ${item.campus}, ${item.projectGroup}, ${item.selectionReason})`;
  }
  const rows = await sql`SELECT * FROM inspection_batches WHERE id = ${id}`;
  return { ...mapBatch(rows[0]), items };
}

export async function createBatch(rawPayload: unknown, actorUserId: string | null = null) {
  const payload = assertHistoryPayload(rawPayload);
  const sql = await database();
  const existing = await currentBatch(payload.batch.businessWeekStart, payload.batch.batchKind);
  if (existing) {
    if (existing.sourceSha256 !== payload.batch.sourceSha256 || existing.rosterSha256 !== payload.batch.rosterSha256) {
      return { conflict: true, reused: false, batch: existing };
    }
    return { reused: true, batch: existing };
  }
  try {
    const result = { reused: false, batch: await insertBatch(sql, payload, crypto.randomUUID()) };
    if (actorUserId) await audit("inspection_batch_created", actorUserId, null, { batchId: result.batch.id, businessWeekStart: payload.batch.businessWeekStart, batchKind: payload.batch.batchKind });
    return result;
  } catch (error) {
    if ((error as { code?: string })?.code === "23505") {
      const concurrent = await currentBatch(payload.batch.businessWeekStart, payload.batch.batchKind);
      if (concurrent) return { reused: true, batch: concurrent };
    }
    throw error;
  }
}

export async function replaceBatch(id: string, rawPayload: unknown, actorUserId: string | null = null) {
  const payload = assertHistoryPayload(rawPayload);
  const sql = await database();
  const rows = await sql`SELECT id FROM inspection_batches WHERE id = ${id} AND status = 'active' LIMIT 1`;
  if (!rows.length) throw new Error("当前抽检批次不存在或已经被替换。");
  await sql`UPDATE inspection_batches SET status = 'voided', voided_at = now() WHERE id = ${id}`;
  const result = { reused: false, batch: await insertBatch(sql, payload, crypto.randomUUID()) };
  if (actorUserId) await audit("inspection_batch_replaced", actorUserId, null, { oldBatchId: id, batchId: result.batch.id, businessWeekStart: payload.batch.businessWeekStart, batchKind: payload.batch.batchKind });
  return result;
}

export { SESSION_MAX_AGE_SECONDS, cookieHeader, env };

// Vercel treats every TypeScript file below /api as a function entrypoint.
// This module is shared by the real handlers, but the fallback keeps its
// accidental direct URL non-operational if the platform packages it too.
export default function sharedModuleHandler(_request: ApiRequest, response: ApiResponse) {
  response.status(404).json({ error: "未找到抽检接口。" });
}
