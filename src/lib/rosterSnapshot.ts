import { emailValue } from "../worker/utils.js";
import { normalizedRoleExcludedEmails } from "../worker/inspectionRoleRules.js";
import type { RosterInfo } from "../worker/inspectionTypes.js";

export interface RosterSnapshotMeta {
  snapshotDate: string;
  sourceFile: string;
  sourceSha256: string;
  rowCount: number;
}

/** Public projection only: no names, employee IDs, organization or original rows. */
export interface InspectionRosterSnapshot extends RosterSnapshotMeta {
  matchedEmailRows: number;
  emails: string[];
  roleExcludedEmails: string[];
  unsubmittedExtraExcludedEmails: string[];
}

export function validateRosterMeta(value: unknown): RosterSnapshotMeta {
  const meta = value as RosterSnapshotMeta | null;
  const date = new Date(`${meta?.snapshotDate}T00:00:00Z`);
  if (!meta || !/^\d{4}-\d{2}-\d{2}$/u.test(meta.snapshotDate)
    || !Number.isFinite(date.getTime()) || date.toISOString().slice(0, 10) !== meta.snapshotDate
    || typeof meta.sourceFile !== "string" || !meta.sourceFile.trim()
    || !/^[a-f0-9]{64}$/u.test(meta.sourceSha256)
    || !Number.isSafeInteger(meta.rowCount) || meta.rowCount < 1) {
    throw new Error("在职版本信息无效，处理已停止。");
  }
  return { snapshotDate: meta.snapshotDate, sourceFile: meta.sourceFile, sourceSha256: meta.sourceSha256, rowCount: meta.rowCount };
}

export function rosterFromSnapshot(value: unknown): RosterInfo {
  const meta = validateRosterMeta(value);
  const asset = value as InspectionRosterSnapshot;
  const lists = [asset.emails, asset.roleExcludedEmails, asset.unsubmittedExtraExcludedEmails];
  if (lists.some((list) => !Array.isArray(list) || list.some((email) => typeof email !== "string" || emailValue(email) !== email)
    || new Set(list).size !== list.length)
    || !asset.emails.length || !Number.isSafeInteger(asset.matchedEmailRows)
    || asset.matchedEmailRows < asset.emails.length || asset.matchedEmailRows > meta.rowCount
    || asset.roleExcludedEmails.some((email) => !asset.emails.includes(email))
    || asset.unsubmittedExtraExcludedEmails.some((email) => !asset.emails.includes(email))) {
    throw new Error("在职名单返回格式无效，处理已停止。");
  }
  return {
    emails: new Set(asset.emails),
    roleExcludedEmails: normalizedRoleExcludedEmails(asset.roleExcludedEmails),
    unsubmittedExtraExcludedEmails: normalizedRoleExcludedEmails(asset.unsubmittedExtraExcludedEmails),
    sourceName: `${meta.sourceFile}（统一接口）`,
    snapshotDate: meta.snapshotDate,
    rowCount: meta.rowCount,
    matchedEmailRows: asset.matchedEmailRows,
  };
}

export async function loadLatestInspectionRoster(): Promise<InspectionRosterSnapshot> {
  const response = await fetch("/api/roster", { cache: "no-store", credentials: "same-origin", signal: AbortSignal.timeout(30_000) });
  const body = await response.json().catch(() => null);
  if (!response.ok) throw new Error(body?.error || "最新在职名单读取失败，请重试或手动上传指定版本。处理已停止。");
  rosterFromSnapshot(body);
  return body as InspectionRosterSnapshot;
}
