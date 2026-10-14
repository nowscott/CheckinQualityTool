import { validateRosterMeta, type InspectionRosterSnapshot } from "../lib/rosterSnapshot.js";
import { emailValue } from "../worker/utils.js";
import { inspectionRoleExcludedEmails, inspectionUnsubmittedExtraExcludedEmails } from "../worker/inspectionRoleRules.js";

const SERVICE_URL = "https://zaizhi.0211120.xyz/api/roster/";
const SOURCE_FIELDS = ["姓名", "员工ID", "师训组长", "师训主管/助理主管", "科目师训负责人", "教研组", "用工类别", "入职日期", "司龄（年）", "教师状态", "校区", "区域", "非课状态", "可带年级", "岗位短描述", "性别", "教师编码", "教师评级", "邮箱"];

export function projectInspectionRoster(value: unknown): InspectionRosterSnapshot {
  const meta = validateRosterMeta(value);
  const rows = (value as { rows?: unknown }).rows;
  if (!Array.isArray(rows) || rows.length !== meta.rowCount) throw new Error("统一在职名单人数校验失败。");
  const keys = new Set<string>();
  const emails = new Set<string>();
  const roleRows: Array<{ email: string; roleDescription: unknown }> = [];
  let matchedEmailRows = 0;
  for (const row of rows) {
    if (!row || typeof row !== "object" || SOURCE_FIELDS.some((field) => !Object.hasOwn(row, field))
      || typeof row.personKey !== "string" || !row.personKey.trim() || keys.has(row.personKey)) {
      throw new Error("统一在职名单字段或人员键校验失败。");
    }
    keys.add(row.personKey);
    const email = emailValue(row["邮箱"]);
    if (!email) continue;
    matchedEmailRows += 1;
    emails.add(email);
    roleRows.push({ email, roleDescription: row["岗位短描述"] });
  }
  if (!emails.size) throw new Error("统一在职名单没有有效邮箱。");
  return {
    ...meta,
    matchedEmailRows,
    emails: [...emails].sort(),
    roleExcludedEmails: [...inspectionRoleExcludedEmails(roleRows)].sort(),
    unsubmittedExtraExcludedEmails: [...inspectionUnsubmittedExtraExcludedEmails(roleRows)].sort(),
  };
}

export async function readLatestInspectionRoster(token: string | undefined, fetcher: typeof fetch = fetch): Promise<InspectionRosterSnapshot> {
  if (!token?.trim()) throw new Error("服务端尚未配置在职接口，请联系管理员或手动上传指定版本。");
  const read = async (endpoint: "latest" | "meta") => {
    try {
      const response = await fetcher(`${SERVICE_URL}${endpoint}`, {
        headers: { Authorization: `Bearer ${token}` }, cache: "no-store", signal: AbortSignal.timeout(12_000), redirect: "error",
      });
      if (!response.ok) {
        if (response.status === 401 || response.status === 403) throw new Error("在职接口认证失败，请联系管理员检查服务端配置。");
        throw new Error("统一在职接口暂不可用，请重试或手动上传指定版本。");
      }
      return await response.json();
    } catch (error) {
      // Never relay upstream bodies, URLs, request headers, or exception details.
      if (error instanceof Error && /^(在职接口认证失败|统一在职接口暂不可用)/u.test(error.message)) throw error;
      throw new Error("统一在职接口读取失败，请重试或手动上传指定版本。");
    }
  };
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const [latest, version] = await Promise.all([read("latest"), read("meta")]);
    const snapshot = projectInspectionRoster(latest);
    const meta = validateRosterMeta(version);
    if (snapshot.snapshotDate === meta.snapshotDate && snapshot.sourceSha256 === meta.sourceSha256
      && snapshot.rowCount === meta.rowCount && snapshot.sourceFile === meta.sourceFile) return snapshot;
  }
  throw new Error("在职版本正在更新，名单与版本信息不一致，请重新开始处理。");
}
