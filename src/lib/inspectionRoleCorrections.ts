import { emailValue } from "../worker/utils.js";
import type { RosterInfo } from "../worker/inspectionTypes.js";

export const ROLE_CORRECTIONS_SHEET_ID = "0y559g";
export const ROLE_CORRECTIONS_URL = `https://docs.qq.com/sheet/DUGFUT3dOckJVeXFm?tab=${ROLE_CORRECTIONS_SHEET_ID}`;

/** Only the email/eligibility projection is sent to the browser. */
export interface InspectionRoleCorrections {
  managerEmails: string[];
  supervisorEmails: string[];
  teacherEmails: string[];
}

export function validateInspectionRoleCorrections(value: unknown): InspectionRoleCorrections {
  const body = value as InspectionRoleCorrections | null;
  const seen = new Set<string>();
  if (!body || [body.managerEmails, body.supervisorEmails, body.teacherEmails].some((list) =>
    !Array.isArray(list) || list.some((email) => {
      if (typeof email !== "string" || !email || emailValue(email) !== email || seen.has(email)) return true;
      seen.add(email);
      return false;
    }),
  )) throw new Error("岗位纠偏名单格式无效，请检查腾讯文档后重新生成。");
  return body;
}

export async function loadInspectionRoleCorrections(): Promise<InspectionRoleCorrections> {
  const response = await fetch("/api/role-corrections", { cache: "no-store", signal: AbortSignal.timeout(20_000) });
  const body = await response.json().catch(() => null);
  if (!response.ok) throw new Error(body?.error || "岗位纠偏表读取失败，请稍后重试。");
  return validateInspectionRoleCorrections(body);
}

export function applyInspectionRoleCorrections(roster: RosterInfo, corrections: InspectionRoleCorrections): RosterInfo {
  validateInspectionRoleCorrections(corrections);
  const roleExcludedEmails = new Set(roster.roleExcludedEmails);
  const unsubmittedExtraExcludedEmails = new Set(roster.unsubmittedExtraExcludedEmails);
  for (const role of ["managerEmails", "supervisorEmails", "teacherEmails"] as const) {
    const emails = corrections[role];
    for (const email of emails) {
      if (!roster.emails.has(email)) continue;
      roleExcludedEmails.delete(email);
      unsubmittedExtraExcludedEmails.delete(email);
      if (role === "managerEmails") roleExcludedEmails.add(email);
      if (role === "supervisorEmails") unsubmittedExtraExcludedEmails.add(email);
    }
  }
  return { ...roster, roleExcludedEmails, unsubmittedExtraExcludedEmails };
}
