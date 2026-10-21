import { parseTencentCsv } from "./tencentDocs.js";
import { emailValue } from "../worker/utils.js";
import type { InspectionRoleCorrections } from "../lib/inspectionRoleCorrections.js";

const normalize = (value: string) => String(value || "").normalize("NFKC").trim();

export function parseInspectionRoleCorrections(csv: string): InspectionRoleCorrections {
  const rows = parseTencentCsv(csv);
  const headerIndex = rows.findIndex((row) => ["教师邮箱", "纠正岗位", "是否启用"].every((field) => row.map(normalize).includes(field)));
  if (headerIndex < 0) throw new Error("岗位纠偏页需要教师邮箱、纠正岗位、是否启用三列。");
  const headers = rows[headerIndex].map(normalize);
  const emailIndex = headers.indexOf("教师邮箱");
  const roleIndex = headers.indexOf("纠正岗位");
  const enabledIndex = headers.indexOf("是否启用");
  const enabled = new Map<string, string>();
  const result: InspectionRoleCorrections = { managerEmails: [], supervisorEmails: [], teacherEmails: [] };
  for (let index = headerIndex + 1; index < rows.length; index += 1) {
    const row = rows[index];
    if (!row.some((cell) => normalize(cell))) continue;
    const status = normalize(row[enabledIndex]);
    if (status === "否") continue;
    if (status !== "是") throw new Error(`岗位纠偏第 ${index + 1} 行“是否启用”请填写是或否。`);
    const email = normalize(row[emailIndex]).toLowerCase();
    const role = normalize(row[roleIndex]);
    if (!email || emailValue(email) !== email) throw new Error(`岗位纠偏第 ${index + 1} 行教师邮箱无效。`);
    if (!["教师", "主管", "助理主管", "经理"].includes(role)) throw new Error(`岗位纠偏第 ${index + 1} 行纠正岗位请填写教师、主管、助理主管或经理。`);
    const previous = enabled.get(email);
    if (previous && previous !== role) throw new Error(`岗位纠偏第 ${index + 1} 行与同一邮箱的其他启用记录冲突。`);
    enabled.set(email, role);
  }
  for (const [email, role] of enabled) {
    if (role === "经理") result.managerEmails.push(email);
    else if (role === "教师") result.teacherEmails.push(email);
    else result.supervisorEmails.push(email);
  }
  for (const emails of Object.values(result)) emails.sort();
  return result;
}
