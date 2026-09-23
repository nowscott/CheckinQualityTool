import { emailValue, text } from "./utils";

export function hasExcludedInspectionRole(roleDescription: unknown) {
  return /经理/u.test(text(roleDescription));
}

export function inspectionRoleExcludedEmails(
  rows: Iterable<{ email: unknown; roleDescription?: unknown }>,
) {
  const excluded = new Set<string>();
  for (const row of rows) {
    const email = emailValue(row.email);
    if (email && hasExcludedInspectionRole(row.roleDescription)) excluded.add(email);
  }
  return excluded;
}

export function normalizedRoleExcludedEmails(emails: Iterable<unknown>) {
  const normalized = new Set<string>();
  for (const value of emails) {
    const email = emailValue(value);
    if (email) normalized.add(email);
  }
  return normalized;
}
