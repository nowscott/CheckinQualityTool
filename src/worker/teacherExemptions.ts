import { normalizeMatchText } from "./utils";

const TEACHER_SUFFIX = decodeURIComponent("%E8%80%81%E5%B8%88");
const EXEMPT_TEACHERS = new Set<string>();

export function normalizeTeacherName(value: unknown) {
  return normalizeMatchText(value)
    .replace(/\s+/g, "")
    .replace(/[0-9０-９]+$/u, "")
    .replace(new RegExp(`${TEACHER_SUFFIX}$`, "u"), "");
}

export function isTeacherExempt(value: unknown) {
  const teacher = normalizeTeacherName(value);
  return Boolean(teacher && EXEMPT_TEACHERS.has(teacher));
}
