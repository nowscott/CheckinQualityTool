import type { CellValue, CleanStudentName } from "./types";
import { weekOfMonthFromDate } from "../lib/week.js";

export function text(value: unknown) {
  return value == null ? "" : String(value).trim();
}

export function emailValue(value: unknown) {
  const match = text(value).toLowerCase().match(/[\w.+-]+@[\w.-]+\.[a-z]{2,}/);
  return match ? match[0] : "";
}

export function cleanStudentName(value: unknown): CleanStudentName {
  const original = text(value);
  let cleaned = original.replace(/\s+/g, "");
  if (/^[\u4e00-\u9fff]{4}$/u.test(cleaned) && cleaned.endsWith("一")) {
    cleaned = cleaned.slice(0, -1);
  }
  const suffixRules = [
    /[（(【\[].+?[）)】\]]$/u,
    /[0-9０-９]+$/u,
    /重复$/u,
    /(?:学员|同学|学生)$/u,
    /(?:家长|妈妈|爸爸|姐姐|哥哥|妹妹|弟弟)$/u,
    /(?:初|高)(?:[一二三]|[1-3１-３])(?:年级)?$/u,
  ];
  let changed = true;
  while (cleaned && changed) {
    changed = false;
    for (const suffix of suffixRules) {
      const next = cleaned.replace(suffix, "");
      if (next !== cleaned) {
        cleaned = next;
        changed = true;
      }
    }
  }
  return {
    original,
    cleaned,
    note: original.replace(/\s+/g, "") === cleaned ? "" : `${original} → ${cleaned || "空"}`,
  };
}

export function normalizeMatchText(value: unknown) {
  return text(value).toLocaleLowerCase("zh-CN");
}

export function normalizeTeacherName(value: unknown) {
  return normalizeMatchText(value)
    .replace(/\s+/g, "")
    .replace(/[0-9０-９]+$/u, "")
    .replace(/老师$/u, "");
}

export function excelDate(value: CellValue) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    const pad = (number: number) => String(number).padStart(2, "0");
    return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
  }
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }
  const match = text(value).match(/\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/);
  return match ? match[0].replace(/[/.]/g, "-") : text(value);
}

export function excelTime(value: CellValue) {
  const pad = (number: number) => String(number).padStart(2, "0");
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    return `${pad(value.getHours())}:${pad(value.getMinutes())}`;
  }
  const match = text(value).match(/(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?(?:\s|$)/);
  if (!match) return text(value);
  return `${pad(Number(match[1]))}:${match[2]}`;
}

export function displayValue(value: unknown) {
  if (!(value instanceof Date) || Number.isNaN(value.valueOf())) return text(value);
  const pad = (number: number) => String(number).padStart(2, "0");
  return (
    `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())} ` +
    `${pad(value.getHours())}:${pad(value.getMinutes())}:${pad(value.getSeconds())}`
  );
}

export function sortDate(value: CellValue) {
  const timestamp = Date.parse(excelDate(value));
  return Number.isNaN(timestamp) ? Number.MAX_SAFE_INTEGER : timestamp;
}

export function weekOfMonth(value: CellValue) {
  const match = excelDate(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return 0;
  return weekOfMonthFromDate(new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
}

export function chineseWeek(number: number) {
  return ["", "第一周", "第二周", "第三周", "第四周", "第五周"][number] || "";
}

export function inferServiceWeek(counts: Map<number, number>) {
  const ranked = [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0] - b[0]);
  return {
    label: chineseWeek(ranked[0]?.[0] || 0),
    distribution: ranked.map(([week, count]) => `${chineseWeek(week)} ${count}课次`).join("；"),
  };
}
