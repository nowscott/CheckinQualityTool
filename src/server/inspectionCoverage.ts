import { parseTencentCsv } from "./tencentDocs.js";

export function shanghaiMonth(now = new Date()) {
  const parts = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit" }).formatToParts(now);
  return `${parts.find((p) => p.type === "year")!.value}-${parts.find((p) => p.type === "month")!.value}`;
}

export function validateCoverageMonth(month: string) {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/u.test(month)) throw new Error("抽检月份格式应为 YYYY-MM。");
}

/** Accept typed dates and spreadsheet date serials, without timezone-dependent parsing. */
export function inspectionDate(value: string) {
  const text = value.normalize("NFKC").trim();
  if (/^\d{5}(?:\.\d+)?$/u.test(text)) {
    const serial = Number(text);
    if (serial < 40000 || serial > 100000) return "";
    return new Date(Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000).toISOString().slice(0, 10);
  }
  const match = /^(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})(?:日)?(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$/u.exec(text);
  if (!match) return "";
  const date = `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
  const parsed = new Date(`${date}T00:00:00Z`);
  return Number.isFinite(parsed.valueOf()) && parsed.toISOString().slice(0, 10) === date ? date : "";
}

const HEADERS = ["教师姓名", "学员姓名", "课次日期", "课次时间"];
export function parseInspectionCoverage(csv: string, month: string) {
  validateCoverageMonth(month);
  const rows = parseTencentCsv(csv);
  // Tencent returns the full requested rectangle, including trailing empty rows.
  while (rows.length && rows[rows.length - 1].every((cell) => !cell.trim())) rows.pop();
  if (rows.length >= 10000) throw new Error("抽检登记超过读取范围，请扩展读取范围后重试。");
  const headers = rows[0]?.map((v) => v.trim()) || [];
  const indexes = HEADERS.map((name) => headers.indexOf(name));
  const dateIndex = headers.indexOf("抽检日期");
  if (indexes.some((i) => i < 0) || dateIndex < 0) throw new Error("抽检登记需要教师姓名、学员姓名、课次日期、课次时间及抽检日期；抽检日期放在第六列。");
  const teachers = new Map<string, { teacherEmail: string; teacherName: string }>();
  let rowCount = 0;
  let undatedRowCount = 0;
  let invalidDateRowCount = 0;
  const invalidDateRows: number[] = [];
  for (const [i, row] of rows.slice(1).entries()) {
    const cells = indexes.map((index) => (row[index] || "").trim());
    if (!cells.some(Boolean)) continue;
    // Teachers can paste multiple batches, including the repeated export headers.
    if (cells.every((cell, index) => cell === HEADERS[index])) continue;
    if (cells.some((v) => !v)) throw new Error(`抽检登记第 ${i + 2} 行缺少信息，请从抽检文件完整粘贴前四列。`);
    const rawDate = (row[dateIndex] || "").trim();
    if (!rawDate) { undatedRowCount++; continue; }
    const date = inspectionDate(rawDate);
    if (!date) {
      invalidDateRowCount++;
      if (invalidDateRows.length < 20) invalidDateRows.push(i + 2);
      continue;
    }
    if (date.slice(0, 7) !== month) continue;
    const teacherName = cells[0].normalize("NFKC").replace(/[\p{White_Space}\p{Cf}]/gu, "");
    if (!teacherName) throw new Error(`抽检登记第 ${i + 2} 行缺少教师姓名。`);
    rowCount++;
    // Coverage uses the manually entered inspection date, never the lesson date.
    teachers.set(teacherName.toLowerCase(), { teacherEmail: "", teacherName });
  }
  return { month, rowCount, undatedRowCount, invalidDateRowCount, invalidDateRows, teachers: [...teachers.values()] };
}
