import { progress } from "./progress";
import type { CellValue } from "./types";
import { text } from "./utils";

export interface FoundSheet {
  name: string;
  rows: CellValue[][];
}

export function headerMap(headerRow: CellValue[]) {
  const map = new Map<string, number[]>();
  headerRow.forEach((value, index) => {
    const key = text(value);
    if (!map.has(key)) map.set(key, []);
    map.get(key)!.push(index);
  });
  return map;
}

export function findSheet(workbook: SheetJsWorkbook, requiredHeaders: string[]): FoundSheet {
  for (const name of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
      header: 1,
      range: 0,
      blankrows: false,
      defval: "",
    });
    if (!rows.length) continue;
    const headers = new Set(rows[0].map(text));
    if (requiredHeaders.every((header) => headers.has(header))) return { name, rows };
  }
  throw new Error(`找不到包含字段“${requiredHeaders.join("、")}”的工作表。`);
}

function isRecoverableZipSizeWarning(args: unknown[]) {
  return args.some((arg) => /^Bad uncompressed size: \d+ != 0$/u.test(text(arg)));
}

function readXlsx(data: ArrayBuffer) {
  const originalWarn = console.warn;
  const originalError = console.error;
  console.warn = (...args: unknown[]) => {
    if (!isRecoverableZipSizeWarning(args)) originalWarn(...args);
  };
  console.error = (...args: unknown[]) => {
    if (!isRecoverableZipSizeWarning(args)) originalError(...args);
  };
  try {
    return XLSX.read(data, { type: "array", dense: true, cellDates: true });
  } finally {
    console.warn = originalWarn;
    console.error = originalError;
  }
}

export async function readWorkbook(
  file: File,
  stageStart: number,
  stageEnd: number,
  label: string,
) {
  progress(
    `正在读取${label}`,
    `${file.name} · ${(file.size / 1024 / 1024).toFixed(1)} MB`,
    stageStart,
  );
  const data = await file.arrayBuffer();
  progress(
    `正在解析${label}`,
    "使用 dense 模式解析 Excel，此步骤耗时取决于文件大小。",
    stageStart + (stageEnd - stageStart) * 0.35,
  );
  const workbook = readXlsx(data);
  progress(`${label}解析完成`, "正在提取必要字段并释放原始工作簿。", stageEnd);
  return workbook;
}
