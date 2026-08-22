import { progress } from "./progress";
import type { CellValue } from "./types";
import { text } from "./utils";

export interface FoundSheet {
  name: string;
  rows: CellValue[][];
}

export interface SheetCandidate extends FoundSheet {
  headers: string[];
  map: Map<string, number[]>;
  order: number;
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

export function sheetCandidates(workbook: SheetJsWorkbook): SheetCandidate[] {
  const candidates: SheetCandidate[] = [];
  workbook.SheetNames.forEach((name, order) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
      header: 1,
      range: 0,
      blankrows: false,
      defval: "",
    });
    if (!rows.length) return;
    candidates.push({ name, rows, headers: rows[0].map(text), map: headerMap(rows[0]), order });
  });
  return candidates;
}

export function findSheet(workbook: SheetJsWorkbook, requiredHeaders: string[]): FoundSheet {
  for (const candidate of sheetCandidates(workbook)) {
    if (requiredHeaders.every((header) => candidate.map.has(header))) return candidate;
  }
  throw new Error(`找不到包含字段“${requiredHeaders.join("、")}”的工作表。`);
}

function isRecoverableZipSizeWarning(args: unknown[]) {
  return args.length === 1 && /^Bad uncompressed size: \d+ != 0$/u.test(text(args[0]));
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
    return XLSX.read(data, { type: "array", dense: true, cellDates: true, cellText: false });
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
