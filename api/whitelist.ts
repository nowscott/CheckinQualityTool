import { parseTencentCsv, readTencentSheetCsv } from "../src/server/tencentDocs.js";

declare const process: { env: Record<string, string | undefined> };

type CellRow = string[];
type HandlerRequest = { method?: string };
type HandlerResponse = {
  status(code: number): HandlerResponse;
  setHeader(name: string, value: string): void;
  json(body: unknown): void;
  send(body: string): void;
};

const EXPECTED_COLUMNS = ["学员号", "学员姓名", "处理方式", "匹配别名", "说明"];
const ALLOWED_MODES = new Set(["免检", "别名", "保留原名", "剔除"]);

function encodeCsvField(value: string) {
  return /[",\r\n]/u.test(value) ? `"${value.replace(/"/gu, '""')}"` : value;
}

function validateCsv(csv: string) {
  const rows = parseTencentCsv(csv).filter((row) => row.some((value) => value.trim()));
  if (!rows.length || rows[0].join(",") !== EXPECTED_COLUMNS.join(",")) {
    throw new Error("腾讯文档表头不符合要求，请保持五列：学员号、学员姓名、处理方式、匹配别名、说明。");
  }

  const studentIds = new Set<string>();
  for (const [offset, row] of rows.slice(1).entries()) {
    if (row.length !== EXPECTED_COLUMNS.length) {
      throw new Error(`腾讯文档第 ${offset + 2} 行列数不正确。`);
    }
    const studentId = row[0].trim();
    const studentName = row[1].trim();
    const mode = row[2].trim();
    if (!studentId && !studentName) throw new Error(`腾讯文档第 ${offset + 2} 行缺少学员号和学员姓名。`);
    if (!ALLOWED_MODES.has(mode)) throw new Error(`腾讯文档第 ${offset + 2} 行处理方式无效。`);
    if (studentId) {
      if (studentIds.has(studentId)) throw new Error(`腾讯文档存在重复学员号：${studentId}。`);
      studentIds.add(studentId);
    }
  }
  return {
    rowCount: rows.length - 1,
    csv: rows.map((row) => row.map(encodeCsvField).join(",")).join("\n"),
  };
}

export default async function handler(request: HandlerRequest, response: HandlerResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }

  try {
    const csv = await readTencentSheetCsv(process.env.TENCENT_DOC_SHEET_ID || "BB08J2", EXPECTED_COLUMNS.length - 1);
    const { csv: normalizedCsv, rowCount } = validateCsv(csv);
    response.setHeader("Cache-Control", "no-store");
    response.setHeader("Content-Type", "text/csv; charset=utf-8");
    response.setHeader("X-Whitelist-Source", "tencent-docs");
    response.setHeader("X-Whitelist-Row-Count", String(rowCount));
    response.status(200).send(normalizedCsv);
  } catch (error) {
    const message = error instanceof Error ? error.message : "腾讯文档白名单读取失败。";
    response.status(502).json({ error: message });
  }
}
