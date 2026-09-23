import { parseTencentCsv, readTencentSheetCsv } from "../src/server/tencentDocs.js";

type HandlerRequest = { method?: string };
type HandlerResponse = {
  status(code: number): HandlerResponse;
  setHeader(name: string, value: string): void;
  json(body: unknown): void;
};

const SHEET_ID = "8gfbuz";
const HEADERS = ["教师姓名", "教研组", "师训组长", "教学服务周", "是否反馈"] as const;
const HEADER_ALIASES: Record<(typeof HEADERS)[number], readonly string[]> = {
  教师姓名: ["教师姓名"],
  教研组: ["教研组"],
  师训组长: ["师训组长"],
  教学服务周: ["教学服务周", "教学服务周期"],
  是否反馈: ["是否反馈"],
};

function normalize(value: string) {
  return value.normalize("NFKC").replace(/[\p{White_Space}\p{Cf}]/gu, "").trim();
}

function isUnreported(value: string) {
  const status = normalize(value);
  return status === "否" || status.includes("未反馈");
}

export default async function handler(request: HandlerRequest, response: HandlerResponse) {
  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");
    response.status(405).json({ error: "仅支持 GET。" });
    return;
  }

  try {
    const csv = await readTencentSheetCsv(SHEET_ID, HEADERS.length - 1);
    const rows = parseTencentCsv(csv).filter((row) => row.some((value) => value.trim()));
    const normalizedRows = rows.map((row) => row.map(normalize));
    const headerRowIndex = normalizedRows.findIndex((candidateHeaders) =>
      HEADERS.every((header) => HEADER_ALIASES[header].some((alias) => candidateHeaders.includes(normalize(alias)))),
    );
    if (headerRowIndex < 0) {
      throw new Error("重点关注页表头不完整，需要教师姓名、教研组、师训组长、教学服务周（或教学服务周期）、是否反馈五列。");
    }
    const actualHeaders = normalizedRows[headerRowIndex];
    const indexes = HEADERS.map((header) =>
      HEADER_ALIASES[header]
        .map((alias) => actualHeaders.indexOf(normalize(alias)))
        .find((index) => index >= 0) ?? -1,
    );
    const dataRows = rows.slice(headerRowIndex + 1);

    const unreportedNames = new Set<string>();
    let unreportedRowCount = 0;
    const missingNameRows = dataRows.filter((row) =>
      row.some((value) => value.trim()) && !row[indexes[0]]?.trim(),
    ).length;
    for (const row of dataRows) {
      const name = row[indexes[0]]?.trim() || "";
      const status = row[indexes[4]] || "";
      if (!name || !isUnreported(status)) continue;
      unreportedRowCount += 1;
      unreportedNames.add(name);
    }

    response.setHeader("Cache-Control", "no-store");
    response.setHeader("X-Feedback-Focus-Source", "tencent-docs");
    response.status(200).json({
      rowCount: dataRows.length,
      unreportedRowCount,
      unreportedTeacherNames: [...unreportedNames],
      missingNameRows,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "重点关注表读取失败。";
    response.setHeader("Cache-Control", "no-store");
    response.status(502).json({ error: message });
  }
}
