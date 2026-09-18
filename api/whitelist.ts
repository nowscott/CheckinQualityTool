declare const process: { env: Record<string, string | undefined> };

type CellRow = string[];

type HandlerRequest = {
  method?: string;
};

type HandlerResponse = {
  status(code: number): HandlerResponse;
  setHeader(name: string, value: string): void;
  json(body: unknown): void;
  send(body: string): void;
};

type McpResponse = {
  error?: { message?: string };
  result?: {
    isError?: boolean;
    content?: Array<{ type?: string; text?: string }>;
  };
};

const MCP_URL = "https://docs.qq.com/openapi/mcp";
const EXPECTED_COLUMNS = ["学员号", "学员姓名", "处理方式", "匹配别名", "说明"];
const ALLOWED_MODES = new Set(["免检", "别名", "保留原名", "剔除"]);

function parseCsv(value: string): CellRow[] {
  const rows: CellRow[] = [];
  let row: CellRow = [];
  let field = "";
  let quoted = false;
  const source = String(value || "").replace(/^\uFEFF/u, "");

  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    if (quoted) {
      if (character === '"' && source[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        field += character;
      }
    } else if (character === '"') {
      quoted = true;
    } else if (character === ",") {
      row.push(field);
      field = "";
    } else if (character === "\n") {
      row.push(field.replace(/\r$/u, ""));
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += character;
    }
  }
  if (field || row.length) {
    row.push(field.replace(/\r$/u, ""));
    rows.push(row);
  }
  return rows;
}

function extractCsv(payload: McpResponse) {
  if (payload.error) throw new Error(payload.error.message || "腾讯文档接口返回错误。");
  if (payload.result?.isError) throw new Error("腾讯文档读取失败。");
  const text = payload.result?.content?.find((item) => item.type === "text")?.text;
  if (!text) throw new Error("腾讯文档没有返回表格内容。");

  try {
    const parsed = JSON.parse(text) as { csv_data?: string; data?: string };
    return parsed.csv_data ?? parsed.data ?? text;
  } catch {
    return text;
  }
}

function encodeCsvField(value: string) {
  return /[",\r\n]/u.test(value) ? `"${value.replace(/"/gu, '""')}"` : value;
}

function validateCsv(csv: string) {
  const rows = parseCsv(csv).filter((row) => row.some((value) => value.trim()));
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

async function readWhitelist() {
  const token = process.env.TENCENT_DOCS_TOKEN;
  if (!token) throw new Error("服务器未配置腾讯文档读取凭证。");

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 10_000);
  try {
    const response = await fetch(MCP_URL, {
      method: "POST",
      headers: {
        Authorization: token,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        jsonrpc: "2.0",
        id: Date.now(),
        method: "tools/call",
        params: {
          name: "sheet.get_cell_data",
          arguments: {
            file_id: process.env.TENCENT_DOC_FILE_ID || "DUGFUT3dOckJVeXFm",
            sheet_id: process.env.TENCENT_DOC_SHEET_ID || "BB08J2",
            start_row: 0,
            end_row: 999,
            start_col: 0,
            end_col: 4,
            return_csv: true,
          },
        },
      }),
      signal: controller.signal,
    });
    if (!response.ok) throw new Error(`腾讯文档接口 HTTP ${response.status}。`);
    const csv = extractCsv((await response.json()) as McpResponse);
    return validateCsv(csv);
  } finally {
    clearTimeout(timeout);
  }
}

export default async function handler(request: HandlerRequest, response: HandlerResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }

  try {
    const { csv, rowCount } = await readWhitelist();
    response.setHeader("Cache-Control", "no-store");
    response.setHeader("Content-Type", "text/csv; charset=utf-8");
    response.setHeader("X-Whitelist-Source", "tencent-docs");
    response.setHeader("X-Whitelist-Row-Count", String(rowCount));
    response.status(200).send(csv);
  } catch (error) {
    const message = error instanceof Error ? error.message : "腾讯文档白名单读取失败。";
    response.status(502).json({ error: message });
  }
}
