declare const process: { env: Record<string, string | undefined> };

interface McpResponse {
  error?: { message?: string };
  result?: {
    isError?: boolean;
    content?: Array<{ type?: string; text?: string }>;
  };
}

const MCP_URL = "https://docs.qq.com/openapi/mcp";

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

export function parseTencentCsv(value: string) {
  const rows: string[][] = [];
  let row: string[] = [];
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

export async function readTencentSheetCsv(sheetId: string, endColumn: number) {
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
            sheet_id: sheetId,
            start_row: 0,
            end_row: 999,
            start_col: 0,
            end_col: endColumn,
            return_csv: true,
          },
        },
      }),
      signal: controller.signal,
    });
    if (!response.ok) throw new Error(`腾讯文档接口 HTTP ${response.status}。`);
    return extractCsv((await response.json()) as McpResponse);
  } finally {
    clearTimeout(timeout);
  }
}
