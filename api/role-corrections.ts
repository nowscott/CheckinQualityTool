import { readTencentSheetCsv } from "../src/server/tencentDocs.js";
import { parseInspectionRoleCorrections } from "../src/server/inspectionRoleCorrections.js";
import { ROLE_CORRECTIONS_SHEET_ID } from "../src/lib/inspectionRoleCorrections.js";

type HandlerRequest = { method?: string };
type HandlerResponse = {
  status(code: number): HandlerResponse;
  setHeader(name: string, value: string): void;
  json(body: unknown): void;
};

export default async function handler(request: HandlerRequest, response: HandlerResponse) {
  response.setHeader("Cache-Control", "private, no-store");
  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");
    response.status(405).json({ error: "仅支持 GET。" });
    return;
  }
  try {
    const csv = await readTencentSheetCsv(ROLE_CORRECTIONS_SHEET_ID, 4);
    response.status(200).json(parseInspectionRoleCorrections(csv));
  } catch (error) {
    const message = error instanceof Error && error.message.startsWith("岗位纠偏")
      ? error.message : "岗位纠偏表读取失败，请稍后重试。";
    response.status(502).json({ error: message });
  }
}
