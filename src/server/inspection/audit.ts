import { auditEntries, requireSession, type ApiRequest, type ApiResponse } from "./shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }
  if (!(await requireSession(request, response, "admin"))) return;
  try {
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json({ entries: await auditEntries() });
  } catch (error) {
    response.status(503).json({ error: error instanceof Error ? error.message : "审计记录读取失败。" });
  }
}
