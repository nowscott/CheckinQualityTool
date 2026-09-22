import { historyBatches, parseHistoryFilters, requireSession, type ApiRequest, type ApiResponse } from "./shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }
  if (!(await requireSession(request, response, "viewer"))) return;
  try {
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json(await historyBatches(parseHistoryFilters(request, { pageSize: 30, status: "all", batchKind: "all" })));
  } catch (error) {
    response.status(503).json({ error: error instanceof Error ? error.message : "抽检历史读取失败。" });
  }
}
