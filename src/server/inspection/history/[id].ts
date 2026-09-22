import { historyBatch, parseHistoryFilters, requireSession, type ApiRequest, type ApiResponse } from "../shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }
  if (!(await requireSession(request, response, "viewer"))) return;
  const id = typeof request.query?.id === "string" ? request.query.id : "";
  if (!id) {
    response.status(400).json({ error: "缺少抽检批次 ID。" });
    return;
  }
  try {
    const filters = parseHistoryFilters(request, { pageSize: 100 });
    const batch = await historyBatch(id, filters.page, filters.pageSize);
    if (!batch) {
      response.status(404).json({ error: "未找到抽检批次。" });
      return;
    }
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json({ batch });
  } catch (error) {
    response.status(503).json({ error: error instanceof Error ? error.message : "抽检批次读取失败。" });
  }
}
