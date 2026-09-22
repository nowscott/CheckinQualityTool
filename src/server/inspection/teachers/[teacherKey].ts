import { parseHistoryFilters, requireSession, teacherDetail, type ApiRequest, type ApiResponse } from "../shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }
  if (!(await requireSession(request, response, "viewer"))) return;
  const rawKey = typeof request.query?.teacherKey === "string" ? request.query.teacherKey : "";
  if (!rawKey) {
    response.status(400).json({ error: "缺少教师标识。" });
    return;
  }
  try {
    const filters = parseHistoryFilters(request, {
      pageSize: 50,
      batchKind: "formal",
      status: "active",
      sort: "recent",
    });
    const detail = await teacherDetail(decodeURIComponent(rawKey), filters);
    if (!detail.teacher) {
      response.status(404).json({ error: "未找到该教师的抽检记录。" });
      return;
    }
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json(detail);
  } catch (error) {
    response.status(503).json({ error: error instanceof Error ? error.message : "教师抽检明细读取失败。" });
  }
}
