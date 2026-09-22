import { parseHistoryFilters, requireSession, teacherSummaries, type ApiRequest, type ApiResponse } from "./shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }
  if (!(await requireSession(request, response, "viewer"))) return;
  try {
    const filters = parseHistoryFilters(request, {
      pageSize: 30,
      batchKind: "formal",
      status: "active",
      sort: "courses",
    });
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json({ ...await teacherSummaries(filters), page: filters.page, pageSize: filters.pageSize });
  } catch (error) {
    response.status(503).json({ error: error instanceof Error ? error.message : "教师抽检统计读取失败。" });
  }
}
