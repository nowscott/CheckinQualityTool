import { currentBatch, requireSession, type ApiRequest, type ApiResponse } from "./shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }
  if (!(await requireSession(request, response, "viewer"))) return;
  try {
    const weekStart = typeof request.query?.weekStart === "string" ? request.query.weekStart : undefined;
    const batchKind = request.query?.batchKind === "trial" ? "trial" : "formal";
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json({ batch: await currentBatch(weekStart, batchKind) });
  } catch (error) {
    response.status(503).json({ error: error instanceof Error ? error.message : "抽检历史读取失败。" });
  }
}
