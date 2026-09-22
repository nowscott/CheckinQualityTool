import { assertSameOrigin, currentAuthUser, parseBody, replaceBatch, requireSession, type ApiRequest, type ApiResponse } from "../../shared.js";

export default async function handler(request: ApiRequest & { query?: Record<string, string | string[] | undefined> }, response: ApiResponse) {
  if (request.method !== "POST") {
    response.status(405);
    response.setHeader("Allow", "POST");
    response.json({ error: "仅支持 POST。" });
    return;
  }
  if (!(await requireSession(request, response, "operator"))) return;
  const id = typeof request.query?.id === "string" ? request.query.id : "";
  if (!id) {
    response.status(400).json({ error: "缺少抽检批次 ID。" });
    return;
  }
  try {
    assertSameOrigin(request);
    const actor = await currentAuthUser(request);
    const result = await replaceBatch(id, parseBody(request), actor?.id || null);
    response.setHeader("Cache-Control", "no-store");
    response.status(201).json(result);
  } catch (error) {
    response.status(error instanceof Error && error.message === "请求来源不受信任。" ? 403 : 400).json({ error: error instanceof Error ? error.message : "抽检批次替换失败。" });
  }
}
