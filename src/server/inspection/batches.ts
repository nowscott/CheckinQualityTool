import { assertSameOrigin, createBatch, currentAuthUser, parseBody, requireSession, type ApiRequest, type ApiResponse } from "./shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "POST") {
    response.status(405);
    response.setHeader("Allow", "POST");
    response.json({ error: "仅支持 POST。" });
    return;
  }
  if (!(await requireSession(request, response, "operator"))) return;
  try {
    assertSameOrigin(request);
    const actor = await currentAuthUser(request);
    const result = await createBatch(parseBody(request), actor?.id || null);
    response.setHeader("Cache-Control", "no-store");
    response.status(result.conflict ? 409 : result.reused ? 200 : 201).json(result);
  } catch (error) {
    response.status(error instanceof Error && error.message === "请求来源不受信任。" ? 403 : 400).json({ error: error instanceof Error ? error.message : "抽检历史保存失败。" });
  }
}
