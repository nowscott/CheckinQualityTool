import { assertSameOrigin, currentAuthUser, parseBody, requireSession, resetUserPassword, type ApiRequest, type ApiResponse, type AuthUser } from "../../shared.js";

export default async function handler(request: ApiRequest & { query?: Record<string, string | string[] | undefined> }, response: ApiResponse) {
  if (request.method !== "POST") {
    response.status(405);
    response.setHeader("Allow", "POST");
    response.json({ error: "仅支持 POST。" });
    return;
  }
  if (!(await requireSession(request, response, "admin"))) return;
  const id = typeof request.query?.id === "string" ? request.query.id : "";
  if (!id) {
    response.status(400).json({ error: "缺少用户 ID。" });
    return;
  }
  try {
    assertSameOrigin(request);
    const actor = await currentAuthUser(request);
    if (!actor) {
      response.status(401).json({ error: "请先登录抽检系统。" });
      return;
    }
    await resetUserPassword(actor as AuthUser, id, parseBody(request).password);
    response.status(200).json({ ok: true });
  } catch (error) {
    response.status(error instanceof Error && error.message === "请求来源不受信任。" ? 403 : 400).json({ error: error instanceof Error ? error.message : "密码重置失败。" });
  }
}
