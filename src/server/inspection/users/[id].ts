import { assertSameOrigin, currentAuthUser, parseBody, requireSession, type ApiRequest, type ApiResponse, type AuthUser, updateUser } from "../shared.js";

export default async function handler(request: ApiRequest & { query?: Record<string, string | string[] | undefined> }, response: ApiResponse) {
  if (request.method !== "PATCH") {
    response.status(405);
    response.setHeader("Allow", "PATCH");
    response.json({ error: "仅支持 PATCH。" });
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
    const user = await updateUser(actor as AuthUser, id, parseBody(request));
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json({ user });
  } catch (error) {
    response.status(error instanceof Error && error.message === "请求来源不受信任。" ? 403 : 400).json({ error: error instanceof Error ? error.message : "用户更新失败。" });
  }
}
