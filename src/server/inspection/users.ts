import { assertSameOrigin, createUser, currentAuthUser, listUsers, parseBody, requireSession, type ApiRequest, type ApiResponse, type AuthUser } from "./shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "GET" && request.method !== "POST") {
    response.status(405);
    response.setHeader("Allow", "GET, POST");
    response.json({ error: "仅支持 GET 或 POST。" });
    return;
  }
  if (!(await requireSession(request, response, "admin"))) return;
  try {
    if (request.method === "GET") {
      response.setHeader("Cache-Control", "no-store");
      response.status(200).json({ users: await listUsers() });
      return;
    }
    assertSameOrigin(request);
    const actor = await currentAuthUser(request);
    if (!actor) {
      response.status(401).json({ error: "请先登录抽检系统。" });
      return;
    }
    const body = parseBody(request);
    const user = await createUser(actor as AuthUser, body);
    response.setHeader("Cache-Control", "no-store");
    response.status(201).json({ user });
  } catch (error) {
    response.status(error instanceof Error && error.message === "请求来源不受信任。" ? 403 : 400).json({ error: error instanceof Error ? error.message : "用户操作失败。" });
  }
}
