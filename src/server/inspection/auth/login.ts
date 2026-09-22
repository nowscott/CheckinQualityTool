import { assertSameOrigin, authenticateUser, authCookieHeader, authHintCookieHeader, publicUser, type ApiRequest, type ApiResponse } from "../shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "POST") {
    response.status(405);
    response.setHeader("Allow", "POST");
    response.json({ error: "仅支持 POST。" });
    return;
  }
  try {
    assertSameOrigin(request);
    const body = request.body && typeof request.body === "object" ? request.body as Record<string, unknown> : {};
    const result = await authenticateUser(body.username, body.password);
    if (!result) {
      response.status(401).json({ error: "用户名或密码不正确，或账号暂时被锁定。" });
      return;
    }
    response.setHeader("Set-Cookie", [authCookieHeader(result.token, 8 * 60 * 60), authHintCookieHeader(result.user, 8 * 60 * 60)]);
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json({ user: publicUser(result.user), expiresIn: 8 * 60 * 60 });
  } catch (error) {
    response.status(error instanceof Error && error.message === "请求来源不受信任。" ? 403 : 503).json({ error: error instanceof Error ? error.message : "登录失败，请稍后重试。" });
  }
}
