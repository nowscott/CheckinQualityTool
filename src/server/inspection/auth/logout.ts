import { assertSameOrigin, audit, authCookieHeader, cookieHeader, currentAuthUser, revokeAuthSession, type ApiRequest, type ApiResponse } from "../shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "POST") {
    response.status(405);
    response.setHeader("Allow", "POST");
    response.json({ error: "仅支持 POST。" });
    return;
  }
  try {
    assertSameOrigin(request);
    const user = await currentAuthUser(request);
    await revokeAuthSession(request);
    if (user) await audit("logout", user.id, user.id, {});
    response.setHeader("Set-Cookie", [authCookieHeader("", 0), cookieHeader("", 0)]);
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json({ ok: true, user: user ? { username: user.username } : null });
  } catch (error) {
    response.status(error instanceof Error && error.message === "请求来源不受信任。" ? 403 : 503).json({ error: error instanceof Error ? error.message : "退出登录失败。" });
  }
}
