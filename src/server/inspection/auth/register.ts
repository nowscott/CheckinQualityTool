import { assertSameOrigin, authModeInfo, authCookieHeader, parseBody, publicUser, registerUser, type ApiRequest, type ApiResponse } from "../shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "POST") {
    response.status(405);
    response.setHeader("Allow", "POST");
    response.json({ error: "仅支持 POST。" });
    return;
  }
  if (authModeInfo().mode === "legacy") {
    response.status(404).json({ error: "当前环境未启用用户登录。" });
    return;
  }
  try {
    assertSameOrigin(request);
    const body = parseBody(request);
    const result = await registerUser({
      username: body.username,
      displayName: body.displayName,
      password: body.password,
      confirmation: body.confirmation ?? body.confirmPassword,
    });
    response.setHeader("Set-Cookie", authCookieHeader(result.token, 8 * 60 * 60));
    response.setHeader("Cache-Control", "no-store");
    response.status(201).json({ user: publicUser(result.user), expiresIn: 8 * 60 * 60 });
  } catch (error) {
    response.status(error instanceof Error && error.message === "请求来源不受信任。" ? 403 : 400).json({ error: error instanceof Error ? error.message : "注册失败，请稍后重试。" });
  }
}
