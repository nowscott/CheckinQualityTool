import { assertSameOrigin, authModeInfo, cookieHeader, env, issueSession, parseBody, type ApiRequest, type ApiResponse } from "./shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "POST") {
    response.status(405);
    response.setHeader("Allow", "POST");
    response.json({ error: "仅支持 POST。" });
    return;
  }
  if (authModeInfo().mode === "login") {
    response.status(404).json({ error: "当前环境已切换为用户登录。" });
    return;
  }
  try {
    assertSameOrigin(request);
  } catch (error) {
    response.status(403).json({ error: error instanceof Error ? error.message : "请求来源不受信任。" });
    return;
  }
  const password = String(parseBody(request).password || "");
  const expected = env("INSPECTION_PASSWORD");
  if (!expected) {
    response.status(503).json({ error: "服务器尚未配置抽检历史密码。" });
    return;
  }
  if (!password || password !== expected) {
    response.status(401).json({ error: "抽检历史密码不正确。" });
    return;
  }
  response.setHeader("Set-Cookie", cookieHeader(issueSession(), 8 * 60 * 60));
  response.setHeader("Cache-Control", "no-store");
  response.status(200).json({ ok: true, expiresIn: 8 * 60 * 60 });
}
