import { authCookieHeader, authHintCookieHeader, currentAuthUser, publicUser, type ApiRequest, type ApiResponse } from "../shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }
  try {
    const user = await currentAuthUser(request);
    if (user) {
      response.setHeader("Set-Cookie", authHintCookieHeader(user, 8 * 60 * 60));
      response.setHeader("Cache-Control", "no-store");
      response.status(200).json({ user: publicUser(user) });
      return;
    }
    response.setHeader("Set-Cookie", [authCookieHeader("", 0), authHintCookieHeader(null, 0)]);
    response.setHeader("Cache-Control", "no-store");
    response.status(401).json({ user: null });
  } catch (error) {
    response.status(503).json({ error: error instanceof Error ? error.message : "登录状态读取失败。" });
  }
}
