import { authModeInfo, currentAuthUser, hasValidSession, publicUser, type ApiRequest, type ApiResponse } from "../shared.js";

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
      response.setHeader("Cache-Control", "no-store");
      response.status(200).json({ user: publicUser(user), legacy: false, legacyAvailable: false });
      return;
    }
    const mode = authModeInfo();
    if (mode.legacyAvailable && hasValidSession(request)) {
      response.setHeader("Cache-Control", "no-store");
      response.status(200).json({
        user: { id: "legacy", username: "legacy", displayName: "兼容密码", role: "admin", lastLoginAt: null },
        legacy: true,
        legacyAvailable: true,
      });
      return;
    }
    response.setHeader("Cache-Control", "no-store");
    response.status(401).json({ user: null, legacy: false, legacyAvailable: mode.legacyAvailable });
  } catch (error) {
    response.status(503).json({ error: error instanceof Error ? error.message : "登录状态读取失败。" });
  }
}
