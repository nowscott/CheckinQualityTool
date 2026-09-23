import { assertSameOrigin, currentAuthUser, requireSession, type ApiRequest, type ApiResponse, voidInspectionBatch } from "../../shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "POST") {
    response.status(405);
    response.setHeader("Allow", "POST");
    response.json({ error: "仅支持 POST。" });
    return;
  }
  if (!(await requireSession(request, response, "admin"))) return;
  const id = typeof request.query?.id === "string" ? request.query.id : "";
  if (!id) {
    response.status(400).json({ error: "缺少抽检批次 ID。" });
    return;
  }
  try {
    assertSameOrigin(request);
    const actor = await currentAuthUser(request);
    if (!actor) {
      response.status(401).json({ error: "请先登录抽检系统。" });
      return;
    }
    if (actor.role !== "admin") {
      response.status(403).json({ error: "当前账号没有执行此操作的权限。" });
      return;
    }
    const batch = await voidInspectionBatch(id, actor.id);
    if (!batch) {
      response.status(404).json({ error: "未找到抽检批次。" });
      return;
    }
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json({ batch });
  } catch (error) {
    const message = error instanceof Error ? error.message : "批次作废失败。";
    const status = message === "请求来源不受信任。" ? 403 : message === "当前批次已不是生效状态，请刷新后重试。" ? 409 : 503;
    response.status(status).json({ error: status === 503 ? "批次作废失败，请稍后重试。" : message });
  }
}
