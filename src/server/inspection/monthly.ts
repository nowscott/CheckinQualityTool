import { monthlyInspection, requireSession, type ApiRequest, type ApiResponse } from "./shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }
  if (!(await requireSession(request, response, "viewer"))) return;
  const month = typeof request.query?.month === "string" ? request.query.month : "";
  const monthNumber = Number(month.slice(5, 7));
  if (!/^\d{4}-\d{2}$/u.test(month) || monthNumber < 1 || monthNumber > 12) {
    response.status(400).json({ error: "月份格式应为 YYYY-MM。" });
    return;
  }
  try {
    response.setHeader("Cache-Control", "no-store");
    response.status(200).json(await monthlyInspection(month));
  } catch (error) {
    response.status(503).json({ error: error instanceof Error ? error.message : "月度抽检记录读取失败。" });
  }
}
