import { readTencentSheetCsv } from "../src/server/tencentDocs.js";
import { parseInspectionCoverage, shanghaiMonth } from "../src/server/inspectionCoverage.js";

type Request = { method?: string; query?: { month?: string | string[] } };
type Response = { status(code: number): Response; setHeader(name: string, value: string): void; json(body: unknown): void };
export default async function handler(request: Request, response: Response) {
  response.setHeader("Cache-Control", "no-store");
  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");
    response.status(405).json({ error: "仅支持 GET。" });
    return;
  }
  const month = request.query?.month ?? shanghaiMonth();
  if (typeof month !== "string" || !/^\d{4}-(0[1-9]|1[0-2])$/u.test(month)) {
    response.status(400).json({ error: "抽检月份格式应为 YYYY-MM。" });
    return;
  }
  try {
    const result = parseInspectionCoverage(await readTencentSheetCsv("ih60c3", 5, 9999), month);
    response.status(200).json({ ...result, sourceUrl: "https://docs.qq.com/sheet/DUGFUT3dOckJVeXFm?tab=ih60c3" });
  } catch (error) {
    response.status(502).json({ error: error instanceof Error ? error.message : "抽检登记读取失败。" });
  }
}
