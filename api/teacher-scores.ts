import { readTencentSheetCsv } from "../src/server/tencentDocs.js";
import { parseTeacherScoreCsv } from "../src/server/teacherScores.js";

type HandlerRequest = { method?: string };
type HandlerResponse = {
  status(code: number): HandlerResponse;
  setHeader(name: string, value: string): void;
  json(body: unknown): void;
};

const SCORE_SHEET_ID = "wlc6il";

export default async function handler(request: HandlerRequest, response: HandlerResponse) {
  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");
    response.status(405).json({ error: "仅支持 GET。" });
    return;
  }

  try {
    const csv = await readTencentSheetCsv(SCORE_SHEET_ID, 3);
    const parsed = parseTeacherScoreCsv(csv);
    const orderedScores = [...new Set(parsed.rows.map((row) => row.score))].sort((left, right) => left - right);
    const priorityRanks = new Map(orderedScores.map((score, index) => [score, index + 1]));
    const rows = parsed.rows.map(({ score, ...row }) => ({
      ...row,
      priorityRank: priorityRanks.get(score) || 0,
    }));
    response.setHeader("Cache-Control", "private, no-store");
    response.setHeader("X-Teacher-Score-Source", "tencent-docs");
    response.status(200).json({ ...parsed, rows });
  } catch (error) {
    const message = error instanceof Error ? error.message : "腾讯文档评分页读取失败。";
    response.setHeader("Cache-Control", "no-store");
    response.status(502).json({ error: message });
  }
}
