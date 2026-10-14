import { readLatestInspectionRoster } from "../src/server/rosterService.js";

declare const process: { env: Record<string, string | undefined> };

type HandlerRequest = { method?: string };
type HandlerResponse = {
  status(code: number): HandlerResponse;
  setHeader(name: string, value: string): void;
  json(body: unknown): void;
};

/** The current app is public. Return only the explicitly approved inspection projection. */
export default async function handler(request: HandlerRequest, response: HandlerResponse) {
  response.setHeader("Cache-Control", "private, no-store");
  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");
    response.status(405).json({ error: "仅支持 GET。" });
    return;
  }
  try {
    response.status(200).json(await readLatestInspectionRoster(process.env.ROSTER_API_TOKEN));
  } catch (error) {
    response.status(502).json({ error: error instanceof Error ? error.message : "最新在职名单读取失败，处理已停止。" });
  }
}
