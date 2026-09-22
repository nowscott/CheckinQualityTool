import scoreSnapshot from "../../../data/inspection/teaching-service-q1.json" with { type: "json" };
import { requireSession, type ApiRequest, type ApiResponse } from "./shared.js";

export default async function handler(request: ApiRequest, response: ApiResponse) {
  if (request.method !== "GET") {
    response.status(405);
    response.setHeader("Allow", "GET");
    response.json({ error: "仅支持 GET。" });
    return;
  }
  if (!(await requireSession(request, response, "viewer"))) return;
  response.setHeader("Cache-Control", "no-store");
  response.status(200).json(scoreSnapshot);
}
