import auditHandler from "../src/server/inspection/audit.js";
import loginHandler from "../src/server/inspection/auth/login.js";
import logoutHandler from "../src/server/inspection/auth/logout.js";
import meHandler from "../src/server/inspection/auth/me.js";
import registerHandler from "../src/server/inspection/auth/register.js";
import batchesHandler from "../src/server/inspection/batches.js";
import replaceHandler from "../src/server/inspection/batches/[id]/replace.js";
import currentHandler from "../src/server/inspection/current.js";
import historyHandler from "../src/server/inspection/history.js";
import historyDetailHandler from "../src/server/inspection/history/[id].js";
import monthlyHandler from "../src/server/inspection/monthly.js";
import sessionHandler from "../src/server/inspection/session.js";
import teachersHandler from "../src/server/inspection/teachers.js";
import teacherDetailHandler from "../src/server/inspection/teachers/[teacherKey].js";
import teachingServiceHandler from "../src/server/inspection/teaching-service.js";
import usersHandler from "../src/server/inspection/users.js";
import userDetailHandler from "../src/server/inspection/users/[id].js";
import resetPasswordHandler from "../src/server/inspection/users/[id]/reset-password.js";
import type { ApiRequest, ApiResponse } from "../src/server/inspection/shared.js";

type Handler = (request: ApiRequest & { query?: Record<string, string | string[] | undefined> }, response: ApiResponse) => unknown;

function queryValue(request: ApiRequest, key: string) {
  const value = request.query?.[key];
  return Array.isArray(value) ? value[0] || "" : value || "";
}

function withParam(request: ApiRequest, key: string, value: string): ApiRequest & { query: Record<string, string | string[] | undefined> } {
  return { ...request, query: { ...(request.query || {}), [key]: value } };
}

function route(request: ApiRequest): { handler: Handler; request: ApiRequest } | null {
  const rawPath = queryValue(request, "path");
  const parts = rawPath.split("/").filter(Boolean);
  if (!parts.length) return null;
  const [root, second, third] = parts;
  if (root === "auth" && second === "login") return { handler: loginHandler as Handler, request };
  if (root === "auth" && second === "logout") return { handler: logoutHandler as Handler, request };
  if (root === "auth" && second === "me") return { handler: meHandler as Handler, request };
  if (root === "auth" && second === "register") return { handler: registerHandler as Handler, request };
  if (root === "audit") return { handler: auditHandler as Handler, request };
  if (root === "batches" && second && third === "replace") return { handler: replaceHandler as Handler, request: withParam(request, "id", second) };
  if (root === "batches") return { handler: batchesHandler as Handler, request };
  if (root === "current") return { handler: currentHandler as Handler, request };
  if (root === "history" && second) return { handler: historyDetailHandler as Handler, request: withParam(request, "id", second) };
  if (root === "history") return { handler: historyHandler as Handler, request };
  if (root === "monthly") return { handler: monthlyHandler as Handler, request };
  if (root === "session") return { handler: sessionHandler as Handler, request };
  if (root === "teachers" && second) return { handler: teacherDetailHandler as Handler, request: withParam(request, "teacherKey", second) };
  if (root === "teachers") return { handler: teachersHandler as Handler, request };
  if (root === "teaching-service") return { handler: teachingServiceHandler as Handler, request };
  if (root === "users" && second && third === "reset-password") return { handler: resetPasswordHandler as Handler, request: withParam(request, "id", second) };
  if (root === "users" && second) return { handler: userDetailHandler as Handler, request: withParam(request, "id", second) };
  if (root === "users") return { handler: usersHandler as Handler, request };
  return null;
}

export default async function handler(request: ApiRequest, response: ApiResponse) {
  const selected = route(request);
  if (!selected) {
    response.status(404).json({ error: "未找到抽检接口。" });
    return;
  }
  await selected.handler(selected.request, response);
}
