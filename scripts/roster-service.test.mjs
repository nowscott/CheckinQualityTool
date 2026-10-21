import assert from "node:assert/strict";
import test from "node:test";
import fs from "node:fs";
import vm from "node:vm";

const { projectInspectionRoster, readLatestInspectionRoster } = await import("../src/server/rosterService.js");
const { loadLatestInspectionRoster, rosterFromSnapshot } = await import("../src/lib/rosterSnapshot.js");
const { default: handler } = await import("../api/roster.js");
globalThis.self = { postMessage() {} };
const { processInspection } = await import("../src/worker/inspectionProcess.js");
const sandbox = { console, ArrayBuffer, Uint8Array, Date };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync("public/vendor/xlsx.full.min.js", "utf8"), sandbox);
globalThis.XLSX = sandbox.XLSX;
const fields = ["姓名", "员工ID", "师训组长", "师训主管/助理主管", "科目师训负责人", "教研组", "用工类别", "入职日期", "司龄（年）", "教师状态", "校区", "区域", "非课状态", "可带年级", "岗位短描述", "性别", "教师编码", "教师评级", "邮箱"];
const row = (id, email, role) => ({ ...Object.fromEntries([...fields].reverse().map((field) => [field, ""])), 姓名: `教师${id}`, 员工ID: id, 邮箱: email, 岗位短描述: role, personKey: `id:${id}` });
const latest = {
  snapshotDate: "2025-02-03", sourceFile: "在职教师明细20250203.xlsx", sourceSha256: "a".repeat(64), rowCount: 5,
  rows: [row("1", "TEACHER@XDF.CN", "高级教师"), row("2", "manager@xdf.cn", "经理"), row("3", "leader@xdf.cn", "高级组长"), row("4", "supervisor@xdf.cn", "助理主管"), row("5", "invalid", "教师")],
};
const snapshot = projectInspectionRoster(latest);
const response = (body, status = 200) => ({ ok: status === 200, status, json: async () => structuredClone(body) });

test("按字段名称投影，沿用邮箱、经理与主管规则，公开数据不含完整行", () => {
  assert.equal(snapshot.rowCount, latest.rows.length);
  assert.equal(snapshot.matchedEmailRows, 4);
  assert.deepEqual(snapshot.roleExcludedEmails, ["manager@xdf.cn"]);
  assert.deepEqual(snapshot.unsubmittedExtraExcludedEmails, ["supervisor@xdf.cn"]);
  assert.equal(snapshot.emails.includes("leader@xdf.cn"), true);
  assert.deepEqual(Object.keys(snapshot).sort(), ["snapshotDate", "sourceFile", "sourceSha256", "rowCount", "matchedEmailRows", "emails", "roleExcludedEmails", "unsubmittedExtraExcludedEmails"].sort());
});

test("人数、缺失字段、重复人员键和版本异常均停止", () => {
  const missing = structuredClone(latest); delete missing.rows[0]["教研组"];
  const duplicate = structuredClone(latest); duplicate.rows[1].personKey = duplicate.rows[0].personKey;
  for (const value of [missing, duplicate, { ...latest, rowCount: 6 }, { ...latest, snapshotDate: "2025-02-31" }, { ...latest, sourceSha256: "bad" }]) {
    assert.throws(() => projectInspectionRoster(value));
  }
});

test("每次读取最新版本，密钥只放请求认证头，不放响应", async () => {
  const calls = [];
  const fetcher = async (url, options) => { calls.push({ url, options }); return response(latest); };
  await readLatestInspectionRoster("test-secret", fetcher);
  const result = await readLatestInspectionRoster("test-secret", fetcher);
  assert.equal(calls.length, 4);
  for (const call of calls) {
    assert.equal(call.options.headers.Authorization, "Bearer test-secret");
    assert.equal(call.options.cache, "no-store");
    assert.equal(call.options.redirect, "error");
  }
  assert.equal(JSON.stringify(result).includes("test-secret"), false);
});

test("版本切换时重新取完整一对，持续不一致则停止", async () => {
  const next = { ...latest, sourceSha256: "b".repeat(64) };
  let calls = 0;
  const fetcher = async () => response(calls++ === 0 ? latest : next);
  assert.equal((await readLatestInspectionRoster("test", fetcher)).sourceSha256, next.sourceSha256);
  assert.equal(calls, 4);
  await assert.rejects(readLatestInspectionRoster("test", async (url) => response(url.endsWith("meta") ? next : latest)), /版本正在更新/u);
});

test("错误密钥、上游故障、超时与异常响应不泄漏服务端详情，也不回退旧名单", async () => {
  for (const status of [401, 403, 500]) {
    await assert.rejects(readLatestInspectionRoster("test-secret", async () => response({ error: "test-secret" }, status)), (error) => !error.message.includes("test-secret"));
  }
  await assert.rejects(readLatestInspectionRoster("test-secret", async () => { throw new Error("test-secret"); }), /接口读取失败/u);
  await assert.rejects(readLatestInspectionRoster(undefined), /服务端尚未配置/u);
  await assert.rejects(readLatestInspectionRoster("test", async () => ({ ok: true, json: async () => { throw new Error("test-secret"); } })), /接口读取失败/u);
});

test("无登录项目匿名访问仅获得批准的邮箱岗位投影；非 GET 拒绝", async (context) => {
  context.mock.method(globalThis, "fetch", async () => response(latest));
  const previous = process.env.ROSTER_API_TOKEN;
  process.env.ROSTER_API_TOKEN = "test-secret";
  context.after(() => { if (previous === undefined) delete process.env.ROSTER_API_TOKEN; else process.env.ROSTER_API_TOKEN = previous; });
  const res = { headers: {}, status(code) { this.code = code; return this; }, setHeader(name, value) { this.headers[name] = value; }, json(body) { this.body = body; } };
  await handler({ method: "GET" }, res);
  assert.equal(res.code, 200);
  assert.deepEqual(res.body, snapshot);
  assert.equal(res.headers["Cache-Control"], "private, no-store");
  await handler({ method: "POST" }, res);
  assert.equal(res.code, 405);
});

test("页面每次开始重新读取；格式错误和请求失败都停止", async (context) => {
  let count = 0;
  context.mock.method(globalThis, "fetch", async (url, options) => {
    count += 1; assert.equal(url, "/api/roster"); assert.equal(options.cache, "no-store"); return response(snapshot);
  });
  await loadLatestInspectionRoster(); await loadLatestInspectionRoster(); assert.equal(count, 2);
  context.mock.method(globalThis, "fetch", async () => response({ error: "接口不可用" }, 502));
  await assert.rejects(loadLatestInspectionRoster(), /接口不可用/u);
  assert.throws(() => rosterFromSnapshot({ ...snapshot, roleExcludedEmails: ["other@xdf.cn"] }));
});

function fileFromRows(name, rows) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "明细");
  return new File([XLSX.write(wb, { type: "array", bookType: "xlsx" })], name);
}
const feedbackFile = () => fileFromRows("课堂反馈.xlsx", [
  ["老师姓名", "学员姓名", "学员号", "课次ID", "课次开始时间", "课次结束时间", "老师邮箱", "是否生成报告"],
  ["教师甲", "学员甲", "1", "COURSE-1", "2025-02-04 10:00", "2025-02-04 12:00", "teacher@xdf.cn", "是"],
  ["经理甲", "学员乙", "2", "COURSE-2", "2025-02-04 10:00", "2025-02-04 12:00", "manager@xdf.cn", "否"],
]);
const request = () => ({ type: "process", mode: "inspection", feedbackFile: feedbackFile(), sampleCount: 1, attempt: 1, includeExplanation: true, includeReason: false, priorityMode: "coverage", focusTeacherNames: [], teacherScoreRows: [] });

test("Worker 固定本次在职快照，独立读取岗位纠偏；导出记录日期与哈希", async (context) => {
  context.mock.method(globalThis, "fetch", async (url) => {
    assert.equal(url, "/api/role-corrections");
    return response({ managerEmails: [], supervisorEmails: [], teacherEmails: [] });
  });
  const messages = [];
  await processInspection({ ...request(), defaultRoster: snapshot }, { postMessage: (m) => messages.push(m) });
  const result = messages.at(-1);
  assert.equal(result.rosterSnapshot.sourceSha256, snapshot.sourceSha256);
  assert.equal(result.rosterSnapshot.snapshotDate, snapshot.snapshotDate);
  assert.equal(result.summary.excludedManagementRows, 1);
  assert.equal(result.summary.selectedRows, 1);
  const wb = XLSX.read(Buffer.concat(result.chunks.map((chunk) => Buffer.from(chunk))), { type: "buffer" });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets["处理说明"], { header: 1 });
  assert.equal(rows.some((r) => r[0] === "在职源文件 SHA-256" && r[1] === snapshot.sourceSha256), true);
});

test("手动上传指定历史版本独立于接口，覆盖默认名单并记录文件哈希", async (context) => {
  context.mock.method(globalThis, "fetch", async (url) => {
    assert.equal(url, "/api/role-corrections");
    return response({ managerEmails: [], supervisorEmails: [], teacherEmails: [] });
  });
  const rosterFile = fileFromRows("在职教师明细20240203.xlsx", [["岗位短描述", "邮箱"], ["教师", "teacher@xdf.cn"]]);
  const messages = [];
  await processInspection({ ...request(), rosterFile }, { postMessage: (m) => messages.push(m) });
  const result = messages.at(-1);
  assert.equal(result.rosterSnapshot.snapshotDate, "2024-02-03");
  assert.match(result.rosterSnapshot.sourceSha256, /^[a-f0-9]{64}$/u);
  await assert.rejects(processInspection(request(), { postMessage() {} }), /没有取得最新在职名单/u);
});
