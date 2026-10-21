import assert from "node:assert/strict";
import test from "node:test";
import fs from "node:fs";
import vm from "node:vm";

const { parseInspectionRoleCorrections } = await import("../src/server/inspectionRoleCorrections.js");
const { applyInspectionRoleCorrections, loadInspectionRoleCorrections } = await import("../src/lib/inspectionRoleCorrections.js");
const { default: handler } = await import("../api/role-corrections.js");
globalThis.self = { postMessage() {} };
const { processInspection } = await import("../src/worker/inspectionProcess.js");
const sandbox = { console, ArrayBuffer, Uint8Array, Date };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync("public/vendor/xlsx.full.min.js", "utf8"), sandbox);
globalThis.XLSX = sandbox.XLSX;
const empty = () => ({ managerEmails: [], supervisorEmails: [], teacherEmails: [] });
const corrections = () => ({ ...empty(), supervisorEmails: ["chenli104@xdf.cn", "zuwanlu@xdf.cn"] });
const response = (body, ok = true) => ({ ok, json: async () => structuredClone(body) });
const csv = "教师姓名,教师邮箱,纠正岗位,是否启用,备注\n祖万露,ZUWANLU@XDF.CN,主管,是,岗位描述不准确\n陈利,chenli104@xdf.cn,助理主管,是,已确认\n停用,disabled@xdf.cn,经理,否,\n";

test("按邮箱解析启用的纠偏，兼容列顺序，忽略停用记录", () => {
  assert.deepEqual(parseInspectionRoleCorrections(csv), corrections());
  assert.deepEqual(parseInspectionRoleCorrections("是否启用,纠正岗位,教师邮箱\n是,经理,manager@xdf.cn\n是,教师,teacher@xdf.cn"), { managerEmails: ["manager@xdf.cn"], supervisorEmails: [], teacherEmails: ["teacher@xdf.cn"] });
  assert.deepEqual(parseInspectionRoleCorrections("教师邮箱,纠正岗位,是否启用\nx@xdf.cn,主管,否"), empty());
});

test("冲突邮箱、岗位错字和缺失启用状态不会静默漏纠偏", () => {
  for (const data of ["a@xdf.cn,主管,是\na@xdf.cn,教师,是", "a@xdf.cn,主观,是", "无邮箱,主管,是", "a@xdf.cn,主管,"]) {
    assert.throws(() => parseInspectionRoleCorrections(`教师邮箱,纠正岗位,是否启用\n${data}`));
  }
});

test("纠偏先修正岗位分类，只作用于本次在职邮箱，不改变原始名单", () => {
  const roster = { emails: new Set(["zuwanlu@xdf.cn", "manager@xdf.cn"]), roleExcludedEmails: new Set(["zuwanlu@xdf.cn"]), unsubmittedExtraExcludedEmails: new Set() };
  const result = applyInspectionRoleCorrections(roster, { ...corrections(), managerEmails: ["manager@xdf.cn"] });
  assert.equal(result.roleExcludedEmails.has("zuwanlu@xdf.cn"), false);
  assert.equal(result.roleExcludedEmails.has("manager@xdf.cn"), true);
  assert.deepEqual([...result.unsubmittedExtraExcludedEmails], ["zuwanlu@xdf.cn"]);
  assert.equal(roster.roleExcludedEmails.has("zuwanlu@xdf.cn"), true);
});

test("接口使用现有腾讯文档页签，仅返回邮箱与抽检分类，不公开姓名备注或凭证", async (context) => {
  const previous = process.env.TENCENT_DOCS_TOKEN;
  process.env.TENCENT_DOCS_TOKEN = "test-secret";
  context.after(() => { if (previous === undefined) delete process.env.TENCENT_DOCS_TOKEN; else process.env.TENCENT_DOCS_TOKEN = previous; });
  context.mock.method(globalThis, "fetch", async (url, options) => {
    assert.equal(url, "https://docs.qq.com/openapi/mcp");
    assert.equal(options.headers.Authorization, "test-secret");
    const call = JSON.parse(options.body);
    assert.equal(call.params.arguments.sheet_id, "0y559g");
    return response({ result: { content: [{ type: "text", text: JSON.stringify({ csv_data: csv }) }] } });
  });
  const res = { headers: {}, status(code) { this.code = code; return this; }, setHeader(name, value) { this.headers[name] = value; }, json(body) { this.body = body; } };
  await handler({ method: "GET" }, res);
  assert.equal(res.code, 200);
  assert.deepEqual(res.body, corrections());
  assert.equal(res.headers["Cache-Control"], "private, no-store");
  assert.doesNotMatch(JSON.stringify(res.body), /祖万露|岗位描述不准确|test-secret/u);
  await handler({ method: "POST" }, res);
  assert.equal(res.code, 405);
});

test("每次重新读取纠偏；停用立即影响下一次生成，服务故障不退回无纠偏状态", async (context) => {
  let current = corrections();
  context.mock.method(globalThis, "fetch", async (url, options) => {
    assert.equal(url, "/api/role-corrections");
    assert.equal(options.cache, "no-store");
    return response(current);
  });
  assert.deepEqual(await loadInspectionRoleCorrections(), corrections());
  current = empty();
  assert.deepEqual(await loadInspectionRoleCorrections(), empty());
  current = null;
  await assert.rejects(loadInspectionRoleCorrections(), /格式无效/u);
});

function fileFromRows(name, rows) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "明细");
  return new File([XLSX.write(wb, { type: "array", bookType: "xlsx" })], name);
}

test("两种优先模式、在线名单及手动上传均在抽取阶段排除祖万露加抽，普通抽检保留", async (context) => {
  context.mock.method(globalThis, "fetch", async (url) => {
    assert.equal(url, "/api/role-corrections");
    return response(corrections());
  });
  const feedbackFile = fileFromRows("反馈.xlsx", [
    ["老师姓名", "学员姓名", "学员号", "课次ID", "课次开始时间", "课次结束时间", "老师邮箱", "是否生成报告"],
    ["祖万露", "学员甲", "1", "NORMAL", "2026-10-19 10:00", "2026-10-19 12:00", "zuwanlu@xdf.cn", "是"],
    ["祖万露", "学员乙", "2", "EXTRA", "2026-10-19 14:00", "2026-10-19 16:00", "zuwanlu@xdf.cn", "否"],
    ["普通教师", "学员丙", "3", "TEACHER", "2026-10-19 10:00", "2026-10-19 12:00", "teacher@xdf.cn", "否"],
  ]);
  const rosterFile = fileFromRows("在职教师明细20261019.xlsx", [["邮箱", "岗位短描述"], ["zuwanlu@xdf.cn", "益智新师师训助理"], ["teacher@xdf.cn", "教师"]]);
  const defaultRoster = { snapshotDate: "2026-10-19", sourceFile: "在职教师明细20261019.xlsx", sourceSha256: "a".repeat(64), rowCount: 2, matchedEmailRows: 2, emails: ["zuwanlu@xdf.cn", "teacher@xdf.cn"], roleExcludedEmails: [], unsubmittedExtraExcludedEmails: [] };
  for (const priorityMode of ["coverage", "unreported"]) {
    for (const source of [{ rosterFile }, { defaultRoster }]) {
      const messages = [];
      await processInspection({ type: "process", mode: "inspection", feedbackFile, ...source, sampleCount: 10, attempt: 1, includeExplanation: true, includeReason: true, priorityMode, focusTeacherNames: ["祖万露"], teacherScoreRows: [] }, { postMessage: message => messages.push(message) });
      const result = messages.at(-1);
      assert.equal(result.summary.excludedSupervisorUnsubmittedRows, 1);
      assert.equal(result.summary.extraSelectedRows, 1);
      assert.equal(result.summary.normalSelectedRows, 1);
      const wb = XLSX.read(Buffer.concat(result.chunks.map(chunk => Buffer.from(chunk))), { type: "buffer" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets["抽检名单"], { header: 1 });
      assert.equal(rows.length, 3);
      assert.equal(rows.some(row => row[0] === "祖万露" && row[4] === "报告未生成容量外加抽"), false);
      assert.equal(rows.some(row => row[0] === "祖万露"), true);
    }
  }
});
