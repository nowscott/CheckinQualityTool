import assert from "node:assert/strict";
import test from "node:test";
import handler from "../api/inspection-coverage.js";
function output() {
  return { code: 0, headers: {}, status(code) { this.code = code; return this; }, setHeader(k,v) { this.headers[k] = v; }, json(value) { this.body = value; } };
}
test("覆盖 API 只返回选定月份身份与计数，不返回学员或课程记录", async () => {
  const original = globalThis.fetch;
  const token = process.env.TENCENT_DOCS_TOKEN;
  process.env.TENCENT_DOCS_TOKEN = "test-only";
  try {
    globalThis.fetch = async (_url, options) => {
      const args = JSON.parse(options.body).params.arguments;
      assert.equal(args.sheet_id, "ih60c3");
      assert.equal(args.end_col, 5);
      return { ok: true, json: async () => ({ result: { content: [{ type: "text", text: JSON.stringify({ csv_data: "教师姓名,学员姓名,课次日期,课次时间,入选原因,抽检日期\n甲,不应返回学员,2026/9/30,10:00-12:00,,2026-10-21\n乙,另一学员,2026/10/1,10:00-12:00,,\n丙,另一学员,2026/10/1,10:00-12:00,,2026-09-30" }) }] } }) };
    };
    const res = output();
    await handler({ method: "GET", query: { month: "2026-10" } }, res);
    assert.equal(res.code, 200);
    assert.equal(res.body.rowCount, 1);
    assert.equal(res.body.undatedRowCount, 1);
    assert.deepEqual(res.body.teachers, [{ teacherEmail: "", teacherName: "甲" }]);
    assert.equal(res.headers["Cache-Control"], "no-store");
    assert.ok(!JSON.stringify(res.body).includes("不应返回学员"));
  } finally { globalThis.fetch = original; if (token === undefined) delete process.env.TENCENT_DOCS_TOKEN; else process.env.TENCENT_DOCS_TOKEN = token; }
});
test("覆盖 API 拒绝无效月份与写入，读取失败返回错误而非空台账", async () => {
  const badMonth = output();
  await handler({ method: "GET", query: { month: "2026-99" } }, badMonth);
  assert.equal(badMonth.code, 400);
  const write = output();
  await handler({ method: "POST" }, write);
  assert.equal(write.code, 405);
  const old = globalThis.fetch;
  const token = process.env.TENCENT_DOCS_TOKEN;
  process.env.TENCENT_DOCS_TOKEN = "test-only";
  globalThis.fetch = async () => { throw new Error("upstream unavailable"); };
  try {
    const failed = output();
    await handler({ method: "GET", query: { month: "2026-10" } }, failed);
    assert.equal(failed.code, 502);
    assert.equal(failed.body.teachers, undefined);
  } finally { globalThis.fetch = old; if (token === undefined) delete process.env.TENCENT_DOCS_TOKEN; else process.env.TENCENT_DOCS_TOKEN = token; }
});
