import assert from "node:assert/strict";
import { mkdtemp, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { spawnSync } from "node:child_process";
import test from "node:test";

globalThis.self = { postMessage() {} };
const { buildWorkbook } = await import("../worker/excelWriter.js");

test("Excel ZIP 生产路径保留分块且可完整解压", async () => {
  const chunks = buildWorkbook([{
    name: "测试",
    rows: [{ 姓名: "张三", 状态: "已发送" }],
    columns: ["姓名", "状态"],
    widths: { 姓名: 18, 状态: 18 },
  }]);
  assert.ok(Array.isArray(chunks));
  assert.ok(chunks.length > 0);
  const bytes = Buffer.concat(chunks.map((chunk) => Buffer.from(chunk)));
  const dir = await mkdtemp(join(tmpdir(), "checkin-xlsx-chunk-"));
  const file = join(dir, "result.xlsx");
  await writeFile(file, bytes);
  const check = spawnSync("unzip", ["-t", file], { encoding: "utf8" });
  assert.equal(check.status, 0, check.stderr || check.stdout);
});
