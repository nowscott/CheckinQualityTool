import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import vm from "node:vm";
import { webcrypto, createHash } from "node:crypto";

const [feedbackPath, rosterPath, projectionPath, outputDir = "/tmp/checkin-roster-regression", baseUrl] = process.argv.slice(2);
if (!feedbackPath || !rosterPath || !projectionPath) throw new Error("用法：node scripts/regression-inspection-roster.mjs <真实课程.xlsx> <同版在职.xlsx> <接口投影.json> [输出目录] [线上站点URL]");
const projection = JSON.parse(await fs.readFile(projectionPath, "utf8"));
const manifest = baseUrl
  ? await fetch(`${baseUrl}/data/inspection-worker.json`, { cache: "no-store" }).then((r) => { if (!r.ok) throw new Error("线上 Worker 清单失败"); return r.json(); })
  : JSON.parse(await fs.readFile("dist/data/inspection-worker.json", "utf8"));
const source = baseUrl ? await fetch(`${baseUrl}${manifest.url}`).then((r) => r.text()) : await fs.readFile(path.join("dist", manifest.url), "utf8");
const vendor = baseUrl ? await fetch(`${baseUrl}/vendor/xlsx.full.min.js`).then((r) => r.text()) : await fs.readFile("public/vendor/xlsx.full.min.js", "utf8");
const file = async (filePath) => {
  const bytes = await fs.readFile(filePath);
  return { name: path.basename(filePath), size: bytes.length, arrayBuffer: async () => Uint8Array.from(bytes).buffer };
};
const feedbackFile = await file(feedbackPath);
const rosterFile = await file(rosterPath);
assert.equal(createHash("sha256").update(await fs.readFile(rosterPath)).digest("hex"), projection.sourceSha256);

async function run(rosterSource, priorityMode) {
  const messages = [];
  const sandbox = { console, setTimeout, clearTimeout, TextEncoder, TextDecoder, ArrayBuffer, Uint8Array, Date, crypto: webcrypto,
    fetch: async (url) => { assert.equal(url, "/vendor/xlsx.full.min.js"); return { ok: true, text: async () => vendor }; },
    postMessage: (message) => messages.push(message),
  };
  const context = vm.createContext(sandbox); context.self = context;
  vm.runInContext(source, context);
  await context.onmessage({ data: { type: "process", mode: "inspection", feedbackFile, ...rosterSource, sampleCount: 1000, attempt: 1, includeExplanation: true, includeReason: true, priorityMode, focusTeacherNames: [], teacherScoreRows: [] } });
  const result = messages.at(-1);
  if (result.type === "error") throw new Error(result.message);
  assert.equal(result.type, "inspectionComplete");
  const bytes = Buffer.concat(result.chunks.map((chunk) => Buffer.from(chunk)));
  const wb = context.XLSX.read(bytes, { type: "buffer" });
  const rows = context.XLSX.utils.sheet_to_json(wb.Sheets["抽检名单"], { header: 1, defval: "" });
  assert.deepEqual(Array.from(rows[0]), ["教师姓名", "学员姓名", "课次日期", "课次时间", "入选原因"]);
  assert.equal(rows.length - 1, result.summary.selectedRows);
  const explanation = context.XLSX.utils.sheet_to_json(wb.Sheets["处理说明"], { header: 1 });
  assert.equal(explanation.some((r) => r[0] === "在职快照日期" && r[1] === projection.snapshotDate), true);
  assert.equal(explanation.some((r) => r[0] === "在职源文件 SHA-256" && r[1] === projection.sourceSha256), true);
  return { result, bytes, rows: JSON.parse(JSON.stringify(rows)) };
}
await fs.mkdir(outputDir, { recursive: true });
for (const mode of ["coverage", "unreported"]) {
  const api = await run({ defaultRoster: projection }, mode);
  const manual = await run({ rosterFile }, mode);
  assert.deepEqual(JSON.parse(JSON.stringify(api.result.summary)), JSON.parse(JSON.stringify(manual.result.summary)));
  assert.deepEqual(api.rows, manual.rows);
  assert.deepEqual(JSON.parse(JSON.stringify(api.result.rosterSnapshot)), JSON.parse(JSON.stringify(manual.result.rosterSnapshot)));
  await fs.writeFile(path.join(outputDir, `${mode}_接口.xlsx`), api.bytes);
  await fs.writeFile(path.join(outputDir, `${mode}_上传.xlsx`), manual.bytes);
  console.log(JSON.stringify({ mode, snapshot: api.result.rosterSnapshot, summary: api.result.summary, equalToManual: true }));
}
