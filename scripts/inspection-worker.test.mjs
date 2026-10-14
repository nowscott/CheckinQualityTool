import assert from "node:assert/strict";
import test from "node:test";

const { loadInspectionWorkerUrl } = await import("../lib/inspectionWorker.js");

test("页面保留打开时的旧地址，每次生成仍读取当前抽检程序", async (context) => {
  let publishedUrl = "/assets/inspectionEntry-first.js";
  const calls = [];
  context.mock.method(globalThis, "fetch", async (url, options) => {
    calls.push({ url, options });
    return { ok: true, json: async () => ({ url: publishedUrl }) };
  });
  assert.equal(await loadInspectionWorkerUrl(), publishedUrl);
  publishedUrl = "/assets/inspectionEntry-second.js";
  assert.equal(await loadInspectionWorkerUrl(), publishedUrl);
  assert.deepEqual(calls, Array(2).fill({ url: "/data/inspection-worker.json", options: { cache: "no-store" } }));
});

test("本地开发使用开发服务器的抽检程序", async (context) => {
  context.mock.method(globalThis, "fetch", () => { throw new Error("不应请求发布清单"); });
  assert.equal(await loadInspectionWorkerUrl("/src/worker/inspectionEntry.ts?worker_file&type=classic"), "/src/worker/inspectionEntry.ts?worker_file&type=classic");
});

test("发布清单不可用时不会回退到旧抽检程序", async (context) => {
  context.mock.method(globalThis, "fetch", async () => ({ ok: false }));
  await assert.rejects(loadInspectionWorkerUrl(), /抽检程序加载失败/u);
});

test("只加载本站发布的抽检程序", async (context) => {
  context.mock.method(globalThis, "fetch", async () => ({ ok: true, json: async () => ({ url: "https://example.com/worker.js" }) }));
  await assert.rejects(loadInspectionWorkerUrl(), /抽检程序地址无效/u);
});
