import { readFile, writeFile } from "node:fs/promises";
import { basename, resolve } from "node:path";
import vm from "node:vm";

const root = resolve(import.meta.dirname, "..");
const [listPath, chatPath, outputPath = "/tmp/typescript-worker-result.xlsx"] = process.argv.slice(2);
if (!listPath || !chatPath) {
  throw new Error("用法：node scripts/regression-worker.mjs <课堂反馈名单.xlsx> <聊天质检明细.xlsx> [输出.xlsx]");
}

const assets = await import("node:fs/promises").then(({ readdir }) => readdir(resolve(root, "dist/assets")));
let workerFile = "";
for (const asset of assets.filter((name) => name.endsWith(".js"))) {
  const source = await readFile(resolve(root, "dist/assets", asset), "utf8");
  if (source.includes("onmessage") && source.includes("postMessage") && !source.includes("createRoot")) {
    workerFile = resolve(root, "dist/assets", asset);
    break;
  }
}
if (!workerFile) throw new Error("找不到构建后的 Worker，请先运行 npm run build。");

const listBuffer = await readFile(resolve(listPath));
const chatBuffer = await readFile(resolve(chatPath));
const whitelistCsv = await readFile(resolve(root, "public/data/whitelist.csv"), "utf8");
const workerSources = new Map([
  [
    resolve(root, "dist/vendor/xlsx.full.min.js"),
    await readFile(resolve(root, "dist/vendor/xlsx.full.min.js"), "utf8"),
  ],
]);

let context;
const result = new Promise((resolveResult, rejectResult) => {
  const startedAt = performance.now();
  globalThis.__workerComplete = (message) => {
    if (message.type === "error") rejectResult(new Error(message.message));
    if (message.type === "complete") {
      resolveResult({
        ...message,
        elapsedMs: Number((performance.now() - startedAt).toFixed(1)),
      });
    }
  };
});

const sandbox = {
  console,
  setTimeout,
  clearTimeout,
  TextEncoder,
  TextDecoder,
  Uint8Array,
  ArrayBuffer,
  fetch: async (url) => {
    if (url !== "/vendor/xlsx.full.min.js") throw new Error(`不支持的 fetch 地址：${url}`);
    return {
      ok: true,
      text: async () => workerSources.get(resolve(root, "dist/vendor/xlsx.full.min.js")),
    };
  },
  postMessage: (message) => globalThis.__workerComplete(message),
  importScripts: (...urls) => {
    for (const url of urls) {
      const sourcePath = resolve(root, "dist", url.replace(/^\//, ""));
      const source = workerSources.get(sourcePath);
      if (!source) throw new Error(`找不到 Worker 依赖：${sourcePath}`);
      vm.runInContext(source, context, { filename: sourcePath });
    }
  },
};

context = vm.createContext(sandbox);
context.self = context;
vm.runInContext(await readFile(workerFile, "utf8"), context, { filename: workerFile });

const file = (path, buffer) => ({
  name: basename(path),
  size: buffer.byteLength,
  arrayBuffer: async () => new Uint8Array(buffer),
});

await context.self.onmessage({
  data: {
    type: "process",
    listFile: file(listPath, listBuffer),
    chatFile: file(chatPath, chatBuffer),
    weekLabel: "auto",
    useSingle: false,
    whitelistCsv,
  },
});

const message = await result;
const chunks = message.chunks || [new Uint8Array(message.buffer)];
const outputBytes = Buffer.concat(chunks.map((chunk) => Buffer.from(chunk)));
await writeFile(resolve(outputPath), outputBytes);
console.log(JSON.stringify({
  worker: basename(workerFile),
  output: resolve(outputPath),
  bytes: message.byteLength || outputBytes.byteLength,
  elapsedMs: message.elapsedMs,
  summary: message.summary,
}, null, 2));
