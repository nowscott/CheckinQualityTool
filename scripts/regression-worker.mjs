import { readFile, writeFile } from "node:fs/promises";
import { basename, resolve } from "node:path";
import vm from "node:vm";

const root = resolve(import.meta.dirname, "..");
const args = process.argv.slice(2);
const includeCleanChatsIndex = args.indexOf("--include-clean-chats");
const includeCleanChats = includeCleanChatsIndex >= 0;
if (includeCleanChats) args.splice(includeCleanChatsIndex, 1);
const appealArgIndex = args.findIndex((arg) => arg.startsWith("--appeal="));
const appealPath = appealArgIndex >= 0 ? args[appealArgIndex].slice("--appeal=".length) : "";
if (appealArgIndex >= 0) args.splice(appealArgIndex, 1);
const summaryPaths = args
  .filter((arg) => arg.startsWith("--summary="))
  .map((arg) => arg.slice("--summary=".length));
for (let index = args.length - 1; index >= 0; index -= 1) {
  if (args[index].startsWith("--summary=")) args.splice(index, 1);
}
const modeArg = args[0]?.startsWith("--mode=") ? args.shift() : "";
const mode = modeArg ? modeArg.split("=")[1] : "checkin";
let listPath = "";
let chatPaths = [];
let outputPath = mode === "reminder"
  ? "/tmp/reminder-worker-result.xlsx"
  : mode === "stageReport"
    ? "/tmp/stage-report-worker-result.xlsx"
    : mode === "stageReportBeautify"
      ? "/tmp/stage-report-beautify-worker-result.xlsx"
      : "/tmp/typescript-worker-result.xlsx";
if (mode === "stageReportBeautify") {
  listPath = args.shift() || "";
  if (args.length) outputPath = args.shift();
} else if (["reminder", "stageReport"].includes(mode)) {
  listPath = args.shift() || "";
  if (args.length > 1) outputPath = args.pop();
  chatPaths = args;
} else {
  const [currentListPath, chatPath, currentOutputPath = outputPath] = args;
  listPath = currentListPath;
  chatPaths = chatPath ? [chatPath] : [];
  outputPath = currentOutputPath;
}

if (!listPath || (mode !== "stageReportBeautify" && !chatPaths.length)) {
  throw new Error("用法：node scripts/regression-worker.mjs [--mode=checkin|reminder|stageReport|stageReportBeautify] <名单或源表.xlsx> [聊天.xlsx...] [输出.xlsx]");
}
if (!["checkin", "reminder", "stageReport", "stageReportBeautify"].includes(mode)) throw new Error(`不支持的 mode：${mode}`);

const assets = await import("node:fs/promises").then(({ readdir }) =>
  readdir(resolve(root, "dist/assets")),
);
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
const chatBuffers = await Promise.all(chatPaths.map((path) => readFile(resolve(path))));
const appealBuffer = appealPath ? await readFile(resolve(appealPath)) : null;
const summaryBuffers = await Promise.all(summaryPaths.map((path) => readFile(resolve(path))));
const whitelistCsv = await readFile(resolve(root, "public/data/whitelist.csv"), "utf8");
const workerSources = new Map([
  [
    resolve(root, "dist/vendor/xlsx.full.min.js"),
    await readFile(resolve(root, "dist/vendor/xlsx.full.min.js"), "utf8"),
  ],
]);

let context;
let complete;
const result = new Promise((resolveResult, rejectResult) => {
  complete = (message) => {
    if (message.type === "complete") resolveResult(message);
    if (message.type === "error") rejectResult(new Error(message.message));
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
  postMessage: (message) => complete(message),
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
  data: mode === "stageReportBeautify"
    ? {
        type: "process",
        mode: "stageReportBeautify",
        sourceFile: file(listPath, listBuffer),
      }
    : mode === "reminder"
    ? {
        type: "process",
        mode: "reminder",
        denominatorFile: file(listPath, listBuffer),
        appealFile: appealPath && appealBuffer ? file(appealPath, appealBuffer) : null,
        summaryFiles: summaryPaths.map((path, index) => file(path, summaryBuffers[index])),
        chatFiles: chatPaths.map((path, index) => file(path, chatBuffers[index])),
        includeCleanChats,
        includeResultColors: false,
        whitelistCsv,
      }
    : mode === "stageReport"
      ? {
          type: "process",
          mode: "stageReport",
          denominatorFile: file(listPath, listBuffer),
          chatFiles: chatPaths.map((path, index) => file(path, chatBuffers[index])),
        }
      : {
        type: "process",
        listFile: file(listPath, listBuffer),
        chatFile: file(chatPaths[0], chatBuffers[0]),
        weekLabel: "auto",
        useSingle: false,
        whitelistCsv,
      },
});

const message = await result;
await writeFile(resolve(outputPath), new Uint8Array(message.buffer));
console.log(JSON.stringify({
  worker: basename(workerFile),
  mode,
  output: resolve(outputPath),
  bytes: message.buffer.byteLength,
  summary: message.summary,
}, null, 2));
