/// <reference path="./sheetjs.d.ts" />

import { processInspection } from "./inspectionProcess";
import { ensureSheetJs } from "./sheetJsLoader";
import type { InspectionRequest } from "../types/worker";

const workerScope = self as unknown as Parameters<typeof processInspection>[1];

self.onmessage = async ({ data }: MessageEvent<InspectionRequest>) => {
  if (data.type !== "process" || data.mode !== "inspection") return;
  try {
    await ensureSheetJs();
    await processInspection(data, workerScope);
  } catch (error) {
    workerScope.postMessage({ type: "error", message: error instanceof Error ? error.message : String(error) });
  }
};
