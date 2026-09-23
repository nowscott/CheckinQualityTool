/// <reference path="./sheetjs.d.ts" />

import { buildInspectionOutput } from "./inspectionWriter";
import { buildInspectionSelection, normalizeInspectionNumber } from "./inspectionSampler";
import { parseInspectionRoster, parseInspectionRows, sha256File } from "./inspectionParser";
import { progress } from "./progress";
import { readWorkbook } from "./excelReader";
import type { InspectionRequest, WorkerResponse } from "../types/worker";
import type { DefaultRosterAsset, DefaultRosterRoleExclusionAsset, RosterInfo } from "./inspectionTypes";

interface WorkerScope {
  postMessage(message: unknown, transfer?: Transferable[]): void;
}

function postInspectionComplete(scope: WorkerScope, output: ReturnType<typeof buildInspectionOutput>) {
  const byteLength = output.chunks.reduce((total, chunk) => total + chunk.byteLength, 0);
  const transfer = [...new Set(output.chunks.map((chunk) => chunk.buffer))] as ArrayBuffer[];
  scope.postMessage({
    type: "inspectionComplete",
    chunks: output.chunks,
    byteLength,
    filename: output.filename,
    summary: output.summary,
    historyPayload: output.historyPayload,
  } satisfies WorkerResponse, transfer);
}

export async function processInspection(request: InspectionRequest, scope: WorkerScope) {
  const sampleCount = normalizeInspectionNumber(request.sampleCount);
  if (!request.feedbackFile) throw new Error("请上传课堂反馈明细。");
  if (sampleCount < 1) throw new Error("抽检数必须是大于 0 的整数。");

  progress("正在读取课程反馈明细", `${request.feedbackFile.name} · ${(request.feedbackFile.size / 1024 / 1024).toFixed(1)} MB`, 4);
  const feedbackWorkbook = await readWorkbook(request.feedbackFile, 4, 28, "课堂反馈明细", true);
  const parsed = parseInspectionRows(feedbackWorkbook);
  const rows = parsed.rows;
  if (!rows.length) throw new Error("课堂反馈明细没有可处理的课程记录。");
  let roster: RosterInfo;
  let rosterSha256: string;
  if (request.rosterFile) {
    progress("正在读取最新在职明细", `${request.rosterFile.name} · ${(request.rosterFile.size / 1024 / 1024).toFixed(1)} MB`, 30);
    const rosterWorkbook = await readWorkbook(request.rosterFile, 30, 46, "最新在职明细");
    roster = parseInspectionRoster(rosterWorkbook, request.rosterFile.name);
    rosterSha256 = await sha256File(request.rosterFile);
  } else {
    progress("正在读取内置在职明细", "未上传在职明细，使用项目内置的最新快照。", 38);
    const response = await fetch("/data/inspection-roster.json", { cache: "no-store" });
    if (!response.ok) throw new Error("内置在职明细读取失败，请上传一份最新在职明细。 ");
    const asset = (await response.json()) as DefaultRosterAsset;
    let roleExcludedEmails = asset.roleExcludedEmails || [];
    if (!roleExcludedEmails.length) {
      const roleResponse = await fetch("/data/inspection-role-exclusions.json", { cache: "no-store" });
      if (roleResponse.ok) {
        const roleAsset = (await roleResponse.json()) as DefaultRosterRoleExclusionAsset;
        roleExcludedEmails = Array.isArray(roleAsset.emails) ? roleAsset.emails : [];
      }
    }
    roster = {
      emails: new Set(asset.emails),
      roleExcludedEmails: new Set(roleExcludedEmails),
      sourceName: `${asset.sourceFile}（项目内置）`,
      snapshotDate: asset.snapshotDate,
      rowCount: asset.rowCount,
      matchedEmailRows: asset.matchedEmailRows,
    };
    rosterSha256 = asset.sourceSha256;
  }
  progress(
    "正在核对在职教师",
    `在职明细识别到 ${roster.emails.size.toLocaleString()} 个有效邮箱，排除主管/经理 ${roster.roleExcludedEmails.size.toLocaleString()} 人。`,
    52,
  );

  const sourceSha256 = await sha256File(request.feedbackFile);
  const selection = buildInspectionSelection(rows, roster, {
    sampleCount,
    attempt: Math.max(1, Math.floor(request.attempt || 1)),
    sourceSha256,
    rosterSha256,
    sourceName: request.feedbackFile.name,
    sourceColumns: parsed.columns,
    batchKind: request.batchKind,
  });
  progress(
    "抽检名单生成完成",
    `候选 ${selection.stats.eligibleRows.toLocaleString()} 条，覆盖 ${selection.stats.eligibleTeachers.toLocaleString()} 位教师，抽检 ${selection.stats.selectedRows.toLocaleString()} 条。`,
    72,
  );
  const output = buildInspectionOutput(selection, request.includeExplanation);
  progress("Excel 已生成", "正在交给页面保存抽检历史。", 90);
  postInspectionComplete(scope, output);
}
