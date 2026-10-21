/// <reference path="./sheetjs.d.ts" />

import { buildInspectionOutput } from "./inspectionWriter";
import { buildInspectionSelection, normalizeInspectionNumber } from "./inspectionSampler";
import { rosterFromSnapshot } from "../lib/rosterSnapshot.js";
import { applyInspectionRoleCorrections, loadInspectionRoleCorrections } from "../lib/inspectionRoleCorrections.js";
import { parseInspectionRoster, parseInspectionRows, sha256File } from "./inspectionParser";
import { progress } from "./progress";
import { readWorkbook } from "./excelReader";
import type { InspectionRequest, WorkerResponse } from "../types/worker";
import type { RosterInfo } from "./inspectionTypes";
import type { RosterSnapshotMeta } from "../lib/rosterSnapshot.js";

interface WorkerScope {
  postMessage(message: unknown, transfer?: Transferable[]): void;
}

function postInspectionComplete(scope: WorkerScope, output: ReturnType<typeof buildInspectionOutput>, rosterSnapshot: RosterSnapshotMeta) {
  const byteLength = output.chunks.reduce((total, chunk) => total + chunk.byteLength, 0);
  const transfer = [...new Set(output.chunks.map((chunk) => chunk.buffer))] as ArrayBuffer[];
  scope.postMessage({
    type: "inspectionComplete",
    chunks: output.chunks,
    byteLength,
    filename: output.filename,
    summary: output.summary,
    priority: output.priority,
    rosterSnapshot,
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
  let rosterSourceFile: string;
  if (request.rosterFile) {
    progress("正在读取最新在职明细", `${request.rosterFile.name} · ${(request.rosterFile.size / 1024 / 1024).toFixed(1)} MB`, 30);
    const rosterWorkbook = await readWorkbook(request.rosterFile, 30, 46, "最新在职明细");
    roster = parseInspectionRoster(rosterWorkbook, request.rosterFile.name);
    rosterSha256 = await sha256File(request.rosterFile);
    rosterSourceFile = request.rosterFile.name;
  } else {
    if (!request.defaultRoster) throw new Error("本次处理没有取得最新在职名单，请重新开始或手动上传指定版本。");
    roster = rosterFromSnapshot(request.defaultRoster);
    rosterSha256 = request.defaultRoster.sourceSha256;
    rosterSourceFile = request.defaultRoster.sourceFile;
    progress("正在使用本次在职快照", `${roster.snapshotDate} · 同一次处理固定使用此版本。`, 38);
  }
  progress("正在读取岗位纠偏", "按腾讯文档中的已启用邮箱记录纠正岗位后再抽取。", 48);
  roster = applyInspectionRoleCorrections(roster, await loadInspectionRoleCorrections());
  const excludedRosterRoleCount = [...roster.roleExcludedEmails].filter((email) => roster.emails.has(email)).length;
  const excludedSupervisorUnsubmittedCount = [...roster.unsubmittedExtraExcludedEmails]
    .filter((email) => roster.emails.has(email) && !roster.roleExcludedEmails.has(email)).length;
  progress(
    "正在核对在职教师",
    `在职明细识别到 ${roster.emails.size.toLocaleString()} 个有效邮箱；岗位含“经理”的 ${excludedRosterRoleCount.toLocaleString()} 人已排除抽检；主管及助理主管 ${excludedSupervisorUnsubmittedCount.toLocaleString()} 人仍参与普通抽检，但未生成报告不加抽。`,
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
    priorityMode: request.priorityMode,
    focusTeacherNames: request.focusTeacherNames,
    teacherScoreRows: request.teacherScoreRows,
    teacherScoresByEmail: request.teacherScoresByEmail,
  });
  progress(
    "抽检名单生成完成",
    `候选 ${selection.stats.eligibleRows.toLocaleString()} 条，覆盖 ${selection.stats.eligibleTeachers.toLocaleString()} 位教师，抽检 ${selection.stats.selectedRows.toLocaleString()} 条。`,
    72,
  );
  const output = buildInspectionOutput(selection, request.includeExplanation, request.includeReason);
  progress("Excel 已生成", "正在准备下载抽检名单。", 90);
  postInspectionComplete(scope, output, { snapshotDate: roster.snapshotDate, sourceFile: rosterSourceFile, sourceSha256: rosterSha256, rowCount: roster.rowCount });
}
