import assert from "node:assert/strict";
import test from "node:test";

globalThis.self = { postMessage() {} };
const { buildStageReportDingTalkOutput } = await import("../worker/stageReportDingTalkWriter.js");

test("阶段性报告有未决行时阻止钉钉文件导出", () => {
  assert.throws(() => buildStageReportDingTalkOutput(
    { columns: ["教师姓名"], trainingLeadGroups: new Map() },
    {
      detailRows: [],
      teacherRows: [],
      counts: {},
      unresolvedRows: [{ sourceRowNumber: 12, reason: "教师姓名对应多个聊天邮箱" }],
    },
  ), /第12行：教师姓名对应多个聊天邮箱/);
});
