import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";

globalThis.self = { postMessage() {} };
globalThis.XLSX = {
  utils: { sheet_to_json: (sheet) => sheet.rows },
  SSF: { parse_date_code: () => null },
};

const { buildStageReportBeautifyOutput } = await import("../worker/stageReportBeautifyWriter.js");

const stageHeaders = [
  "教师姓名", "学生姓名", "学号", "暑假最后一节课时间", "师训组长", "助理主管", "教研组",
  "是否发送阶段性报告（系统数据）", "是否申诉", "申诉情况详情", "是否需要发送",
  "数据变动时间（并非最终导入数据时间）", "申诉是否生效", "是否已发送（申诉+系统）",
];
const appealHeaders = ["教师姓名", "学生姓名", "申诉情况说明", "申诉情况详情", "相关截图上传处"];
const teacherHeaders = [
  "教师姓名", "师训组长", "助理主管", "教研组", "窗口期报告应发送", "窗口期报告已发送", "窗口期报告发送率",
  "阶段性报告应发送0824", "阶段性报告已发送0824", "阶段性报告发送率0824", "阶段性报告未发送学生姓名0824",
  "阶段性报告应发送0819", "阶段性报告应发送0805", "阶段性报告已发送0805", "阶段性报告已发送0819",
  "阶段性报告发送率0805", "阶段性报告发送率0819", "阶段性报告未发送学生姓名0805", "窗口期报告未发送学生姓名",
  "阶段性报告未发送学生姓名0819", "是否全部完成", "是否已通知", "阶段性报告数据更新时间0824",
  "阶段性报告数据更新时间0819", "阶段性报告数据更新时间0805",
];
const trainingHeaders = [
  "师训组长", "助理主管", "教研组", "窗口期报告应发送", "窗口期报告已发送", "窗口期报告发送率",
  "阶段性报告应发送0824", "阶段性报告已发送0824", "阶段性报告发送率0824", "阶段性报告应发送0819",
  "阶段性报告已发送0819", "阶段性报告发送率0819", "阶段性报告应发送0805", "阶段性报告已发送0805",
  "阶段性报告发送率0805",
];

function row(headers, values) {
  return headers.map((header) => values[header] ?? "");
}

function fixtureWorkbook({ newDetailOrder, windowUpdateHeader }) {
  const detailHeaders = [...stageHeaders];
  if (newDetailOrder) detailHeaders.splice(6, 2, detailHeaders[7], detailHeaders[6]);
  const sheets = {};
  const sheetNames = [];
  const add = (name, headers, rows) => {
    sheetNames.push(name);
    sheets[name] = { rows: [headers, ...rows] };
  };
  const detailValues = {
    教师姓名: "张老师", 学生姓名: "陈一", 学号: "GZ1", 暑假最后一节课时间: "2026-08-24 10:00-12:00",
    师训组长: "组长", 助理主管: "主管", 教研组: "高中双语", "是否发送阶段性报告（系统数据）": "是",
    "是否已发送（申诉+系统）": "是",
  };
  add("窗口期报告发送明细", ["教师姓名", "学生姓名", "学号", "师训组长", "助理主管", "是否发送窗口期报告", "数据变动时间（并非最终导入数据时间）"], [
    ["张老师", "陈一", "GZ1", "组长", "主管", "是", "2026-08-25 00:20"],
  ]);
  ["0824非窗口期暑期在读学员阶段性报告发送明细", "0819非窗口期暑期在读学员阶段性报告发送明细", "0805非窗口期暑期在读学员阶段性报告发送明细"]
    .forEach((name) => add(name, detailHeaders, [row(detailHeaders, detailValues)]));
  ["0824申诉结果", "0819申诉结果", "0805申诉结果"]
    .forEach((name) => add(name, appealHeaders, [["张老师", "陈一", "已发送", "", ""]]));

  const summaryValues = {
    教师姓名: "张老师", 师训组长: "组长", 助理主管: "主管", 教研组: "高中双语",
    窗口期报告应发送: 1, 窗口期报告已发送: 1, 窗口期报告发送率: "100%",
    阶段性报告应发送0824: 1, 阶段性报告已发送0824: 1, 阶段性报告发送率0824: "100%",
    阶段性报告应发送0819: 1, 阶段性报告已发送0819: 1, 阶段性报告发送率0819: "100%",
    阶段性报告应发送0805: 1, 阶段性报告已发送0805: 1, 阶段性报告发送率0805: "100%",
    阶段性报告数据更新时间0824: "2026-08-25 00:21", 阶段性报告数据更新时间0819: "2026-08-25 00:20",
    阶段性报告数据更新时间0805: "2026-08-25 00:19", [windowUpdateHeader]: "2026-08-25 00:22",
  };
  add("教师维度明细", [...teacherHeaders, windowUpdateHeader], [row([...teacherHeaders, windowUpdateHeader], summaryValues)]);
  add("组长维度汇总", trainingHeaders, [row(trainingHeaders, summaryValues)]);
  return { SheetNames: sheetNames, Sheets: sheets };
}

async function writeOutput(output) {
  const dir = await mkdtemp(join(tmpdir(), "stage-report-beautify-"));
  const file = join(dir, "result.xlsx");
  await writeFile(file, Buffer.concat(output.chunks.map((chunk) => Buffer.from(chunk))));
  return { dir, file };
}

function unzipText(file, entry) {
  const result = spawnSync("unzip", ["-p", file, entry], { encoding: "utf8" });
  assert.equal(result.status, 0, result.stderr || result.stdout);
  return result.stdout;
}

test("公示版兼容 0824、0819、0805 三批数据及新版更新时间字段", async () => {
  for (const fixture of [
    { newDetailOrder: false, windowUpdateHeader: "窗口期数据更新时间" },
    { newDetailOrder: true, windowUpdateHeader: "窗口期报告数据更新时间" },
  ]) {
    const output = buildStageReportBeautifyOutput(fixtureWorkbook(fixture));
    assert.equal(output.dataTime, "2026-08-25 00:22");
    assert.deepEqual(output.counts, { stageRows: 3, windowRows: 1, teacherRows: 1, appealRows: 3, sheets: 11 });

    const written = await writeOutput(output);
    try {
      const check = spawnSync("unzip", ["-t", written.file], { encoding: "utf8" });
      assert.equal(check.status, 0, check.stderr || check.stdout);
      assert.match(unzipText(written.file, "xl/workbook.xml"), /0819之后结课阶段性报告明细/u);
      const groupXml = unzipText(written.file, "xl/worksheets/sheet8.xml");
      assert.match(groupXml, /阶段性报告应发送情况（0824，0819之后结课）/u);
      assert.equal((groupXml.match(/阶段性报告应发送数/g) || []).length, 3);
      assert.match(unzipText(written.file, "xl/worksheets/sheet9.xml"), /阶段性报告应发送数0824/u);
    } finally {
      await rm(written.dir, { recursive: true, force: true });
    }
  }
});
