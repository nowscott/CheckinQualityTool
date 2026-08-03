import assert from "node:assert/strict";
import test from "node:test";

globalThis.XLSX = {
  utils: {
    sheet_to_json(sheet) {
      return sheet.__rows;
    },
  },
};

const { buildStageReportTargets } = await import("../worker/stageReportListParser.js");
const { matchStageReportData } = await import("../worker/stageReportMatching.js");

function workbook(rows) {
  return { SheetNames: ["分母"], Sheets: { 分母: { __rows: rows } } };
}

function chat(overrides = {}) {
  return {
    有效教师邮箱: overrides.email || "teacher@xdf.cn",
    邮箱来源: "群聊发送人邮箱",
    发送人名称: overrides.sender || "老师",
    "群名/好友昵称": overrides.group || "陈一学习群",
    聊天时间: "2026-08-02 09:00:00",
    聊天内容: overrides.content || "",
    来源文件: "聊天.xlsx",
    源聊天行号: 2,
  };
}

test("分母支持常见表头别名，整行重复只保留一次", () => {
  const list = buildStageReportTargets(workbook([
    ["校区", "授课教师", "教师邮箱", "学员号", "学生姓名", "是否申诉", "申诉情况详情"],
    ["越秀", "张老师", "teacher@xdf.cn", "GZ1", "陈一", "", ""],
    ["越秀", "张老师", "teacher@xdf.cn", "GZ1", "陈一", "", ""],
    ["越秀", "张老师", "teacher@xdf.cn", "GZ1", "陈一", "是", "已提交"],
  ]));
  assert.equal(list.targets.length, 2);
  assert.equal(list.counts.整行重复行数, 1);
  assert.equal(list.targets[0].duplicateRows, "3");
  assert.equal(list.targets[1].appeal, true);
  assert.deepEqual(list.columns.slice(0, 5), ["校区", "授课教师", "教师邮箱", "学员号", "学生姓名"]);
});

test("同一教师下，学员后两字命中即可判已发送，不重复审核导出关键词", () => {
  const list = buildStageReportTargets(workbook([
    ["教师姓名", "邮箱", "学号", "学员姓名", "是否发送阶段性报告", "是否申诉", "申诉情况详情"],
    ["张老师", "teacher@xdf.cn", "GZ1", "陈一", "旧值", "是", "待核"],
    ["李老师", "other@xdf.cn", "GZ2", "王二", "", "", ""],
  ]));
  const result = matchStageReportData(list, [
    chat({ group: "其他学习群", content: "家长您好" }),
    chat({ content: "王二家长您好" }),
    chat({ email: "other@xdf.cn", content: "王二家长您好" }),
  ]);
  assert.equal(result.counts.已发送数, 2);
  assert.equal(result.detailRows[0].本次检查结论, "已发送");
  assert.equal(result.detailRows[0]["是否发送阶段性报告"], "旧值");
  assert.equal(result.detailRows[0].命中群名, "陈一学习群");
  assert.equal(result.teacherRows.find((row) => row.教师姓名 === "张老师").申诉数, 1);
});

test("阶段性报告复用打卡口径：学员后两字可命中群名或聊天内容", () => {
  const list = buildStageReportTargets(workbook([
    ["教师姓名", "邮箱", "学号", "学员姓名"],
    ["牛老师", "niu@xdf.cn", "GZ1", "潘胜源"],
  ]));
  const result = matchStageReportData(list, [
    chat({ email: "niu@xdf.cn", group: "潘胜源学习群", content: "阶段性报告总结：胜源本阶段表现良好" }),
  ]);
  assert.equal(result.counts.已发送数, 1);
  assert.equal(result.detailRows[0].本次检查结论, "已发送");
  assert.equal(result.detailRows[0].命中关键词, "胜源（强匹配）");
});

test("钉钉表单无邮箱时，教师姓名去尾号后可兜底匹配聊天发送人", () => {
  const list = buildStageReportTargets(workbook([
    ["教师姓名", "学号", "学生姓名", "是否发送阶段性报告"],
    ["张凝72", "GZ1", "张昊", ""],
  ]));
  const result = matchStageReportData(list, [chat({ sender: "张凝", content: "张昊家长您好" })]);
  assert.equal(result.counts.已发送数, 1);
  assert.equal(result.counts.姓名兜底匹配数, 1);
  assert.equal(result.detailRows[0].教师匹配方式, "姓名兜底匹配");
});

test("没有学员命中或邮箱不一致均不能判已发送", () => {
  const list = buildStageReportTargets(workbook([
    ["教师姓名", "邮箱", "学号", "学员姓名"],
    ["张老师", "teacher@xdf.cn", "GZ1", "陈一"],
  ]));
  const result = matchStageReportData(list, [
    chat({ group: "其他学习群", content: "已发送" }),
    chat({ group: "其他学习群", content: "家长您好" }),
    chat({ email: "another@xdf.cn", content: "阶段性报告：陈一家长您好" }),
  ]);
  assert.equal(result.counts.已发送数, 0);
  assert.equal(result.counts.未发送数, 1);
  assert.equal(result.detailRows[0].本次检查结论, "未发送");
});
