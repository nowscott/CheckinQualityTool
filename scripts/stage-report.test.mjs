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
const { findTrainingLeadGroups } = await import("../worker/stageReportGroups.js");
const {
  buildAssistantHierarchy,
  buildResearchGroupHierarchy,
  isStageReportAppealed,
} = await import("../worker/stageReportHierarchy.js");

const dualMetrics = [
  {
    totalColumn: "阶段应发",
    sentColumn: "阶段已发",
    rateColumn: "阶段发送率",
    unsentColumn: "阶段未发",
    appealColumn: "阶段申诉",
  },
  {
    totalColumn: "窗口应发",
    sentColumn: "窗口已发",
    rateColumn: "窗口发送率",
    unsentColumn: "窗口未发",
  },
];

function workbook(rows) {
  return { SheetNames: ["分母"], Sheets: { 分母: { __rows: rows } } };
}

function workbookWithSheets(sheets) {
  return {
    SheetNames: Object.keys(sheets),
    Sheets: Object.fromEntries(Object.entries(sheets).map(([name, rows]) => [name, { __rows: rows }])),
  };
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

test("多 Sheet 文件按阶段性报告主列选择分母，不误选窗口期发送明细", () => {
  const list = buildStageReportTargets(workbookWithSheets({
    "窗口期报告发送明细": [
      ["教师姓名", "学生姓名", "学号", "师训组长", "助理主管", "是否发送窗口期报告"],
      ["张老师", "陈一", "GZ1", "张老师", "李老师", "否"],
    ],
    "教师维度明细（截至8_7）": [
      ["教师姓名", "师训组长", "助理主管", "教研组", "阶段性报告需发送"],
      ["张老师", "张老师", "李老师", "高中双语", "1"],
    ],
    "非窗口期暑期在读学员阶段性报告发送明细": [
      ["教师姓名", "学生姓名", "学号", "暑假最后一节课时间", "师训组长", "助理主管", "教研组", "是否发送阶段性报告", "是否需要发送"],
      ["张老师", "陈一", "GZ1", "2026-08-05 18:30-20:30", "张老师", "李老师", "高中双语", "否", "是"],
    ],
    "组长维度汇总": [
      ["师训组长", "助理主管", "教研组", "窗口期报告需发送数", "窗口期报告已发送数"],
      ["张老师", "李老师", "高中双语", "1", "0"],
    ],
  }));
  assert.equal(list.sheetName, "非窗口期暑期在读学员阶段性报告发送明细");
  assert.equal(list.targets.length, 1);
  assert.equal(list.targets[0].original.教研组, "高中双语");
  assert.equal(list.targets[0].original["是否发送阶段性报告"], "否");
  assert.deepEqual(list.trainingLeadGroups.get("张"), ["高中双语"]);
});

test("组长不在阶段性报告分母时，使用同文件组长维度汇总的自身教研组", () => {
  const list = buildStageReportTargets(workbookWithSheets({
    "阶段性报告分母": [
      ["教师姓名", "学生姓名", "学号", "师训组长", "教研组", "是否发送阶段性报告"],
      ["李老师", "陈一", "GZ1", "张组长", "初中双语", "否"],
    ],
    "组长维度汇总": [
      ["师训组长", "教研组", "非窗口期暑期在读阶段性报告需发送数"],
      ["张组长", "高中双语", "5"],
    ],
  }));
  assert.deepEqual(list.trainingLeadGroups.get("张组长"), ["高中双语"]);
});

test("完整教师明细中的组长本人教研组优先于其负责范围汇总", () => {
  const list = buildStageReportTargets(workbookWithSheets({
    "阶段性报告分母": [
      ["教师姓名", "学生姓名", "学号", "师训组长", "教研组", "是否发送阶段性报告"],
      ["李老师", "陈一", "GZ1", "张组长", "初中双语", "否"],
    ],
    "教师维度明细": [
      ["教师姓名", "师训组长", "教研组", "阶段性报告需发送"],
      ["张组长", "张组长", "初中双语", "1"],
      ["李老师", "张组长", "高中双语", "1"],
    ],
    "组长维度汇总": [
      ["师训组长", "教研组", "非窗口期暑期在读阶段性报告需发送数"],
      ["张组长", "高中双语", "5"],
    ],
  }));
  assert.deepEqual(list.trainingLeadGroups.get("张组长"), ["初中双语"]);
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

test("师训组长维度只展示组长本人所属教研组，不展开其负责教师的其他教研组", () => {
  const teacherRows = [
    { 教师姓名: "张组长", 教研组: "高中双语" },
    { 教师姓名: "李老师", 教研组: "初中双语" },
    { 教师姓名: "王老师", 教研组: "高中双语" },
  ];
  assert.deepEqual(findTrainingLeadGroups(teacherRows, "张组长"), ["高中双语"]);
  assert.deepEqual(findTrainingLeadGroups(teacherRows, "未授课组长"), []);
});

test("组长本人姓名末尾数字不同于其他教师时，不串用其他人的教研组", () => {
  const teacherRows = [
    { 教师姓名: "黄磊44", 教研组: "实验P" },
    { 教师姓名: "黄磊23", 教研组: "初中益智" },
  ];
  assert.deepEqual(findTrainingLeadGroups(teacherRows, "黄磊23"), ["初中益智"]);
});

test("双口径助理主管汇总保留本人教研组、小计、项目小计和总计", () => {
  const rows = [
    { 教师姓名: "张主管", 助理主管: "张主管1", 教研组: "高中双语", 阶段应发: 2, 阶段已发: 1, 阶段申诉: 1, 窗口应发: 4, 窗口已发: 3 },
    { 教师姓名: "李教师", 助理主管: "张主管1", 教研组: "初中益智", 阶段应发: 3, 阶段已发: 3, 阶段申诉: 2, 窗口应发: 1, 窗口已发: 0 },
    { 教师姓名: "王教师", 助理主管: "王主管2", 教研组: "初中益智", 阶段应发: 5, 阶段已发: 4, 阶段申诉: 0, 窗口应发: 5, 窗口已发: 5 },
  ];
  const table = buildAssistantHierarchy(rows, dualMetrics);
  const manager = table.rows.find((row) => row["师训主管/助理主管"] === "张主管");
  const total = table.rows.find((row) => row.__rowType === "grandTotal");
  assert.equal(manager.教研组, "高中双语");
  assert.equal(manager.阶段应发, 5);
  assert.equal(manager.阶段申诉, 3);
  assert.equal(total.阶段应发, 10);
  assert.equal(total.阶段已发, 8);
  assert.equal(total.阶段发送率, 0.8);
  assert.equal(total.窗口应发, 10);
  assert.equal(total.窗口已发, 8);
  assert.ok(table.rows.some((row) => row.__rowType === "groupTotal"));
  assert.ok(table.rows.some((row) => row.__rowType === "projectTotal"));
  assert.ok(table.mergeCells.some((range) => /^A\d+:B\d+$/u.test(range)));
});

test("教研组汇总按项目排序并去重负责人尾号", () => {
  const rows = [
    { 教师姓名: "甲", 助理主管: "李主管12", 教研组: "初中益智", 阶段应发: 2, 阶段已发: 1, 阶段申诉: 0, 窗口应发: 3, 窗口已发: 2 },
    { 教师姓名: "乙", 助理主管: "李主管12", 教研组: "初中益智", 阶段应发: 1, 阶段已发: 1, 阶段申诉: 1, 窗口应发: 1, 窗口已发: 1 },
    { 教师姓名: "丙", 助理主管: "王主管3", 教研组: "高中双语", 阶段应发: 4, 阶段已发: 0, 阶段申诉: 0, 窗口应发: 4, 窗口已发: 0 },
  ];
  const table = buildResearchGroupHierarchy(rows, dualMetrics);
  const group = table.rows.find((row) => row.教研组 === "初中益智" && row.__rowType === "detail");
  const projectRows = table.rows.filter((row) => row.__rowType === "projectTotal");
  assert.equal(group.负责人, "李主管");
  assert.equal(group.阶段应发, 3);
  assert.equal(group.阶段申诉, 1);
  assert.deepEqual(projectRows.map((row) => row.教研组), ["双语项目", "益智项目"]);
});

test("教研组美化汇总可隐藏负责人且不跨指标合并", () => {
  const table = buildResearchGroupHierarchy([
    { 教师姓名: "甲", 助理主管: "李主管12", 教研组: "初中益智", 阶段应发: 2, 阶段已发: 1, 阶段申诉: 0, 窗口应发: 3, 窗口已发: 2 },
  ], dualMetrics, false);
  assert.ok(table.rows.every((row) => !("负责人" in row)));
  assert.deepEqual(table.mergeCells, []);
  const total = table.rows.find((row) => row.__rowType === "grandTotal");
  assert.equal(total.阶段应发, 2);
  assert.equal(total.窗口应发, 3);
});

test("申诉判定与零应发发送率口径保持一致", () => {
  assert.equal(isStageReportAppealed({ 是否申诉: "否", 申诉情况详情: "" }), false);
  assert.equal(isStageReportAppealed({ 是否申诉: "短期冲刺课/三次课内学员" }), true);
  assert.equal(isStageReportAppealed({ 是否申诉: "", 申诉情况详情: "已提交说明" }), true);
  const table = buildResearchGroupHierarchy([
    { 教师姓名: "甲", 助理主管: "主管1", 教研组: "高中双语", 阶段应发: 0, 阶段已发: 0, 阶段申诉: 1, 窗口应发: 0, 窗口已发: 0 },
  ], dualMetrics);
  const total = table.rows.find((row) => row.__rowType === "grandTotal");
  assert.equal(total.阶段发送率, 1);
  assert.equal(total.窗口发送率, 1);
});
