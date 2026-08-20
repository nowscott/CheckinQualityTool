import assert from "node:assert/strict";
import test from "node:test";

const { matchData } = await import("../worker/matching.js");
const { buildWhitelist } = await import("../worker/whitelist.js");

function target() {
  return {
    教师姓名: "张老师",
    教师邮箱: "teacher@xdf.cn",
    学员姓名: "卢思妍",
    原始学员姓名: "卢思妍",
    姓名清洗说明: "",
    学员号: "GZ2401573428",
    上课日期: "2026-08-14",
    上课开始: "16:00",
    上课结束: "18:00",
    校区: "",
    项目组: "",
    科目: "",
    源名单行号: 2,
    该周课次数: 1,
  };
}

function chat(content) {
  return {
    有效教师邮箱: "teacher@xdf.cn",
    邮箱来源: "群聊发送人邮箱",
    发送人名称: "张老师",
    "群名/好友昵称": "",
    聊天时间: "2026-08-14 18:30:00",
    聊天内容: content,
    源聊天行号: 2,
  };
}

test("打卡白名单原名或别名任一命中即可发送", () => {
  const whitelist = buildWhitelist([
    "学员号,学员姓名,处理方式,匹配别名,说明",
    "GZ2401573428,卢思妍,别名,周炳燊,系统登记名与实际学员名不一致",
  ].join("\n"));

  const aliasResult = matchData([target()], [chat("炳燊家长您好，以下是课程反馈。")], false, "第三周", whitelist);
  assert.equal(aliasResult.finalRows[0].发送情况, "已发送");
  assert.equal(aliasResult.finalRows[0].匹配结论, "别名匹配");
  assert.equal(aliasResult.finalRows[0].命中关键词, "炳燊");

  const originalResult = matchData([target()], [chat("卢思妍家长您好，以下是课程反馈。")], false, "第三周", whitelist);
  assert.equal(originalResult.finalRows[0].发送情况, "已发送");
  assert.equal(originalResult.finalRows[0].匹配结论, "强匹配");
  assert.equal(originalResult.finalRows[0].命中关键词, "思妍");
});

test("打卡白名单学员姓名也按后两字作为候选", () => {
  const whitelist = buildWhitelist([
    "学员号,学员姓名,处理方式,匹配别名,说明",
    "GZ2401573428,卢思妍,别名,周炳燊,系统登记名与实际学员名不一致",
  ].join("\n"));
  const row = { ...target(), 学员姓名: "历史登记名", 原始学员姓名: "历史登记名" };
  const result = matchData([row], [chat("思妍家长您好，以下是课程反馈。")], false, "第三周", whitelist);
  assert.equal(result.finalRows[0].发送情况, "已发送");
  assert.equal(result.finalRows[0].匹配结论, "强匹配");
  assert.equal(result.finalRows[0].命中关键词, "思妍");
});
