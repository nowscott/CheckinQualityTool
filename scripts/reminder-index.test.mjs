import assert from "node:assert/strict";
import test from "node:test";

const { buildStudentHitIndex } = await import("../worker/reminderMatching.js");

test("提醒索引在 8,000 关键词和 90,000 聊天规模完成", { timeout: 60_000 }, () => {
  const keywords = Array.from({ length: 8_000 }, (_, index) => `学${String.fromCharCode(0x4e00 + index)}`);
  const chats = Array.from({ length: 90_000 }, (_, index) => ({
    index,
    chat: {},
    group: index === 45_000 ? `学习群 ${keywords[123]}` : `学习群 ${index}`,
    content: "课程反馈",
  }));
  const started = performance.now();
  const result = buildStudentHitIndex(chats, keywords);
  const elapsed = performance.now() - started;
  assert.equal(result.size, keywords.length);
  assert.equal(result.get(keywords[123]).length, 1);
  assert.ok(elapsed < 60_000, `索引耗时 ${elapsed.toFixed(0)}ms`);
});
