import assert from "node:assert/strict";
import test from "node:test";
import {
  assertHistoryPayload,
  buildInspectionItemsInsertQueries,
  insertBatch,
  voidBatch,
} from "../server/inspection/shared.js";

const sha = "a".repeat(64);

function rawItem(index) {
  return {
    position: index + 1,
    teacherName: `教师${index + 1}`,
    teacherEmail: `teacher${index + 1}@xdf.cn`,
    studentName: `学员${index + 1}`,
    studentId: `student-${index + 1}`,
    courseId: `course-${index + 1}`,
    lessonStart: "2026-09-07 10:00",
    lessonEnd: "2026-09-07 11:00",
    submittedValue: index % 2 ? "是" : "否",
    productGroup: "产品",
    campus: "校区",
    projectGroup: "项目组",
    selectionReason: "测试",
  };
}

function rawPayload(items) {
  return {
    batch: {
      businessWeekStart: "2026-09-07",
      businessWeekEnd: "2026-09-13",
      sourceName: "课堂反馈.xlsx",
      sourceSha256: sha,
      rosterName: "在职明细.xlsx",
      rosterSha256: sha,
      rosterSnapshotDate: "2026-09-15",
      sampleLimit: items.length,
      eligibleCount: items.length,
      selectedCount: items.length,
      teacherCount: items.length,
      unsubmittedCount: 0,
      ruleVersion: "inspection-v2",
      attempt: 1,
      batchKind: "formal",
    },
    items,
  };
}

function batchRow(id) {
  return {
    id,
    business_week_start: "2026-09-07",
    business_week_end: "2026-09-13",
    source_name: "课堂反馈.xlsx",
    source_sha256: sha,
    roster_name: "在职明细.xlsx",
    roster_sha256: sha,
    roster_snapshot_date: "2026-09-15",
    sample_limit: 1000,
    eligible_count: 1000,
    selected_count: 1000,
    teacher_count: 1000,
    unsubmitted_count: 0,
    rule_version: "inspection-v2",
    attempt: 1,
    batch_kind: "formal",
    status: "active",
    created_at: "2026-09-23T00:00:00.000Z",
    voided_at: null,
  };
}

function fakeSql({ fail = false } = {}) {
  const calls = [];
  const transactions = [];
  const sql = (strings, ...values) => {
    const text = strings.raw.join("?");
    calls.push({ type: "template", text, values });
    return { type: "template", text, values };
  };
  sql.query = (query, params) => {
    calls.push({ type: "query", query, params });
    return { type: "query", query, params };
  };
  sql.transaction = async (queries) => {
    transactions.push(queries);
    if (fail) throw new Error("模拟事务失败");
    return [...queries.slice(0, -1).map(() => []), [batchRow("batch-1")]];
  };
  return { sql, calls, transactions };
}

test("1000 条明细使用一次批量 INSERT，并保留顺序", async () => {
  const items = Array.from({ length: 1000 }, (_, index) => rawItem(index));
  const payload = assertHistoryPayload(rawPayload(items));
  const { sql, transactions } = fakeSql();
  const result = await insertBatch(sql, payload, "batch-1");

  assert.equal(transactions.length, 1);
  assert.equal(transactions[0].length, 3);
  const itemQuery = transactions[0][1];
  assert.equal(itemQuery.type, "query");
  assert.equal(itemQuery.params.length, 14_000);
  assert.equal((itemQuery.query.match(/\(\$\d+/g) || []).length, 1000);
  assert.equal(result.items.length, 1000);
  assert.equal(result.items[0].position, 1);
  assert.equal(result.items[999].position, 1000);
});

test("空明细跳过 inspection_items INSERT，但仍在事务内保存批次", async () => {
  const payload = assertHistoryPayload(rawPayload([]));
  const { sql, transactions } = fakeSql();
  const result = await insertBatch(sql, payload, "batch-1");

  assert.equal(transactions.length, 1);
  assert.equal(transactions[0].length, 2);
  assert.equal(result.items.length, 0);
});

test("超过参数安全范围时分块，但仍只提交一个事务", () => {
  const items = Array.from({ length: 2001 }, (_, index) => rawItem(index));
  const queries = buildInspectionItemsInsertQueries("batch-1", assertHistoryPayload(rawPayload(items)).items);

  assert.equal(queries.length, 2);
  assert.equal(queries[0].params.length, 28_000);
  assert.equal(queries[1].params.length, 14);
});

test("替换批次的旧批次作废和新批次写入共用事务", async () => {
  const payload = assertHistoryPayload(rawPayload([rawItem(0)]));
  const { sql, transactions } = fakeSql();
  await insertBatch(sql, payload, "batch-2", "batch-1");

  assert.equal(transactions.length, 1);
  assert.equal(transactions[0].length, 4);
  assert.match(transactions[0][0].text, /UPDATE inspection_batches/);
});

test("事务失败直接抛错，不执行事务后的读取", async () => {
  const payload = assertHistoryPayload(rawPayload([rawItem(0)]));
  const { sql, transactions } = fakeSql({ fail: true });

  await assert.rejects(() => insertBatch(sql, payload, "batch-1"), /模拟事务失败/);
  assert.equal(transactions.length, 1);
});

function fakeVoidSql(results) {
  const calls = [];
  const sql = (strings, ...values) => {
    calls.push({ text: strings.raw.join("?"), values });
    return Promise.resolve(results.shift() || []);
  };
  return { sql, calls };
}

test("作废只转换仍生效的批次、保留课程明细并原子写审计", async () => {
  const row = { ...batchRow("batch-void"), status: "voided", voided_at: "2026-09-23T00:00:00.000Z" };
  const { sql, calls } = fakeVoidSql([[row]]);
  const result = await voidBatch(sql, "batch-void", "admin-user");

  assert.equal(result.id, "batch-void");
  assert.equal(result.status, "voided");
  assert.equal(calls.length, 1);
  assert.match(calls[0].text, /UPDATE inspection_batches/);
  assert.match(calls[0].text, /status = 'active'/);
  assert.match(calls[0].text, /INSERT INTO inspection_audit_log/);
  assert.match(calls[0].text, /inspection_batch_voided/);
  assert.doesNotMatch(calls[0].text, /DELETE FROM inspection_items/);
  assert.ok(calls[0].values.includes("batch-void"));
  assert.ok(calls[0].values.includes("admin-user"));
});

test("并发状态已变化时返回冲突，不重复写审计", async () => {
  const { sql, calls } = fakeVoidSql([[], [{ id: "batch-void", status: "voided" }]]);

  await assert.rejects(() => voidBatch(sql, "batch-void", "admin-user"), /已不是生效状态/);
  assert.equal(calls.length, 2);
  assert.doesNotMatch(calls[1].text, /INSERT INTO inspection_audit_log/);
});

test("作废不存在的批次返回空结果", async () => {
  const { sql, calls } = fakeVoidSql([[], []]);

  assert.equal(await voidBatch(sql, "missing-batch", "admin-user"), null);
  assert.equal(calls.length, 2);
});
