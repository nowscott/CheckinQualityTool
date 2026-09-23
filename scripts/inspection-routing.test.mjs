import assert from "node:assert/strict";
import test from "node:test";
import { route } from "../api/inspection.js";

function request(path, method = "POST") {
  return {
    method,
    headers: { cookie: "inspection_auth=example-session", origin: "https://example.test" },
    body: { example: true },
    query: { path, keep: "value" },
  };
}

test("batch void routing preserves method, headers, body, and existing query fields", () => {
  const original = request("batches/batch-123/void");
  const selected = route(original);

  assert.ok(selected);
  assert.equal(selected.request.method, "POST");
  assert.equal(selected.request.headers, original.headers);
  assert.equal(selected.request.body, original.body);
  assert.equal(selected.request.query.id, "batch-123");
  assert.equal(selected.request.query.keep, "value");
  assert.equal(selected.request.query.path, "batches/batch-123/void");
});

test("history detail routing also retains authentication headers", () => {
  const original = request("history/batch-456", "GET");
  const selected = route(original);

  assert.ok(selected);
  assert.equal(selected.request.method, "GET");
  assert.equal(selected.request.headers, original.headers);
  assert.equal(selected.request.query.id, "batch-456");
});
