import assert from "node:assert/strict";
import test from "node:test";
import { assertSameOrigin, hashPassword, verifyPassword } from "../shared.js";

test("scrypt password hashes verify and use a fresh salt", async () => {
  const first = await hashPassword("a-secure-test-password");
  const second = await hashPassword("a-secure-test-password");
  assert.notEqual(first, second);
  assert.equal(await verifyPassword("a-secure-test-password", first), true);
  assert.equal(await verifyPassword("wrong-password", first), false);
});

test("password policy rejects empty passwords and allows the requested short password", async () => {
  await assert.rejects(() => hashPassword(""), /不能为空/);
  const hash = await hashPassword("admin");
  assert.equal(await verifyPassword("admin", hash), true);
});

test("same-origin validation accepts same host and rejects foreign origin", () => {
  assert.doesNotThrow(() => assertSameOrigin({ headers: { origin: "https://example.test", host: "example.test" } }));
  assert.throws(() => assertSameOrigin({ headers: { origin: "https://attacker.test", host: "example.test" } }), /请求来源不受信任/);
});
