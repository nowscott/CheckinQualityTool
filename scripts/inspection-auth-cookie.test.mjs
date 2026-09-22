import assert from "node:assert/strict";
import test from "node:test";
import {
  INSPECTION_AUTH_REVALIDATION_MS,
  isInspectionAuthCookieStale,
  readInspectionAuthCookie,
} from "./inspectionAuthCookie.js";

const user = {
  id: "user-1",
  username: "operator",
  displayName: "操作员",
  role: "operator",
  lastLoginAt: "2026-09-22T10:00:00.000Z",
};

function encodeCookie(value) {
  return Buffer.from(JSON.stringify(value)).toString("base64url");
}

test("reads the public identity from the browser cookie", () => {
  const cookie = encodeCookie({ user, issuedAt: 1000 });
  assert.deepEqual(readInspectionAuthCookie(`other=value; inspection_auth_hint=${cookie}`), { user, issuedAt: 1000 });
});

test("cookie becomes stale at the ten-minute boundary", () => {
  const cookie = { user, issuedAt: 1000 };
  assert.equal(isInspectionAuthCookieStale(cookie, 1000 + INSPECTION_AUTH_REVALIDATION_MS - 1), false);
  assert.equal(isInspectionAuthCookieStale(cookie, 1000 + INSPECTION_AUTH_REVALIDATION_MS), true);
});

test("malformed or forged public cookies are ignored by the client parser", () => {
  assert.equal(readInspectionAuthCookie("inspection_auth_hint=not-json"), null);
  assert.equal(readInspectionAuthCookie(`inspection_auth_hint=${encodeCookie({ user: { role: "admin" }, issuedAt: 1000 })}`), null);
});
