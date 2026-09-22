import { randomBytes, scrypt } from "node:crypto";
import { promisify } from "node:util";
import { createInterface } from "node:readline/promises";
import { neon } from "@neondatabase/serverless";

const scryptAsync = promisify(scrypt);
const minPasswordLength = 1;

function validateUsername(value) {
  const username = String(value || "").trim().toLowerCase();
  if (!/^[a-z0-9][a-z0-9._+@-]{1,118}$/u.test(username)) throw new Error("用户名应为 2～120 位字母、数字或常用符号。");
  return username;
}

async function hashPassword(value) {
  const password = String(value || "");
  if (Array.from(password).length < minPasswordLength) throw new Error("密码不能为空。");
  const salt = randomBytes(16);
  const derived = await scryptAsync(password, salt, 64, { N: 32768, r: 8, p: 1, maxmem: 64 * 1024 * 1024 });
  return `scrypt$v1$32768$8$1$${salt.toString("base64url")}$${Buffer.from(derived).toString("base64url")}`;
}

function hiddenQuestion(prompt) {
  const input = process.stdin;
  return new Promise((resolve, reject) => {
    let value = "";
    const previousRaw = input.isRaw;
    const cleanup = () => {
      input.removeListener("data", onData);
      if (input.setRawMode) input.setRawMode(Boolean(previousRaw));
      input.pause();
    };
    const onData = (chunk) => {
      for (const character of String(chunk)) {
        if (character === "\u0003") {
          cleanup();
          reject(new Error("操作已取消。"));
          return;
        }
        if (character === "\r" || character === "\n") {
          cleanup();
          process.stdout.write("\n");
          resolve(value);
          return;
        }
        if (character === "\u007f" || character === "\b") {
          if (value) value = value.slice(0, -1);
          continue;
        }
        value += character;
      }
    };
    process.stdout.write(prompt);
    if (input.setRawMode) input.setRawMode(true);
    input.resume();
    input.on("data", onData);
  });
}

const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) throw new Error("请先提供 DATABASE_URL，例如使用 node --env-file=.env.local 运行。");

const sql = neon(databaseUrl);
await sql`CREATE TABLE IF NOT EXISTS inspection_users (
  id text PRIMARY KEY,
  username text NOT NULL UNIQUE,
  display_name text NOT NULL,
  password_hash text NOT NULL,
  role text NOT NULL DEFAULT 'viewer' CHECK (role IN ('admin', 'operator', 'viewer')),
  is_active boolean NOT NULL DEFAULT true,
  session_version integer NOT NULL DEFAULT 1,
  last_login_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
)`;
const existing = await sql`SELECT COUNT(*)::int AS count FROM inspection_users`;
if (Number(existing[0]?.count || 0) > 0) throw new Error("已经存在用户，初始化脚本不会覆盖或新增首个管理员。");

const readline = createInterface({ input: process.stdin, output: process.stdout });
try {
  const username = validateUsername(await readline.question("管理员用户名: "));
  const displayName = String(await readline.question("显示名称（可留空）: ")).trim() || username;
  readline.close();
  const password = await hiddenQuestion("管理员密码（不能为空）: ");
  const confirmation = await hiddenQuestion("再次输入密码: ");
  if (password !== confirmation) throw new Error("两次密码不一致。");
  const passwordHash = await hashPassword(password);
  await sql`INSERT INTO inspection_users (id, username, display_name, password_hash, role) VALUES (${crypto.randomUUID()}, ${username}, ${displayName.slice(0, 120)}, ${passwordHash}, 'admin')`;
  console.log(`管理员 ${username} 已创建。请使用该账号登录抽检系统。`);
} finally {
  readline.close();
}
