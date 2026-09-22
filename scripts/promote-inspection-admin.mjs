import { neon } from "@neondatabase/serverless";

function normalizeUsername(value) {
  const username = String(value || "").trim().toLowerCase();
  if (!/^[a-z0-9][a-z0-9._+@-]{1,118}$/u.test(username)) throw new Error("请提供合法的用户名。用法：npm run auth:promote -- 用户名");
  return username;
}

const username = normalizeUsername(process.argv[2]);
const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) throw new Error("请先提供 DATABASE_URL，例如使用 node --env-file=.env.local 运行。");

const sql = neon(databaseUrl);
const rows = await sql`SELECT id, username, display_name, role, is_active FROM inspection_users WHERE username = ${username} LIMIT 1`;
if (!rows.length) throw new Error(`没有找到用户名为 ${username} 的账号，请先在页面完成注册。`);

const user = rows[0];
if (user.is_active === false) throw new Error(`账号 ${username} 已停用，请先恢复账号后再提升管理员。`);
if (user.role === "admin") {
  console.log(`账号 ${username} 已经是管理员，无需重复操作。`);
  process.exit(0);
}

await sql`UPDATE inspection_users SET role = 'admin', session_version = session_version + 1, updated_at = now() WHERE id = ${user.id}`;
await sql`UPDATE inspection_sessions SET revoked_at = now() WHERE user_id = ${user.id} AND revoked_at IS NULL`;
await sql`INSERT INTO inspection_audit_log (id, actor_user_id, target_user_id, action, metadata) VALUES (${crypto.randomUUID()}, NULL, ${user.id}, 'user_promoted_to_admin', ${JSON.stringify({ username, previousRole: user.role, source: 'controlled-script' })})`;

console.log(`账号 ${username} 已提升为管理员。该账号的旧会话已失效，请重新登录。`);
