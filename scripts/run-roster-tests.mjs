import fs from "node:fs/promises";
import { execFileSync } from "node:child_process";
import path from "node:path";
import os from "node:os";

const output = await fs.mkdtemp(path.join(os.tmpdir(), "checkin-roster-tests-"));
try {
  execFileSync("node_modules/.bin/tsc", ["--ignoreConfig", "--target", "ES2022", "--module", "ESNext", "--moduleResolution", "Bundler", "--outDir", output, "--skipLibCheck", "--strict", "api/roster.ts", "src/worker/inspectionProcess.ts", "src/worker/sheetjs.d.ts"], { stdio: "inherit" });
  async function fixImports(dir) {
    for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
      const file = path.join(dir, entry.name);
      if (entry.isDirectory()) await fixImports(file);
      else if (entry.name.endsWith(".js")) {
        const source = await fs.readFile(file, "utf8");
        await fs.writeFile(file, source.replace(/from "(\.\.?\/[^".]+)"/gu, 'from "$1.js"'));
      }
    }
  }
  await fixImports(output);
  await fs.writeFile(path.join(output, "package.json"), JSON.stringify({ type: "module" }));
  await fs.mkdir(path.join(output, "scripts"));
  await fs.mkdir(path.join(output, "node_modules"));
  await fs.symlink(path.resolve("node_modules/fflate"), path.join(output, "node_modules/fflate"));
  await fs.copyFile("scripts/roster-service.test.mjs", path.join(output, "scripts/roster-service.test.mjs"));
  execFileSync(process.execPath, ["--test", path.join(output, "scripts/roster-service.test.mjs")], { stdio: "inherit" });
} finally {
  await fs.rm(output, { recursive: true, force: true });
}
