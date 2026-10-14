import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import packageJson from "./package.json" with { type: "json" };

function inspectionWorkerManifest(): Plugin {
  return {
    name: "inspection-worker-manifest",
    apply: "build",
    generateBundle(_options, bundle) {
      const worker = Object.values(bundle).find((file) => /^assets\/inspectionEntry-[\w-]+\.js$/u.test(file.fileName));
      if (!worker) this.error("Missing inspection worker asset.");
      this.emitFile({
        type: "asset",
        fileName: "data/inspection-worker.json",
        source: JSON.stringify({ url: `/${worker.fileName}` }),
      });
    },
  };
}

export default defineConfig({
  plugins: [react(), inspectionWorkerManifest()],
  define: {
    __APP_VERSION__: JSON.stringify(packageJson.version),
  },
  worker: {
    format: "iife",
  },
});
