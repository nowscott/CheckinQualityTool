export async function loadInspectionWorkerUrl(developmentUrl?: string) {
  if (developmentUrl) return developmentUrl;
  const response = await fetch("/data/inspection-worker.json", { cache: "no-store" });
  if (!response.ok) throw new Error("抽检程序加载失败，请稍后重试。");
  const manifest = await response.json() as { url?: unknown };
  if (typeof manifest.url !== "string" || !/^\/assets\/inspectionEntry-[\w-]+\.js$/u.test(manifest.url)) {
    throw new Error("抽检程序地址无效，请稍后重试。");
  }
  return manifest.url;
}
