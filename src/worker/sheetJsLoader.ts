let sheetJsPromise: Promise<void> | undefined;

function sheetJsGlobal() {
  return globalThis as typeof globalThis & { XLSX?: SheetJsGlobal };
}

export async function ensureSheetJs() {
  if (sheetJsGlobal().XLSX) return;
  sheetJsPromise ||= fetch("/vendor/xlsx.full.min.js", { cache: "force-cache" }).then(async (response) => {
    if (!response.ok) throw new Error("Excel 解析依赖加载失败，请刷新页面后重试。");
    const source = await response.text();
    (0, eval)(source);
    if (!sheetJsGlobal().XLSX) throw new Error("Excel 解析依赖初始化失败，请刷新页面后重试。");
  });
  await sheetJsPromise;
}
