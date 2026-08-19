import { analyseXlsx } from "./xlsx.ts";
self.onmessage = async (event: MessageEvent<{ file: File }>) => { try { self.postMessage({ ok: true, analysis: await analyseXlsx(event.data.file) }); } catch (error) { self.postMessage({ ok: false, error: error instanceof Error ? error.message : "Workbook analysis failed." }); } };
