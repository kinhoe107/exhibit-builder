import JSZip from "jszip";

export const XLSX_LIMITS = { bytes: 25 * 1024 * 1024, entries: 400, inflated: 80 * 1024 * 1024, sheets: 40, rows: 5000, cols: 80, cells: 100000, sharedStrings: 50000 } as const;
export type WorkbookCell = { row: number; col: number; value: string; formula?: boolean; style?: number; wrap?: boolean; numberFormat?: string };
export type CellRange = { left: number; top: number; right: number; bottom: number };
export type SheetRenderPlan = { relationId: string; path: string; sourceHash: string; range: string; bounds: CellRange; titleRows?: CellRange; orientation: "portrait" | "landscape"; tiles: CellRange[]; fontSize: number; warnings: string[]; predictedPageCount: number; planHash: string };
export type WorkbookSheet = { id: string; path: string; name: string; state: "visible" | "hidden" | "veryHidden"; range: string; printArea?: string; titleRows?: string; rows: number; cols: number; cells: WorkbookCell[]; merges: CellRange[]; rowHeights: Record<number, number>; columnWidths: Record<number, number>; warnings: string[]; renderPlan: SheetRenderPlan };
export type WorkbookAnalysis = { kind: "xlsx"; sourceHash: string; sheets: WorkbookSheet[]; warnings: string[] };
const xml = (value: string) => value.replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n))).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const attr = (tag: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`, "i").exec(tag)?.[1];
function address(ref: string) { const m = /^([A-Z]+)(\d+)$/i.exec(ref); if (!m) return null; let col = 0; for (const char of m[1].toUpperCase()) col = col * 26 + char.charCodeAt(0) - 64; return { row: Number(m[2]), col }; }
function column(n: number) { let s=""; while(n){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26)} return s; }
function parseRange(ref: string): CellRange | null { const m = /^([A-Z]+\d+):([A-Z]+\d+)$/i.exec(ref.replace(/\$/g,"")); const a=m&&address(m[1]), b=m&&address(m[2]); return a&&b?{left:Math.min(a.col,b.col),top:Math.min(a.row,b.row),right:Math.max(a.col,b.col),bottom:Math.max(a.row,b.row)}:null; }
export function printAreaForSheet(raw: string | undefined, sheetName: string): { range?: CellRange; warning?: string } { if (!raw) return {}; if (raw.includes(",")) return { warning: "Discontiguous print areas are not supported; select one contiguous area in Excel." }; const match = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i.exec(raw.trim()); if (!match) return { warning: "Print area is invalid and was not used." }; const owner = (match[1] ?? match[2] ?? "").replace(/''/g, "'"); if (owner !== sheetName) return { warning: "Print area belongs to another sheet and was not used." }; const range = parseRange(`${match[3]}${match[4]}:${match[5]}${match[6]}`); return range ? { range } : { warning: "Print area is invalid and was not used." }; }
function rangeText(r: CellRange) { return `${column(r.left)}${r.top}:${column(r.right)}${r.bottom}`; }
function intersects(a: CellRange, b: CellRange) { return a.left <= b.right && a.right >= b.left && a.top <= b.bottom && a.bottom >= b.top; }
/** Excel renders a merged cell over its complete rectangle even though only its top-left cell stores the value. */
function expandAcrossMergedCells(initial: CellRange, merges: CellRange[]) {
  let expanded = { ...initial };
  let changed = true;
  while (changed) {
    changed = false;
    for (const merge of merges) {
      if (!intersects(expanded, merge)) continue;
      const next = { left: Math.min(expanded.left, merge.left), top: Math.min(expanded.top, merge.top), right: Math.max(expanded.right, merge.right), bottom: Math.max(expanded.bottom, merge.bottom) };
      if (rangeText(next) !== rangeText(expanded)) { expanded = next; changed = true; }
    }
  }
  return expanded;
}
function stableHash(value:string) { let h=2166136261; for(let i=0;i<value.length;i++)h=Math.imul(h^value.charCodeAt(i),16777619); return (`00000000${(h>>>0).toString(16)}`).slice(-8); }
function makePlan(id:string,path:string,sourceHash:string,bounds:CellRange,titleRows:CellRange|undefined,widths:Record<number,number>,heights:Record<number,number>,warnings:string[]):SheetRenderPlan { const total=Array.from({length:bounds.right-bounds.left+1},(_,i)=>widths[bounds.left+i]??8.43).reduce((a,b)=>a+b,0), orientation=total>70?"landscape":"portrait", cols=orientation==="landscape"?9:7, tiles:CellRange[]=[]; const titleHeight=titleRows?Array.from({length:titleRows.bottom-titleRows.top+1},(_,i)=>Math.max(15,Math.min(42,heights[titleRows.top+i]??15))).reduce((a,b)=>a+b,0):0; for(let l=bounds.left;l<=bounds.right;l+=cols) { let top=bounds.top; while(top<=bounds.bottom) { let bottom=top-1, used=0, budget=orientation==="landscape"?430:675; if(top>bounds.top) budget-=titleHeight; while(bottom<bounds.bottom) { const next=Math.max(15,Math.min(42,heights[bottom+1]??15)); if(bottom>=top&&used+next>budget) break; used+=next; bottom++; } tiles.push({left:l,top,right:Math.min(bounds.right,l+cols-1),bottom:Math.max(top,bottom)}); top=Math.max(top+1,bottom+1); } } const planHash=stableHash(JSON.stringify({id,path,sourceHash,bounds,titleRows,orientation,tiles,warnings,heights})); return {relationId:id,path,sourceHash,range:rangeText(bounds),bounds,titleRows,orientation,tiles,fontSize:7,warnings,predictedPageCount:tiles.length,planHash}; }
function textOf(fragment: string) { return xml(Array.from(fragment.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)).map((m) => m[1]).join("")); }
/** Parses OOXML cell data only. Formulas, macros, and external content are never executed. */
export async function analyseXlsx(file: File): Promise<WorkbookAnalysis> {
  if (!/\.xlsx$/i.test(file.name)) throw new Error("Only .xlsx workbooks are supported.");
  if (file.size > XLSX_LIMITS.bytes) throw new Error("Workbook exceeds the 25MB safety limit.");
  let zip: JSZip; try { zip = await JSZip.loadAsync(await file.arrayBuffer(), { createFolders: false }); } catch { throw new Error("Workbook is malformed or not a safe OOXML archive."); }
  const sourceHash = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", await file.arrayBuffer()))).map((part) => part.toString(16).padStart(2, "0")).join("");
const entries = Object.values(zip.files); if (entries.length > XLSX_LIMITS.entries || entries.some((e) => e.name.includes(".."))) throw new Error("Workbook archive has unsafe paths or too many entries.");
  const inflated = entries.reduce((n, e) => n + ((e as any)._data?.uncompressedSize ?? 0), 0); if (inflated > XLSX_LIMITS.inflated || inflated / Math.max(file.size, 1) > 120) throw new Error("Workbook expands beyond the safety compression limit.");
  if (zip.file("EncryptedPackage") || zip.file("xl/vbaProject.bin")) throw new Error("Encrypted or macro-enabled workbooks are not supported.");
  const workbook = zip.file("xl/workbook.xml"); if (!workbook) throw new Error("Workbook has no workbook definition."); const wb = (await workbook.async("text")).replace(/<(\/?)[\w-]+:/g, "<$1"); if (/externalLink|<externalReferences/i.test(wb)) throw new Error("Workbooks with external links are not supported.");
  const stringsEntry = zip.file("xl/sharedStrings.xml"); const strings = stringsEntry ? Array.from((await stringsEntry.async("text")).matchAll(/<si[^>]*>([\s\S]*?)<\/si>/g)).map((m) => textOf(m[1])) : []; if (strings.length > XLSX_LIMITS.sharedStrings) throw new Error("Workbook has too many shared strings.");
  const relEntry = zip.file("xl/_rels/workbook.xml.rels"); if (!relEntry) throw new Error("Workbook sheet relationships are missing."); const rels = (await relEntry.async("text")).replace(/<(\/?)[\w-]+:/g, "<$1"); const links = new Map(Array.from(rels.matchAll(/<Relationship\b[^>]*\/>/g)).map((m) => [attr(m[0], "Id"), attr(m[0], "Target")]));
  const tags = Array.from(wb.matchAll(/<sheet\b[^>]*\/?>/g)); if (!tags.length || tags.length > XLSX_LIMITS.sheets) throw new Error("Workbook has an unsupported number of sheets."); const sheets: WorkbookSheet[] = []; const paths = new Set<string>(); const defined = Array.from(wb.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g));
  for (const [sheetIndex, tagMatch] of tags.entries()) { const tag = tagMatch[0], name = xml(attr(tag, "name") ?? "Unnamed sheet"), id = attr(tag, "r:id") ?? attr(tag, "id"); const printName = defined.find((item) => attr(item[1], "name") === "_xlnm.Print_Area" && Number(attr(item[1], "localSheetId")) === sheetIndex)?.[2]?.replace(/['$]/g, ""); const titles = defined.find((item) => attr(item[1], "name") === "_xlnm.Print_Titles" && Number(attr(item[1], "localSheetId")) === sheetIndex)?.[2]?.replace(/['$]/g, ""); const rawTarget = id ? links.get(id) : undefined; const target = rawTarget?.replace(/^\/?xl\//, ""); if (!id || !target || !/^worksheets\/[A-Za-z0-9._-]+\.xml$/.test(target) || paths.has(target)) throw new Error(`Workbook sheet relationship for ${name} is unsafe or missing.`); paths.add(target); const rawState = attr(tag, "state") ?? "visible", state = rawState === "veryHidden" ? "veryHidden" : rawState === "hidden" ? "hidden" : "visible" as const; const path = `xl/${target}`; const entry = zip.file(path); if (!entry) throw new Error(`Workbook sheet ${name} is missing.`); const source = (await entry.async("text")).replace(/<(\/?)[\w-]+:/g, "<$1"); const cells: WorkbookCell[] = []; const rowHeights: Record<number, number> = {}, columnWidths: Record<number, number> = {}; for (const m of source.matchAll(/<row\b([^>]*)/g)) { const r=Number(attr(m[1],"r")??0), h=Number(attr(m[1],"ht")??0); if(r&&h>0) rowHeights[r]=h; } for (const m of source.matchAll(/<col\b([^>]*)\/?/g)) { const lo=Number(attr(m[1],"min")??0), hi=Number(attr(m[1],"max")??lo), w=Number(attr(m[1],"width")??0); for(let c=lo;c<=hi&&w>0;c++) columnWidths[c]=w; } let rows = 0, cols = 0;
    for (const match of source.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) { const pos = address(attr(match[1], "r") ?? ""); if (!pos) throw new Error(`Sheet ${name} contains an invalid cell address.`); if (pos.row > XLSX_LIMITS.rows || pos.col > XLSX_LIMITS.cols) throw new Error(`Sheet ${name} exceeds the ${XLSX_LIMITS.rows} row or ${XLSX_LIMITS.cols} column limit.`); if (cells.length >= XLSX_LIMITS.cells) throw new Error(`Sheet ${name} exceeds the cell limit.`); const type = attr(match[1], "t"), body = match[2], formula = /<f\b/i.test(body), hasValue = /<v[^>]*>[\s\S]*?<\/v>/.test(body); if (formula && !hasValue) throw new Error(`Sheet ${name} contains a formula without a cached displayed value.`); const raw = /<v[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1] ?? (type === "inlineStr" ? textOf(body) : ""), value = type === "s" ? strings[Number(raw)] ?? "" : xml(raw); if (!value && !formula) continue; cells.push({ ...pos, value, formula }); rows = Math.max(rows, pos.row); cols = Math.max(cols, pos.col); }
    const range = /<dimension[^>]*ref="([^"]+)"/.exec(source)?.[1] ?? "A1"; const rawPrint = defined.find((item) => attr(item[1], "name") === "_xlnm.Print_Area" && Number(attr(item[1], "localSheetId")) === sheetIndex)?.[2]; const print = printAreaForSheet(rawPrint, name); const merges = Array.from(source.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)).map((m) => parseRange(m[1])).filter((m): m is CellRange => !!m); const detected = parseRange(range) ?? { left: 1, top: 1, right: Math.max(1, cols), bottom: Math.max(1, rows) }; const bounds = print.range ?? expandAcrossMergedCells(detected, merges); const clippedMerge = print.range ? merges.find((merge) => intersects(print.range!, merge) && (merge.left < print.range!.left || merge.top < print.range!.top || merge.right > print.range!.right || merge.bottom > print.range!.bottom)) : undefined; const title = titles?.match(/(\d+):(\d+)/); const titleRows = title ? { left: bounds.left, right: bounds.right, top: Number(title[1]), bottom: Number(title[2]) } : undefined; const warnings = [print.warning ?? "", clippedMerge ? `Fidelity check failed: Excel's print area ${rangeText(print.range!)} cuts through merged cell ${rangeText(clippedMerge)}.` : "", /(<drawing|<legacyDrawing)/i.test(source) ? "This worksheet contains drawings; Microsoft Excel will include printable drawings in its native PDF output." : "", /conditionalFormatting/i.test(source) ? "Conditional formatting is preserved by the native Microsoft Excel print engine." : "", /<dataValidations/i.test(source) ? "Data-validation rules do not change the saved cell values printed by Microsoft Excel." : ""].filter(Boolean); sheets.push({ id, path, name, state, range: rangeText(bounds), printArea: print.range ? rangeText(print.range) : undefined, titleRows: titles, rows, cols, cells, merges, rowHeights, columnWidths, warnings, renderPlan: makePlan(id, path, sourceHash, bounds, titleRows, columnWidths, rowHeights, warnings) }); }
  return { kind: "xlsx", sourceHash, sheets, warnings: ["Formula results use cached displayed values; formulas are not calculated."] };
}

/** Browser callers isolate OOXML decompression and XML parsing from the UI thread. */
export function analyseXlsxInWorker(file: File): Promise<WorkbookAnalysis> {
  const timeoutError = () => new Error("Workbook analysis timed out after 20 seconds.");
  if (typeof Worker === "undefined") {
    return Promise.race([
      analyseXlsx(file),
      new Promise<WorkbookAnalysis>((_, reject) => {
        setTimeout(() => reject(timeoutError()), 20_000);
      }),
    ]);
  }
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("./xlsx-worker.ts", import.meta.url), { type: "module" });
    const finish = () => worker.terminate();
    const timeout = window.setTimeout(() => { finish(); reject(new Error("Workbook analysis timed out after 20 seconds.")); }, 20_000);
    const close = () => { window.clearTimeout(timeout); finish(); };
    worker.onmessage = (event: MessageEvent<{ ok: boolean; analysis?: WorkbookAnalysis; error?: string }>) => { close(); event.data.ok && event.data.analysis ? resolve(event.data.analysis) : reject(new Error(event.data.error ?? "Workbook analysis failed.")); };
    worker.onerror = () => { close(); reject(new Error("Workbook analysis worker failed.")); };
    worker.postMessage({ file });
  });
}
