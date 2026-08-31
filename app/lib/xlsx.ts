import JSZip from "jszip";

export const XLSX_LIMITS = { bytes: 25 * 1024 * 1024, entries: 400, inflated: 80 * 1024 * 1024, sheets: 40, rows: 5000, cols: 80, cells: 100000, sharedStrings: 50000, merges: 10000, predictedPages: 20000 } as const;
export type WorkbookCell = { row: number; col: number; value: string; formula?: boolean; style?: number; wrap?: boolean; numberFormat?: string };
export type CellRange = { left: number; top: number; right: number; bottom: number };
export type SheetPageMargins = { left: number; right: number; top: number; bottom: number; header: number; footer: number };
export type SheetPageOrder = "downThenOver" | "overThenDown";
export type SheetGeometryCheck = { axis: "horizontal" | "vertical"; ranges: string[] };
export type SheetRenderPlan = { relationId: string; path: string; sourceHash: string; range: string; bounds: CellRange; titleRows?: CellRange; titleColumns?: CellRange; pageMargins: SheetPageMargins; pageOrder: SheetPageOrder; printableWidthPoints: number; printableHeightPoints: number; geometryChecks: SheetGeometryCheck[]; orientation: "portrait" | "landscape"; tiles: CellRange[]; scalePercent: number; warnings: string[]; predictedPageCount: number; planHash: string };
export type WorkbookSheet = { id: string; path: string; name: string; state: "visible" | "hidden" | "veryHidden"; range: string; printArea?: string; titleRows?: string; titleColumns?: string; pageMargins: SheetPageMargins; pageOrder: SheetPageOrder; rows: number; cols: number; cells: WorkbookCell[]; merges: CellRange[]; rowHeights: Record<number, number>; columnWidths: Record<number, number>; warnings: string[]; renderPlan: SheetRenderPlan };
export type WorkbookAnalysis = { kind: "xlsx"; sourceHash: string; sheets: WorkbookSheet[]; warnings: string[] };
const xml = (value: string) => value.replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n))).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
function decodeWorkbookXml(value: string) {
  let current = value;
  for (let pass = 0; pass < 4; pass += 1) {
    const next = current
      .replace(/&#x([0-9a-f]+);/gi, (_, n) => {
        const code = parseInt(n, 16);
        return Number.isInteger(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : "";
      })
      .replace(/&#(\d+);/g, (_, n) => {
        const code = Number(n);
        return Number.isInteger(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : "";
      })
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&");
    if (next === current) return current;
    current = next;
  }
  return current;
}
const DISALLOWED_WORKBOOK_ENTRY = /^(?:xl\/(?:vbaProject\.bin|macrosheets\/|intlmacrosheets\/|activeX\/|embeddings\/|externalLinks\/|connections\.xml)|customXml\/)/i;
const DISALLOWED_WORKBOOK_CONTENT = /(?:macroEnabled|macrosheet|intlmacrosheet|activeX|oleObject|externalLink|connections?)/i;
const WORKBOOK_EXTERNAL_RELATIONSHIP = /TargetMode\s*=\s*(?:["']\s*External\s*["']|External\b)/i;
async function assertSafeWorkbookPackage(zip: JSZip) {
  const entries = Object.values(zip.files);
  if (entries.some((entry) => DISALLOWED_WORKBOOK_ENTRY.test(entry.name))) {
    throw new Error("Workbook contains macros, embedded objects, external links or connections and cannot be opened for printing.");
  }
  const contentTypes = decodeWorkbookXml(await zip.file("[Content_Types].xml")?.async("text") ?? "");
  if (DISALLOWED_WORKBOOK_CONTENT.test(contentTypes)) {
    throw new Error("Workbook contains active or externally connected content and cannot be opened for printing.");
  }
  for (const relationship of entries.filter((entry) => /\.rels$/i.test(entry.name))) {
    const text = decodeWorkbookXml(await relationship.async("text"));
    if (WORKBOOK_EXTERNAL_RELATIONSHIP.test(text)) {
      throw new Error("Workbook contains an external relationship and cannot be opened for printing.");
    }
  }
}
const attr = (tag: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`, "i").exec(tag)?.[1];
function address(ref: string) { const m = /^([A-Z]+)(\d+)$/i.exec(ref); if (!m) return null; let col = 0; for (const char of m[1].toUpperCase()) col = col * 26 + char.charCodeAt(0) - 64; return { row: Number(m[2]), col }; }
function column(n: number) { let s=""; while(n){s=String.fromCharCode(65+(n-1)%26)+s;n=Math.floor((n-1)/26)} return s; }
function parseRange(ref: string): CellRange | null { const cleaned = ref.replace(/\$/g,""); const m = /^([A-Z]+\d+)(?::([A-Z]+\d+))?$/i.exec(cleaned); const a=m&&address(m[1]), b=m&&address(m[2] ?? m[1]); return a&&b?{left:Math.min(a.col,b.col),top:Math.min(a.row,b.row),right:Math.max(a.col,b.col),bottom:Math.max(a.row,b.row)}:null; }
function withinWorksheetLimits(range: CellRange) { return range.left >= 1 && range.top >= 1 && range.right <= XLSX_LIMITS.cols && range.bottom <= XLSX_LIMITS.rows; }
export function printAreaForSheet(raw: string | undefined, sheetName: string): { range?: CellRange; warning?: string } { if (!raw) return {}; if (raw.includes(",")) return { warning: "Discontiguous print areas are not supported; select one contiguous area in Excel." }; const match = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i.exec(raw.trim()); if (!match) return { warning: "Print area is invalid and was not used." }; const owner = (match[1] ?? match[2] ?? "").replace(/''/g, "'"); if (owner !== sheetName) return { warning: "Print area belongs to another sheet and was not used." }; const range = parseRange(`${match[3]}${match[4]}:${match[5]}${match[6]}`); return range && withinWorksheetLimits(range) ? { range } : { warning: `Print area exceeds the ${XLSX_LIMITS.rows}-row or ${XLSX_LIMITS.cols}-column worksheet analysis limit and was not used.` }; }
function rangeText(r: CellRange) { return `${column(r.left)}${r.top}:${column(r.right)}${r.bottom}`; }
function intersects(a: CellRange, b: CellRange) { return a.left <= b.right && a.right >= b.left && a.top <= b.bottom && a.bottom >= b.top; }
const OOXML_DEFAULT_PAGE_MARGINS: SheetPageMargins = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
const CANONICAL_PAGE_MARGINS: SheetPageMargins = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };

function printTitlesForSheet(raw: string | undefined, sheetName: string, bounds: CellRange): { rows?: CellRange; columns?: CellRange; rowText?: string; columnText?: string; warning?: string } {
  if (!raw) return {};
  const pieces = raw.split(",").map((item) => item.trim()).filter(Boolean);
  let rows: CellRange | undefined;
  let columns: CellRange | undefined;
  for (const piece of pieces) {
    const match = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+|\d+):\$?([A-Z]+|\d+)$/i.exec(piece);
    if (!match || (match[1] ?? match[2] ?? "").replace(/''/g, "'") !== sheetName) {
      return { warning: "Fidelity check failed: Excel print titles are invalid or belong to another worksheet." };
    }
    if (/^\d+$/.test(match[3]) && /^\d+$/.test(match[4])) {
      const top = Number(match[3]), bottom = Number(match[4]);
      if (!top || bottom < top || bottom > XLSX_LIMITS.rows || rows) return { warning: "Fidelity check failed: Excel print-title rows are invalid or exceed the worksheet analysis limit." };
      rows = { left: bounds.left, right: bounds.right, top, bottom };
    } else if (/^[A-Z]+$/i.test(match[3]) && /^[A-Z]+$/i.test(match[4])) {
      const left = address(`${match[3]}1`)?.col ?? 0, right = address(`${match[4]}1`)?.col ?? 0;
      if (!left || right < left || right > XLSX_LIMITS.cols || columns) return { warning: "Fidelity check failed: Excel print-title columns are invalid or exceed the worksheet analysis limit." };
      columns = { left, right, top: bounds.top, bottom: bounds.bottom };
    } else return { warning: "Fidelity check failed: Excel print titles mix row and column coordinates." };
  }
  const unsupported = columns
    ? "Fidelity check failed: repeated print-title columns are not supported for evidential workbook output. Remove them in Excel or print this worksheet separately."
    : rows && (rows.top !== bounds.top || rows.bottom > bounds.bottom)
      ? "Fidelity check failed: repeated print-title rows must be a leading prefix of the approved print range."
      : undefined;
  return {
    rows,
    columns,
    rowText: rows ? `$${rows.top}:$${rows.bottom}` : undefined,
    columnText: columns ? `$${column(columns.left)}:$${column(columns.right)}` : undefined,
    warning: unsupported,
  };
}

function pageGeometry(source: string): { margins: SheetPageMargins; order: SheetPageOrder; warnings: string[] } {
  const warnings: string[] = [];
  const marginTag = /<pageMargins\b([^>]*)\/?\s*>/i.exec(source)?.[1];
  const margins = { ...OOXML_DEFAULT_PAGE_MARGINS };
  if (marginTag) {
    for (const key of Object.keys(margins) as Array<keyof SheetPageMargins>) {
      const parsed = Number(attr(marginTag, key));
      if (!Number.isFinite(parsed) || parsed < 0 || parsed > 10) warnings.push(`Worksheet ${key} page margin is invalid; Exhibit Builder's disclosed canonical A4 margins will be used.`);
      else margins[key] = parsed;
    }
  }
  const setupTag = /<pageSetup\b([^>]*)\/?\s*>/i.exec(source)?.[1] ?? "";
  const savedOrder = attr(setupTag, "pageOrder");
  const order: SheetPageOrder = savedOrder === "overThenDown" ? "overThenDown" : "downThenOver";
  if (savedOrder && savedOrder !== "overThenDown" && savedOrder !== "downThenOver") warnings.push("Worksheet page order is invalid; Exhibit Builder's canonical down-then-over order will be used.");
  const savedComments = attr(setupTag, "comments");
  if (savedComments && savedComments !== "none") warnings.push("Fidelity check failed: worksheet comments are configured to print and are not supported for evidential output.");
  const headerFooter = /<headerFooter\b[^>]*>([\s\S]*?)<\/headerFooter>/i.exec(source)?.[1] ?? "";
  const printableHeaderFooter = xml(headerFooter.replace(/<[^>]+>/g, "")).trim();
  if (printableHeaderFooter) warnings.push("Fidelity check failed: worksheet headers or footers contain printable content; remove them or print this worksheet separately.");
  const nonstandardMargins = (Object.keys(CANONICAL_PAGE_MARGINS) as Array<keyof SheetPageMargins>).some((key) => Math.abs(margins[key] - CANONICAL_PAGE_MARGINS[key]) > 0.0001);
  if (nonstandardMargins) warnings.push("Saved worksheet margins will be normalised to Exhibit Builder's disclosed canonical A4 margins for deterministic output.");
  if (order === "overThenDown") warnings.push("Saved over-then-down page order will be normalised to Exhibit Builder's disclosed down-then-over order.");
  return { margins, order, warnings };
}

function workbookStyleGeometry(stylesSource: string | undefined): { wrapStyles: Set<number>; maxDigitWidthPixels: number; warning?: string } {
  if (!stylesSource) return { wrapStyles: new Set(), maxDigitWidthPixels: 7 };
  const source = stylesSource.replace(/<(\/?)[\w-]+:/g, "<$1");
  const fontsBody = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i.exec(source)?.[1] ?? "";
  const fonts = Array.from(fontsBody.matchAll(/<font\b[^>]*>([\s\S]*?)<\/font>/gi)).map((item) => item[1]);
  const normalStyle = Array.from(source.matchAll(/<cellStyle\b([^>]*)\/?\s*>/gi)).find((item) => attr(item[1], "name") === "Normal");
  const normalXfId = Number(attr(normalStyle?.[1] ?? "", "xfId") ?? 0);
  const styleXfsBody = /<cellStyleXfs\b[^>]*>([\s\S]*?)<\/cellStyleXfs>/i.exec(source)?.[1] ?? "";
  const styleXfs = Array.from(styleXfsBody.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/gi));
  const normalFontId = Number(attr(styleXfs[normalXfId]?.[0] ?? "", "fontId") ?? 0);
  const normalFont = fonts[normalFontId];
  let warning: string | undefined;
  let maxDigitWidthPixels = 7;
  if (normalFont) {
    const name = xml(attr(/<name\b([^>]*)\/?\s*>/i.exec(normalFont)?.[1] ?? "", "val") ?? "");
    const size = Number(attr(/<sz\b([^>]*)\/?\s*>/i.exec(normalFont)?.[1] ?? "", "val"));
    // OOXML column widths are character counts relative to the Normal style's
    // maximum digit width. Excel uses 7 px for Calibri/Aptos 11 and 8 px for
    // Carlito 11. Unknown fonts use a deliberately conservative whole-pixel
    // estimate; native Excel still verifies every resulting geometry check.
    maxDigitWidthPixels = (name === "Calibri" || name === "Aptos") && size === 11
      ? 7
      : name === "Carlito" && size === 11
        ? 8
        : Number.isFinite(size) && size > 0
          ? Math.max(7, Math.ceil(size * 0.72))
          : 8;
    if (!((name === "Calibri" || name === "Aptos") && size === 11)) warning = `Workbook Normal style uses ${name || "an unknown font"} ${Number.isFinite(size) ? `${size}pt` : "at an unknown size"}; native Microsoft Excel dimensions will be verified before output.`;
  }
  const cellXfsBody = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i.exec(source)?.[1] ?? "";
  const wrapStyles = new Set<number>();
  Array.from(cellXfsBody.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/gi)).forEach((item, index) => {
    if (/\bwrapText="(?:1|true)"/i.test(item[0])) wrapStyles.add(index);
  });
  return { wrapStyles, maxDigitWidthPixels, warning };
}
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
function columnWidthPoints(value: number | undefined, defaultWidth: number, maxDigitWidthPixels: number) {
  // Excel's stored width is approximately a character count. This conservative
  // conversion is used only to plan page breaks; native Excel remains the
  // authoritative print engine at the approved 85–100% scale.
  const characterWidth = Math.max(1, value ?? defaultWidth);
  const pixels = Math.floor(((256 * characterWidth + Math.floor(128 / maxDigitWidthPixels)) / 256) * maxDigitWidthPixels);
  return pixels * 0.75;
}

function rowHeightPoints(value: number | undefined, defaultHeight: number) {
  return Math.max(1, Math.min(409, value ?? defaultHeight));
}

function makePlan(
  id: string,
  path: string,
  sourceHash: string,
  bounds: CellRange,
  titleRows: CellRange | undefined,
  titleColumns: CellRange | undefined,
  widths: Record<number, number>,
  heights: Record<number, number>,
  defaultColumnWidth: number,
  defaultRowHeight: number,
  maxDigitWidthPixels: number,
  merges: CellRange[],
  warnings: string[],
): SheetRenderPlan {
  const minimumScalePercent = 85;
  const portraitWidth = 480;
  const landscapeWidth = 725;
  const portraitHeight = 675;
  const landscapeHeight = 430;
  const plannedColumnWidths = Array.from(
    { length: bounds.right - bounds.left + 1 },
    (_, index) => columnWidthPoints(widths[bounds.left + index], defaultColumnWidth, maxDigitWidthPixels),
  );
  const totalWidth = plannedColumnWidths.reduce((sum, width) => sum + width, 0);
  const mergeWidths = merges
    .filter((merge) => intersects(bounds, merge))
    .map((merge) => Array.from(
      { length: merge.right - merge.left + 1 },
      (_, index) => columnWidthPoints(widths[merge.left + index], defaultColumnWidth, maxDigitWidthPixels),
    ).reduce((sum, width) => sum + width, 0));
  const widestUnbreakableWidth = Math.max(...plannedColumnWidths, ...mergeWidths, 0);
  const titleHeight = titleRows
    ? Array.from(
        { length: titleRows.bottom - titleRows.top + 1 },
        (_, index) => rowHeightPoints(heights[titleRows.top + index], defaultRowHeight),
      ).reduce((sum, height) => sum + height, 0)
    : 0;
  const rowHeights = Array.from(
    { length: bounds.bottom - bounds.top + 1 },
    (_, index) => ({ row: bounds.top + index, height: rowHeightPoints(heights[bounds.top + index], defaultRowHeight) }),
  );
  const mergeHeights = merges
    .filter((merge) => intersects(bounds, merge))
    .map((merge) => ({
      merge,
      height: Array.from(
        { length: merge.bottom - merge.top + 1 },
        (_, index) => rowHeightPoints(heights[merge.top + index], defaultRowHeight),
      ).reduce((sum, height) => sum + height, 0),
    }));
  const bodyStartsAfter = titleRows?.bottom ?? bounds.top - 1;
  const tallestFirstPageItem = Math.max(titleHeight, ...rowHeights.map((row) => row.height), ...mergeHeights.map((item) => item.height), 0);
  const tallestRepeatedTitleItem = titleRows
    ? Math.max(
        ...rowHeights.filter((row) => row.row > bodyStartsAfter).map((row) => titleHeight + row.height),
        ...mergeHeights.filter((item) => item.merge.bottom > bodyStartsAfter).map((item) => titleHeight + item.height),
        titleHeight,
      )
    : tallestFirstPageItem;
  const tallestUnbreakableHeight = Math.max(tallestFirstPageItem, tallestRepeatedTitleItem);
  const planOrientation = (orientation: "portrait" | "landscape") => {
    const physicalWidthBudget = orientation === "landscape" ? landscapeWidth : portraitWidth;
    const physicalHeightBudget = orientation === "landscape" ? landscapeHeight : portraitHeight;
    const widthScaleLimit = widestUnbreakableWidth > physicalWidthBudget
      ? Math.floor((physicalWidthBudget / widestUnbreakableWidth) * 100)
      : 100;
    const heightScaleLimit = tallestUnbreakableHeight > physicalHeightBudget
      ? Math.floor((physicalHeightBudget / tallestUnbreakableHeight) * 100)
      : 100;
    const requiredScalePercent = Math.min(widthScaleLimit, heightScaleLimit);
    const scalePercent = Math.max(minimumScalePercent, Math.min(100, requiredScalePercent));
    const compliant = requiredScalePercent >= minimumScalePercent;
    const plannedWarnings = [...warnings];
    const widthBudget = physicalWidthBudget * (100 / scalePercent);
    if (!compliant) {
      plannedWarnings.push(`Fidelity check failed: unbreakable worksheet content needs ${Math.max(1, requiredScalePercent)}% scale to fit one ${orientation} A4 printable page; the readable minimum is ${minimumScalePercent}%.`);
      if (titleHeight > physicalHeightBudget * (100 / minimumScalePercent)) {
        plannedWarnings.push(`Fidelity check failed: repeated title rows need more than one ${orientation} A4 printable page at the readable ${minimumScalePercent}% minimum scale.`);
      }
    } else if (scalePercent < 100) {
      plannedWarnings.push(`Microsoft Excel will print this worksheet at ${scalePercent}% scale so unbreakable rows, columns and merged cells stay within one ${orientation} A4 page.`);
    }

    const horizontal: Array<{ left: number; right: number }> = [];
    let left = bounds.left;
    while (left <= bounds.right) {
      let right = left - 1;
      let used = 0;
      while (right < bounds.right) {
        const candidate = right + 1;
        const candidateWidth = columnWidthPoints(widths[candidate], defaultColumnWidth, maxDigitWidthPixels);
        if (right >= left && used + candidateWidth > widthBudget) break;
        used += candidateWidth;
        right = candidate;
        if (candidateWidth > widthBudget) {
          plannedWarnings.push(`Fidelity check failed: column ${column(candidate)} is wider than one ${orientation} A4 printable page at the readable ${minimumScalePercent}% minimum scale.`);
          break;
        }
      }
      const crossing = merges.find((merge) => merge.left <= right && merge.right > right && merge.right <= bounds.right);
      if (crossing) {
        const mergedWidth = Array.from(
          { length: crossing.right - crossing.left + 1 },
          (_, index) => columnWidthPoints(widths[crossing.left + index], defaultColumnWidth, maxDigitWidthPixels),
        ).reduce((sum, width) => sum + width, 0);
        if (mergedWidth <= widthBudget) right = crossing.left > left ? crossing.left - 1 : crossing.right;
        else plannedWarnings.push(`Fidelity check failed: merged cell ${rangeText(crossing)} is wider than one ${orientation} A4 printable page at the readable ${minimumScalePercent}% minimum scale.`);
      }
      horizontal.push({ left, right: Math.max(left, right) });
      left = Math.max(left + 1, right + 1);
    }

    const vertical: Array<{ top: number; bottom: number }> = [];
    let top = bounds.top;
    while (top <= bounds.bottom) {
      let bottom = top - 1;
      let used = 0;
      let heightBudget = physicalHeightBudget * (100 / scalePercent);
      if (top > bounds.top) heightBudget -= titleHeight;
      while (bottom < bounds.bottom) {
        const nextHeight = rowHeightPoints(heights[bottom + 1], defaultRowHeight);
        if (bottom >= top && used + nextHeight > heightBudget) break;
        used += nextHeight;
        bottom += 1;
      }
      let crossing = merges.find((merge) => merge.top <= bottom && merge.bottom > bottom && merge.bottom <= bounds.bottom);
      while (crossing) {
        const activeMerge = crossing;
        const mergedHeight = Array.from(
          { length: activeMerge.bottom - activeMerge.top + 1 },
          (_, index) => rowHeightPoints(heights[activeMerge.top + index], defaultRowHeight),
        ).reduce((sum, height) => sum + height, 0);
        if (mergedHeight <= heightBudget && activeMerge.top > top) bottom = activeMerge.top - 1;
        else {
          if (mergedHeight > heightBudget) plannedWarnings.push(`Fidelity check failed: merged cell ${rangeText(activeMerge)} is taller than one ${orientation} A4 printable page at the readable ${minimumScalePercent}% minimum scale.`);
          bottom = activeMerge.bottom;
        }
        crossing = merges.find((merge) => merge.top <= bottom && merge.bottom > bottom && merge.bottom <= bounds.bottom);
      }
      vertical.push({ top, bottom: Math.max(top, bottom) });
      top = Math.max(top + 1, bottom + 1);
    }
    const predictedPageCount = horizontal.length * vertical.length;
    if (predictedPageCount > XLSX_LIMITS.predictedPages) plannedWarnings.push(`Fidelity check failed: worksheet pagination predicts ${predictedPageCount} pages, above the ${XLSX_LIMITS.predictedPages}-page native Excel safety limit.`);
    const tiles = predictedPageCount <= XLSX_LIMITS.predictedPages
      ? horizontal.flatMap((columns) => vertical.map((rows) => ({ ...columns, ...rows })))
      : [{ ...bounds }];
    return { orientation, requiredScalePercent, scalePercent, compliant: compliant && predictedPageCount <= XLSX_LIMITS.predictedPages, plannedWarnings, tiles, predictedPageCount };
  };

  const widthHeuristic: "portrait" | "landscape" = totalWidth > portraitWidth ? "landscape" : "portrait";
  const candidates = (["portrait", "landscape"] as const).map(planOrientation);
  const compliantCandidates = candidates.filter((candidate) => candidate.compliant);
  const pool = compliantCandidates.length ? compliantCandidates : candidates;
  pool.sort((left, right) => {
    if (compliantCandidates.length && left.scalePercent !== right.scalePercent) return right.scalePercent - left.scalePercent;
    if (!compliantCandidates.length && left.requiredScalePercent !== right.requiredScalePercent) return right.requiredScalePercent - left.requiredScalePercent;
    if (left.predictedPageCount !== right.predictedPageCount) return left.predictedPageCount - right.predictedPageCount;
    return left.orientation === widthHeuristic ? -1 : right.orientation === widthHeuristic ? 1 : 0;
  });
  const selected = pool[0];
  const { orientation, scalePercent, plannedWarnings, tiles, predictedPageCount } = selected;
  const printableWidthPoints = orientation === "landscape" ? landscapeWidth : portraitWidth;
  const printableHeightPoints = orientation === "landscape" ? landscapeHeight : portraitHeight;
  const titleRangeText = titleRows ? rangeText(titleRows) : undefined;
  const geometryChecks: SheetGeometryCheck[] = [
    ...Array.from({ length: bounds.right - bounds.left + 1 }, (_, index) => ({ axis: "horizontal" as const, ranges: [rangeText({ left: bounds.left + index, right: bounds.left + index, top: bounds.top, bottom: bounds.top })] })),
    ...merges.filter((merge) => intersects(bounds, merge)).map((merge) => ({ axis: "horizontal" as const, ranges: [rangeText(merge)] })),
    ...Array.from({ length: bounds.bottom - bounds.top + 1 }, (_, index) => ({ axis: "vertical" as const, ranges: [rangeText({ left: bounds.left, right: bounds.left, top: bounds.top + index, bottom: bounds.top + index })] })),
    ...merges.filter((merge) => intersects(bounds, merge)).map((merge) => ({ axis: "vertical" as const, ranges: [rangeText(merge)] })),
    ...(titleRangeText ? [
      ...Array.from({ length: Math.max(0, bounds.bottom - titleRows!.bottom) }, (_, index) => ({ axis: "vertical" as const, ranges: [titleRangeText, rangeText({ left: bounds.left, right: bounds.left, top: titleRows!.bottom + index + 1, bottom: titleRows!.bottom + index + 1 })] })),
      ...merges.filter((merge) => intersects(bounds, merge) && merge.bottom > titleRows!.bottom).map((merge) => ({ axis: "vertical" as const, ranges: [titleRangeText, rangeText(merge)] })),
    ] : []),
  ];
  const planHash = stableHash(JSON.stringify({ id, path, sourceHash, bounds, titleRows, titleColumns, orientation, tiles, plannedWarnings, widths, heights, defaultColumnWidth, defaultRowHeight, maxDigitWidthPixels, merges, scalePercent, geometryChecks }));
  return {
    relationId: id,
    path,
    sourceHash,
    range: rangeText(bounds),
    bounds,
    titleRows,
    titleColumns,
    pageMargins: { ...CANONICAL_PAGE_MARGINS },
    pageOrder: "downThenOver",
    printableWidthPoints,
    printableHeightPoints,
    geometryChecks,
    orientation,
    tiles,
    scalePercent,
    warnings: plannedWarnings,
    predictedPageCount,
    planHash,
  };
}
function textOf(fragment: string) { return xml(Array.from(fragment.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)).map((m) => m[1]).join("")); }
/** Parses OOXML cell data only. Formulas, macros, and external content are never executed. */
export async function analyseXlsx(file: File): Promise<WorkbookAnalysis> {
  if (!/\.xlsx$/i.test(file.name)) throw new Error("Only .xlsx workbooks are supported.");
  if (file.size > XLSX_LIMITS.bytes) throw new Error("Workbook exceeds the 25MB safety limit.");
  let zip: JSZip; try { zip = await JSZip.loadAsync(await file.arrayBuffer(), { createFolders: false }); } catch { throw new Error("Workbook is malformed or not a safe OOXML archive."); }
  const sourceHash = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", await file.arrayBuffer()))).map((part) => part.toString(16).padStart(2, "0")).join("");
const entries = Object.values(zip.files); if (entries.length > XLSX_LIMITS.entries || entries.some((e) => e.name.includes(".."))) throw new Error("Workbook archive has unsafe paths or too many entries.");
  const inflated = entries.reduce((n, e) => n + ((e as any)._data?.uncompressedSize ?? 0), 0); if (inflated > XLSX_LIMITS.inflated || inflated / Math.max(file.size, 1) > 120) throw new Error("Workbook expands beyond the safety compression limit.");
  if (zip.file("EncryptedPackage")) throw new Error("Encrypted workbooks are not supported.");
  await assertSafeWorkbookPackage(zip);
  const workbook = zip.file("xl/workbook.xml"); if (!workbook) throw new Error("Workbook has no workbook definition."); const wb = (await workbook.async("text")).replace(/<(\/?)[\w-]+:/g, "<$1"); if (/externalLink|<externalReferences/i.test(wb)) throw new Error("Workbooks with external links are not supported.");
  const stringsEntry = zip.file("xl/sharedStrings.xml"); const strings = stringsEntry ? Array.from((await stringsEntry.async("text")).matchAll(/<si[^>]*>([\s\S]*?)<\/si>/g)).map((m) => textOf(m[1])) : []; if (strings.length > XLSX_LIMITS.sharedStrings) throw new Error("Workbook has too many shared strings.");
  const stylesEntry = zip.file("xl/styles.xml");
  const styleGeometry = workbookStyleGeometry(stylesEntry ? await stylesEntry.async("text") : undefined);
  const relEntry = zip.file("xl/_rels/workbook.xml.rels"); if (!relEntry) throw new Error("Workbook sheet relationships are missing."); const rels = (await relEntry.async("text")).replace(/<(\/?)[\w-]+:/g, "<$1"); const links = new Map(Array.from(rels.matchAll(/<Relationship\b[^>]*\/>/g)).map((m) => [attr(m[0], "Id"), attr(m[0], "Target")]));
  const tags = Array.from(wb.matchAll(/<sheet\b[^>]*\/?>/g)); if (!tags.length || tags.length > XLSX_LIMITS.sheets) throw new Error("Workbook has an unsupported number of sheets."); const sheets: WorkbookSheet[] = []; const paths = new Set<string>(); const defined = Array.from(wb.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g));
  for (const [sheetIndex, tagMatch] of tags.entries()) {
    const tag = tagMatch[0];
    const name = xml(attr(tag, "name") ?? "Unnamed sheet");
    const id = attr(tag, "r:id") ?? attr(tag, "id");
    const rawTitles = defined.find((item) => attr(item[1], "name") === "_xlnm.Print_Titles" && Number(attr(item[1], "localSheetId")) === sheetIndex)?.[2];
    const rawTarget = id ? links.get(id) : undefined;
    const target = rawTarget?.replace(/^\/?xl\//, "");
    if (!id || !target || !/^worksheets\/[A-Za-z0-9._-]+\.xml$/.test(target) || paths.has(target)) throw new Error(`Workbook sheet relationship for ${name} is unsafe or missing.`);
    paths.add(target);
    const rawState = attr(tag, "state") ?? "visible";
    const state = rawState === "veryHidden" ? "veryHidden" : rawState === "hidden" ? "hidden" : "visible" as const;
    const path = `xl/${target}`;
    const entry = zip.file(path);
    if (!entry) throw new Error(`Workbook sheet ${name} is missing.`);
    const source = (await entry.async("text")).replace(/<(\/?)[\w-]+:/g, "<$1");
    const sheetFormat = /<sheetFormatPr\b([^>]*)\/?/i.exec(source)?.[1] ?? "";
    const parsedDefaultColumnWidth = Number(attr(sheetFormat, "defaultColWidth") ?? 8.43);
    const parsedDefaultRowHeight = Number(attr(sheetFormat, "defaultRowHeight") ?? 15);
    const defaultColumnWidth = Number.isFinite(parsedDefaultColumnWidth) && parsedDefaultColumnWidth > 0 ? parsedDefaultColumnWidth : 8.43;
    const defaultRowHeight = Number.isFinite(parsedDefaultRowHeight) && parsedDefaultRowHeight > 0 ? parsedDefaultRowHeight : 15;
    const savedDimension = /<dimension[^>]*ref="([^"]+)"/.exec(source)?.[1];
    const detectedDimension = savedDimension ? parseRange(savedDimension) : undefined;
    if (savedDimension && (!detectedDimension || !withinWorksheetLimits(detectedDimension))) throw new Error(`Sheet ${name} dimension exceeds the ${XLSX_LIMITS.rows}-row or ${XLSX_LIMITS.cols}-column analysis limit.`);
    const cells: WorkbookCell[] = [];
    const rowHeights: Record<number, number> = {};
    const columnWidths: Record<number, number> = {};
    for (const rowMatch of source.matchAll(/<row\b([^>]*)/g)) {
      const row = Number(attr(rowMatch[1], "r") ?? 0), height = Number(attr(rowMatch[1], "ht") ?? 0);
      if (!Number.isInteger(row) || row < 1 || row > XLSX_LIMITS.rows) throw new Error(`Sheet ${name} contains a row outside the worksheet analysis limit.`);
      if (row && height > 0) rowHeights[row] = height;
    }
    for (const columnMatch of source.matchAll(/<col\b([^>]*)\/?/g)) {
      const low = Number(attr(columnMatch[1], "min") ?? 0), high = Number(attr(columnMatch[1], "max") ?? low), width = Number(attr(columnMatch[1], "width") ?? 0);
      if (!Number.isInteger(low) || !Number.isInteger(high) || low < 1 || high < low || high > XLSX_LIMITS.cols) throw new Error(`Sheet ${name} contains a column definition outside the worksheet analysis limit.`);
      for (let columnIndex = low; columnIndex <= high && width > 0; columnIndex += 1) columnWidths[columnIndex] = width;
    }
    let rows = 0, cols = 0;
    for (const match of source.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const pos = address(attr(match[1], "r") ?? "");
      if (!pos) throw new Error(`Sheet ${name} contains an invalid cell address.`);
      if (pos.row > XLSX_LIMITS.rows || pos.col > XLSX_LIMITS.cols) throw new Error(`Sheet ${name} exceeds the ${XLSX_LIMITS.rows} row or ${XLSX_LIMITS.cols} column limit.`);
      if (cells.length >= XLSX_LIMITS.cells) throw new Error(`Sheet ${name} exceeds the cell limit.`);
      const type = attr(match[1], "t"), body = match[2], formula = /<f\b/i.test(body), hasValue = /<v[^>]*>[\s\S]*?<\/v>/.test(body);
      if (formula && !hasValue) throw new Error(`Sheet ${name} contains a formula without a cached displayed value.`);
      const raw = /<v[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1] ?? (type === "inlineStr" ? textOf(body) : "");
      const value = type === "s" ? strings[Number(raw)] ?? "" : xml(raw);
      if (!value && !formula) continue;
      const style = Number(attr(match[1], "s"));
      cells.push({ ...pos, value, formula, style: Number.isInteger(style) ? style : undefined, wrap: Number.isInteger(style) ? styleGeometry.wrapStyles.has(style) : false });
      rows = Math.max(rows, pos.row); cols = Math.max(cols, pos.col);
    }
    const rawPrint = defined.find((item) => attr(item[1], "name") === "_xlnm.Print_Area" && Number(attr(item[1], "localSheetId")) === sheetIndex)?.[2];
    const print = printAreaForSheet(rawPrint, name);
    if (rawPrint && /exceeds the .*analysis limit/i.test(print.warning ?? "")) throw new Error(`Sheet ${name} print area exceeds the worksheet analysis limit.`);
    const merges: CellRange[] = [];
    for (const match of source.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)) {
      if (merges.length >= XLSX_LIMITS.merges) throw new Error(`Sheet ${name} exceeds the ${XLSX_LIMITS.merges}-merge analysis limit.`);
      const merge = parseRange(match[1]);
      if (!merge || !withinWorksheetLimits(merge)) throw new Error(`Sheet ${name} contains a merged range outside the worksheet analysis limit.`);
      merges.push(merge);
    }
    const detected = detectedDimension ?? { left: 1, top: 1, right: Math.max(1, cols), bottom: Math.max(1, rows) };
    const bounds = print.range ?? expandAcrossMergedCells(detected, merges);
    const clippedMerge = print.range ? merges.find((merge) => intersects(print.range!, merge) && (merge.left < print.range!.left || merge.top < print.range!.top || merge.right > print.range!.right || merge.bottom > print.range!.bottom)) : undefined;
    const titles = printTitlesForSheet(rawTitles, name, bounds);
    const geometry = pageGeometry(source);
    const relationshipsPath = `xl/worksheets/_rels/${target.split("/").at(-1)}.rels`;
    const sheetRelationships = await zip.file(relationshipsPath)?.async("text") ?? "";
    const warnings = [
      print.warning ?? "",
      titles.warning ?? "",
      ...geometry.warnings,
      styleGeometry.warning ?? "",
      cells.some((cell) => cell.wrap && rowHeights[cell.row] === undefined) ? "Worksheet contains wrapped text with automatic row heights; native Microsoft Excel dimensions will be verified before output." : "",
      clippedMerge ? `Fidelity check failed: Excel's print area ${rangeText(print.range!)} cuts through merged cell ${rangeText(clippedMerge)}.` : "",
      /<(?:legacyDrawing(?:HF)?|picture)\b/i.test(source) || /comments|threadedComment|notes/i.test(sheetRelationships) ? "Fidelity check failed: worksheet comments, notes, or header/footer pictures are not supported for evidential output." : "",
      /<printOptions\b[^>]*\bheadings="(?:1|true)"/i.test(source) ? "Fidelity check failed: printed row or column headings are not supported for evidential output." : "",
      /<drawing\b/i.test(source) ? "Fidelity check failed: worksheet drawings or charts are not supported because their printable anchors cannot yet be verified inside the approved range." : "",
      /conditionalFormatting/i.test(source) ? "Conditional formatting is preserved by the native Microsoft Excel print engine." : "",
      /<dataValidations/i.test(source) ? "Data-validation rules do not change the saved cell values printed by Microsoft Excel." : "",
    ].filter(Boolean);
    const renderPlan = makePlan(id, path, sourceHash, bounds, titles.rows, titles.columns, columnWidths, rowHeights, defaultColumnWidth, defaultRowHeight, styleGeometry.maxDigitWidthPixels, merges, warnings);
    sheets.push({ id, path, name, state, range: rangeText(bounds), printArea: print.range ? rangeText(print.range) : undefined, titleRows: titles.rowText, titleColumns: titles.columnText, pageMargins: geometry.margins, pageOrder: geometry.order, rows, cols, cells, merges, rowHeights, columnWidths, warnings: renderPlan.warnings, renderPlan });
  }
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
