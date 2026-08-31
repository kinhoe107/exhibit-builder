import JSZip, { type JSZipObject } from "jszip";

type WordRelationship = { target: string; type: string };

const A4_WIDTH_MM = 210;
const A4_HEIGHT_MM = 297;
const MAX_TEMPLATE_BYTES = 25 * 1024 * 1024;
const MAX_ARCHIVE_ENTRIES = 200;
const MAX_INFLATED_BYTES = 50 * 1024 * 1024;
const MAX_IMAGE_BYTES = 12 * 1024 * 1024;
const TEMPLATE_OPERATION_TIMEOUT_MS = 20_000;

type SafeDocxBudget = {
  counted: Set<string>;
  inflatedBytes: number;
  imageBytes: number;
};

const safeDocxBudgets = new WeakMap<JSZip, SafeDocxBudget>();

async function withTemplateTimeout<T>(operation: Promise<T>, label: string): Promise<T> {
  let timeout: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      operation,
      new Promise<T>((_resolve, reject) => {
        timeout = setTimeout(() => reject(new Error(`The Word template took too long to read ${label}.`)), TEMPLATE_OPERATION_TIMEOUT_MS);
      }),
    ]);
  } finally {
    if (timeout) clearTimeout(timeout);
  }
}

function recordActualExpansion(zip: JSZip, entry: JSZipObject, byteLength: number) {
  const budget = safeDocxBudgets.get(zip);
  if (!budget || budget.counted.has(entry.name)) return;
  budget.counted.add(entry.name);
  budget.inflatedBytes += byteLength;
  if (/^word\/media\//i.test(entry.name)) budget.imageBytes += byteLength;
  if (byteLength > MAX_INFLATED_BYTES || budget.inflatedBytes > MAX_INFLATED_BYTES) {
    throw new Error("The Word template expands beyond the 50 MB offline safety limit.");
  }
  if (budget.imageBytes > MAX_IMAGE_BYTES) {
    throw new Error("The Word template contains more than 12 MB of embedded images.");
  }
}

async function safeEntryBytes(zip: JSZip, entry: JSZipObject) {
  const bytes = await withTemplateTimeout(entry.async("uint8array"), entry.name);
  recordActualExpansion(zip, entry, bytes.byteLength);
  return bytes;
}

async function safeEntryText(zip: JSZip, entry: JSZipObject) {
  return new TextDecoder().decode(await safeEntryBytes(zip, entry));
}

function extensionOf(name: string) {
  return name.split(".").pop()?.toLowerCase() ?? "";
}

function attribute(source: string, name: string) {
  const escaped = name.replace(/[.*+?^$()|[\]\\]/g, "\\$&");
  const match = source.match(
    new RegExp("(?:^|\\s)(?:[A-Za-z_][\\w.-]*:)?" + escaped + "\\s*=\\s*[\"']([^\"']*)[\"']", "i"),
  );
  return match?.[1] ?? "";
}

function decodeXml(value: string) {
  return value
    .replace(/&#x([0-9a-f]+);/gi, (_match, code: string) => String.fromCodePoint(Number.parseInt(code, 16)))
    .replace(/&#(\d+);/g, (_match, code: string) => String.fromCodePoint(Number(code)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function findTagStart(xml: string, tag: string, from = 0) {
  const match = new RegExp("<(?:(?:[A-Za-z_][\\w.-]*):)?" + tag + "(?=\\s|/|>)", "g");
  match.lastIndex = from;
  const found = match.exec(xml);
  return found?.index ?? -1;
}

function elementAt(xml: string, start: number, tag: string) {
  if (start < 0) return "";
  const token = new RegExp("</?(?:(?:[A-Za-z_][\\w.-]*):)?" + tag + "(?:\\s[^>]*)?/?>", "gi");
  token.lastIndex = start;
  let depth = 0;
  let first = true;
  let match: RegExpExecArray | null;
  while ((match = token.exec(xml))) {
    const raw = match[0];
    const closing = raw.startsWith("</");
    const selfClosing = /\/\s*>$/.test(raw);
    if (!closing && !selfClosing) depth += 1;
    if (closing) depth -= 1;
    if (first) first = false;
    if (depth === 0 && !first) return xml.slice(start, match.index + raw.length);
  }
  return "";
}

function firstElement(xml: string, tag: string) {
  const start = findTagStart(xml, tag);
  return start < 0 ? "" : elementAt(xml, start, tag);
}

function innerXml(element: string) {
  const openEnd = element.indexOf(">");
  const closeStart = element.lastIndexOf("</");
  if (openEnd < 0 || closeStart < 0 || closeStart <= openEnd) return "";
  return element.slice(openEnd + 1, closeStart);
}

function topLevelElements(xml: string, tags: string) {
  const blocks: string[] = [];
  const token = new RegExp("<(?:(?:[A-Za-z_][\\w.-]*):)?(" + tags + ")(?=\\s|>)", "gi");
  let match: RegExpExecArray | null;
  while ((match = token.exec(xml))) {
    const block = elementAt(xml, match.index, match[1]);
    if (!block) break;
    blocks.push(block);
    token.lastIndex = match.index + block.length;
  }
  return blocks;
}

function topLevelBlocks(xml: string) {
  return topLevelElements(xml, "p|tbl");
}

function twipsToMm(value: string, fallback: number) {
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric > 0 ? (numeric / 1440) * 25.4 : fallback;
}

function halfPointsToPt(value: string, fallback: number) {
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric > 0 ? numeric / 2 : fallback;
}

function pointsToCss(value: string, fallback: number) {
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric >= 0 ? (numeric / 20) + "pt" : fallback + "pt";
}

function relationshipMap(xml: string) {
  const map = new Map<string, WordRelationship>();
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?>/gi)) {
    const attrs = match[1];
    const id = attribute(attrs, "Id");
    const target = attribute(attrs, "Target");
    const type = attribute(attrs, "Type");
    if (id && target) map.set(id, { target, type });
  }
  return map;
}

function resolveZipPath(base: string, target: string) {
  const parts = [...base.split("/"), ...target.replace(/^\/+/, "").split("/")];
  const resolved: string[] = [];
  for (const part of parts) {
    if (!part || part === ".") continue;
    if (part === "..") resolved.pop();
    else resolved.push(part);
  }
  return resolved.join("/");
}

function mimeTypeFor(name: string) {
  const extension = extensionOf(name);
  if (extension === "jpg" || extension === "jpeg") return "image/jpeg";
  if (extension === "gif") return "image/gif";
  if (extension === "svg") return "image/svg+xml";
  return "image/png";
}

function bytesToBase64(bytes: Uint8Array) {
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}

async function imageDataUri(zip: JSZip, documentDirectory: string, relationship: WordRelationship) {
  const path = resolveZipPath(documentDirectory, relationship.target);
  const entry = zip.file(path);
  if (!entry) return "";
  const bytes = await safeEntryBytes(zip, entry);
  return "data:" + mimeTypeFor(path) + ";base64," + bytesToBase64(bytes);
}

async function loadSafeDocxZip(file: File) {
  const source = await file.arrayBuffer();
  if (source.byteLength > MAX_TEMPLATE_BYTES) {
    throw new Error("The Word template is larger than the 25 MB offline safety limit.");
  }
  let zip: JSZip;
  try {
    zip = await withTemplateTimeout(JSZip.loadAsync(source, { checkCRC32: false, createFolders: false }), "the document archive");
  } catch {
    throw new Error(file.name + " is not a readable Word archive.");
  }
  const entries = Object.values(zip.files).filter((entry) => !entry.dir);
  if (entries.length > MAX_ARCHIVE_ENTRIES) {
    throw new Error("The Word template contains too many archive entries for safe offline conversion.");
  }
  let totalInflated = 0;
  let totalImages = 0;
  for (const entry of entries) {
    const unsafeName = String((entry as unknown as { unsafeOriginalName?: string }).unsafeOriginalName ?? entry.name);
    if (unsafeName.includes("..\\") || unsafeName.includes("../") || unsafeName.startsWith("/")) {
      throw new Error("The Word template contains an unsafe archive path.");
    }
    const declaredSize = Number((entry as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize);
    if (Number.isFinite(declaredSize) && declaredSize >= 0) {
      totalInflated += declaredSize;
      if (/^word\/media\//i.test(entry.name)) totalImages += declaredSize;
    }
    if (declaredSize > MAX_INFLATED_BYTES || totalInflated > MAX_INFLATED_BYTES) {
      throw new Error("The Word template expands beyond the 50 MB offline safety limit.");
    }
    if (totalImages > MAX_IMAGE_BYTES) {
      throw new Error("The Word template contains more than 12 MB of embedded images.");
    }
  }
  safeDocxBudgets.set(zip, { counted: new Set(), inflatedBytes: 0, imageBytes: 0 });
  return zip;
}

function runCss(run: string) {
  const properties = firstElement(run, "rPr");
  const css: string[] = [];
  if (firstElement(properties, "b")) css.push("font-weight:700");
  if (firstElement(properties, "i")) css.push("font-style:italic");
  const underline = firstElement(properties, "u");
  if (underline && attribute(underline, "val") !== "none") css.push("text-decoration:underline");
  const size = firstElement(properties, "sz");
  if (size) css.push("font-size:" + halfPointsToPt(attribute(size, "val"), 10) + "pt");
  const color = attribute(firstElement(properties, "color"), "val");
  if (/^[0-9a-f]{6}$/i.test(color)) css.push("color:#" + color);
  const highlight = attribute(firstElement(properties, "highlight"), "val");
  if (highlight && highlight !== "none") css.push("background-color:" + highlight);
  return css.join(";");
}

async function runHtml(run: string, zip: JSZip, relationships: Map<string, WordRelationship>) {
  const parts: string[] = [];
  const token = /<(?:(?:[A-Za-z_][\w.-]*):)?(t|tab|br)(?:\s[^>]*)?(?:\/?>[\s\S]*?<\/*(?:(?:[A-Za-z_][\w.-]*):)?\1>)?/gi;
  let match: RegExpExecArray | null;
  while ((match = token.exec(run))) {
    const raw = match[0];
    const type = match[1].toLowerCase();
    if (type === "t") {
      const openEnd = raw.indexOf(">");
      const closeStart = raw.lastIndexOf("</");
      const value = closeStart > openEnd ? raw.slice(openEnd + 1, closeStart) : "";
      parts.push(escapeHtml(decodeXml(value)));
    } else if (type === "tab") {
      parts.push("<span class=\"word-tab\">&nbsp;&nbsp;&nbsp;&nbsp;</span>");
    } else {
      const breakType = attribute(raw, "type");
      parts.push(breakType.toLowerCase() === "page" ? "<span class=\"word-page-break\"></span>" : "<br>");
    }
  }
  const drawing = firstElement(run, "drawing");
  const blip = drawing.match(/<[^>]*blip\b([^>]*)>/i);
  const embedId = attribute(blip?.[1] ?? "", "embed");
  const relationship = relationships.get(embedId);
  if (relationship) {
    const uri = await imageDataUri(zip, "word", relationship);
    if (uri) {
      const extent = firstElement(drawing, "extent");
      const widthMm = Number(attribute(extent, "cx")) / 914400 * 25.4;
      const heightMm = Number(attribute(extent, "cy")) / 914400 * 25.4;
      let dimensions = "";
      if (widthMm > 0 && heightMm > 0) {
        dimensions = " width=\"" + Math.min(widthMm, 180).toFixed(2) + "mm\" height=\"" + Math.min(heightMm, 250).toFixed(2) + "mm\"";
      }
      parts.push("<img class=\"word-image\" src=\"" + uri + "\"" + dimensions + " alt=\"\" />");
    }
  }
  return "<span style=\"" + runCss(run) + "\">" + parts.join("") + "</span>";
}

function paragraphCss(paragraph: string) {
  const properties = firstElement(paragraph, "pPr");
  const css: string[] = [];
  const alignment = attribute(firstElement(properties, "jc"), "val");
  if (["left", "center", "right", "both", "justify"].includes(alignment)) {
    css.push("text-align:" + (alignment === "both" ? "justify" : alignment));
  }
  const spacing = firstElement(properties, "spacing");
  if (spacing) {
    css.push("margin-top:" + pointsToCss(attribute(spacing, "before"), 0));
    css.push("margin-bottom:" + pointsToCss(attribute(spacing, "after"), 6));
    const line = Number(attribute(spacing, "line"));
    if (Number.isFinite(line) && line > 0) css.push("line-height:" + Math.max(1, line / 240).toFixed(3));
  }
  const indent = firstElement(properties, "ind");
  if (indent) {
    const left = Number(attribute(indent, "left"));
    const right = Number(attribute(indent, "right"));
    const firstLine = Number(attribute(indent, "firstLine"));
    if (left > 0) css.push("padding-left:" + (left / 1440 * 25.4).toFixed(2) + "mm");
    if (right > 0) css.push("padding-right:" + (right / 1440 * 25.4).toFixed(2) + "mm");
    if (firstLine > 0) css.push("text-indent:" + (firstLine / 1440 * 25.4).toFixed(2) + "mm");
  }
  return css.join(";");
}

async function paragraphHtml(paragraph: string, zip: JSZip, relationships: Map<string, WordRelationship>) {
  const runs = [...paragraph.matchAll(/<(?:(?:[A-Za-z_][\w.-]*):)?r(?=\s|>)[\s\S]*?<\/*(?:(?:[A-Za-z_][\w.-]*):)?r>/gi)].map((match) => match[0]);
  const content = (await Promise.all(runs.map((run) => runHtml(run, zip, relationships)))).join("");
  const paragraphProperties = firstElement(paragraph, "pPr");
  const pageBreakBefore = firstElement(paragraphProperties, "pageBreakBefore") ? " word-page-break-before" : "";
  return "<p class=\"word-paragraph" + pageBreakBefore + "\" style=\"" + paragraphCss(paragraph) + "\">" + (content || "&nbsp;") + "</p>";
}

async function tableHtml(table: string, zip: JSZip, relationships: Map<string, WordRelationship>) {
  const properties = firstElement(table, "tblPr");
  const borders = firstElement(properties, "tblBorders");
  const inside = firstElement(borders, "insideH");
  const border = attribute(inside, "val") === "nil" ? "none" : "1px solid #777";
  const rows = topLevelElements(innerXml(table), "tr");
  const body: string[] = [];
  for (const row of rows) {
    const cells = topLevelElements(innerXml(row), "tc");
    const cellHtml: string[] = [];
    for (const cell of cells) {
      const cellProperties = firstElement(cell, "tcPr");
      const shading = attribute(firstElement(cellProperties, "shd"), "fill");
      const span = Number(attribute(firstElement(cellProperties, "gridSpan"), "val"));
      const width = Number(attribute(firstElement(cellProperties, "tcW"), "w"));
      const styles = ["border:" + border, "padding:4pt 5pt", "vertical-align:top"];
      if (/^[0-9a-f]{6}$/i.test(shading)) styles.push("background-color:#" + shading);
      if (width > 0) styles.push("width:" + (width / 1440 * 25.4).toFixed(2) + "mm");
      const cellBlocks = await Promise.all(topLevelBlocks(innerXml(cell)).map((block) => {
        return /<(?:(?:[A-Za-z_][\w.-]*):)?tbl(?:\s|>)/i.test(block)
          ? tableHtml(block, zip, relationships)
          : paragraphHtml(block, zip, relationships);
      }));
      const colspan = span > 1 ? " colspan=\"" + span + "\"" : "";
      cellHtml.push("<td" + colspan + " style=\"" + styles.join(";") + "\">" + cellBlocks.join("") + "</td>");
    }
    body.push("<tr>" + cellHtml.join("") + "</tr>");
  }
  return "<table class=\"word-table\" style=\"border-collapse:collapse;border:" + border + ";width:100%\"><tbody>" + body.join("") + "</tbody></table>";
}

function sectionSettings(body: string) {
  const section = firstElement(body, "sectPr");
  const margins = firstElement(section, "pgMar");
  const pageSize = firstElement(section, "pgSz");
  return {
    orientation: attribute(pageSize, "orient").toLowerCase() === "landscape" ? "landscape" : "portrait",
    marginTop: twipsToMm(attribute(margins, "top"), 25.4),
    marginRight: twipsToMm(attribute(margins, "right"), 25.4),
    marginBottom: twipsToMm(attribute(margins, "bottom"), 25.4),
    marginLeft: twipsToMm(attribute(margins, "left"), 25.4),
  };
}

/**
 * Convert the readable subset of WordprocessingML needed for cover, index and
 * divider templates into a self-contained print document. Images are embedded
 * as data URIs; no remote CSS, fonts or hyperlinks are retained.
 */
export async function renderWordTemplateHtml(file: File) {
  const zip = await loadSafeDocxZip(file);
  const documentEntry = zip.file("word/document.xml");
  if (!documentEntry) throw new Error(file.name + " is not a valid Word document.");
  const documentXml = await safeEntryText(zip, documentEntry);
  const relationshipEntry = zip.file("word/_rels/document.xml.rels");
  const relationships = relationshipMap(relationshipEntry ? await safeEntryText(zip, relationshipEntry) : "");
  const body = firstElement(documentXml, "body");
  if (!body) throw new Error(file.name + " does not contain a readable Word document body.");
  const settings = sectionSettings(body);
  const blocks = topLevelBlocks(innerXml(body));
  const rendered = await Promise.all(blocks.map((block) => {
    return /<(?:(?:[A-Za-z_][\w.-]*):)?tbl(?:\s|>)/i.test(block)
      ? tableHtml(block, zip, relationships)
      : paragraphHtml(block, zip, relationships);
  }));
  const pageSize = settings.orientation === "landscape"
    ? A4_HEIGHT_MM + "mm " + A4_WIDTH_MM + "mm"
    : A4_WIDTH_MM + "mm " + A4_HEIGHT_MM + "mm";
  const contentWidth = settings.orientation === "landscape" ? A4_HEIGHT_MM : A4_WIDTH_MM;
  const contentHeight = settings.orientation === "landscape" ? A4_WIDTH_MM : A4_HEIGHT_MM;
  return "<!doctype html><html><head><meta charset=\"utf-8\"><style>" +
    "@page { size: " + pageSize + "; margin: 0; }" +
    "* { box-sizing: border-box; }" +
    "html, body { margin: 0; padding: 0; background: white; color: #111; }" +
    "body { font-family: Arial, Helvetica, sans-serif; font-size: 10pt; line-height: 1.2; }" +
    ".word-document { width: " + contentWidth + "mm; min-height: " + contentHeight + "mm; padding: " +
      settings.marginTop.toFixed(2) + "mm " + settings.marginRight.toFixed(2) + "mm " +
      settings.marginBottom.toFixed(2) + "mm " + settings.marginLeft.toFixed(2) + "mm; }" +
    ".word-paragraph { white-space: pre-wrap; overflow-wrap: anywhere; }" +
    ".word-paragraph:first-child { margin-top: 0; }" +
    ".word-page-break, .word-page-break-before { break-before: page; page-break-before: always; }" +
    ".word-tab { display: inline-block; min-width: 1.5em; }" +
    ".word-image { max-width: 100%; height: auto; object-fit: contain; vertical-align: middle; }" +
    ".word-table { table-layout: fixed; margin: 7pt 0; }" +
    ".word-table p { margin: 0 0 3pt; }" +
    ".word-table tr { break-inside: avoid; page-break-inside: avoid; }" +
    "</style></head><body><main class=\"word-document\">" + rendered.join("") + "</main></body></html>";
}

/** Convert a DOCX template through Electron's local Chromium print engine. */
export async function convertWordTemplate(file: File) {
  const extension = extensionOf(file.name);
  if (extension !== "docx") {
    throw new Error("Unsupported Word template: " + file.name);
  }
  const html = await renderWordTemplateHtml(file);
  if (!window.bundleBuilderDesktop?.convertTemplate) {
    throw new Error("Word templates can only be converted in the packaged offline desktop application.");
  }
  const bytes = await window.bundleBuilderDesktop.convertTemplate(html, file.name);
  const pdfName = file.name.replace(/\.docx$/i, ".pdf");
  const pdfBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
  return new File([pdfBuffer], pdfName, { type: "application/pdf" });
}
