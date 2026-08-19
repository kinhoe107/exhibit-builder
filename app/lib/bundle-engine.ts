"use client";

import JSZip from "jszip";
import {
  PDFDocument,
  PDFFont,
  PDFHexString,
  PDFName,
  PDFPage,
  StandardFonts,
  degrees,
  rgb,
} from "pdf-lib";
import type { BuildResolution, BundleLayoutSettings, NonA4PageHandling, PageNumberSettings, TemplateFile } from "./bundle-types.ts";
import { BUNDLE_PROFILES, DEFAULT_BUNDLE_LAYOUT, DEFAULT_PAGINATION, coverPrintsPageNumber, coverPrintsVolumeLabel, coverWritesMatterText } from "./bundle-types.ts";
import { hasBlockingPreflight, runPreflight } from "./preflight.ts";
import { applyBuildResolutions, candidateIsExcluded, findResolution, shouldSkipOcr, sourceIsExcluded, templateFallbackSlots } from "./build-resolutions.ts";
import { deriveExhibitGroups, orderExhibitGroups } from "./exhibit-groups.ts";
import { createBuildPlan, type BuildPlanIndexNode, type BundleBuildPlan, type PlannedBuildVolume } from "./build-plan.ts";
import {
  bundleArrangementFromLegacyOrder,
  flattenBundleArrangement,
  reconcileBundleArrangement,
  type BundleArrangement,
} from "./bundle-arrangement.ts";
import {
  applyDetectedDateColumn,
  createIndexLayoutPlan,
  CUSTOM_TEMPLATE_INDEX_GEOMETRY,
  detectIndexTemplateDateColumn,
  type IndexLayoutGeometryProfile,
  type IndexLayoutPlan,
  type IndexLayoutRowInput,
  type PlannedIndexExhibit,
} from "./index-layout.ts";
import { parseBundleEmail } from "./email.ts";
import { rederiveEmailChildren, type EmailAttachmentChild, type EmailChildDisposition, emailChildrenForDisposition } from "./email-attachments.ts";
import { ocrPdfLocally } from "./ocr.ts";
import { analyseXlsx, analyseXlsxInWorker, type WorkbookAnalysis } from "./xlsx.ts";
import { convertWordTemplate } from "./template-converter.ts";
import { assertPdfActionsSafe, unsafePdfActions, type UnsafePdfAction } from "./pdf-action-safety.ts";
import { applyTemplateMatterPatches, type TemplatePagePlacement } from "./template-matter-writeback.ts";
import { effectiveMatterValues, matterValuesFromConfirmation, resolvedBundleTitle } from "./template-matter-review.ts";

const INPUT_LIMITS = {
  docxBytes: 50 * 1024 * 1024,
  docxEntries: 500,
  docxInflatedBytes: 128 * 1024 * 1024,
  docxInflationRatio: 100,
  docxXmlBytes: 32 * 1024 * 1024,
  docxSelectedXmlBytes: 48 * 1024 * 1024,
  emlBytes: 32 * 1024 * 1024,
  pdfBytes: 256 * 1024 * 1024,
  pdfPages: 1_500,
  pdfDimensionPoints: 10_000,
  pdfAreaPoints: 25_000_000,
  pdfTotalMilliseconds: 5 * 60 * 1_000,
  pdfStepMilliseconds: 10_000,
  textBytes: 25 * 1024 * 1024,
  evidenceFiles: 500,
} as const;

async function withinTime<T>(
  operation: Promise<T>,
  milliseconds: number,
  message: string,
  onTimeout?: () => void | Promise<void>,
): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      operation,
      new Promise<T>((_resolve, reject) => {
        timer = setTimeout(() => {
          if (onTimeout) void Promise.resolve(onTimeout()).catch(() => undefined);
          reject(new Error(message));
        }, milliseconds);
      }),
    ]);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

function assertInputFileSize(file: File, extension = extensionOf(file.name)) {
  const limit = extension === "pdf" ? INPUT_LIMITS.pdfBytes : extension === "docx" ? INPUT_LIMITS.docxBytes : extension === "eml" ? INPUT_LIMITS.emlBytes : extension === "xlsx" ? 25 * 1024 * 1024 : INPUT_LIMITS.textBytes;
  if (!file.size || file.size > limit) throw new Error(`${file.name} is empty or exceeds the ${Math.round(limit / (1024 * 1024))} MB ${extension.toUpperCase() || "file"} safety limit.`);
}

async function safeDocxZip(file: File) {
  assertInputFileSize(file, "docx");
  const zip = await withinTime(
    JSZip.loadAsync(await file.arrayBuffer(), { createFolders: false }),
    30_000,
    `${file.name} did not open within the 30-second DOCX safety limit.`,
  );
  const entries = Object.values(zip.files);
  const inflated = entries.reduce((total, entry) => total + ((entry as any)._data?.uncompressedSize ?? 0), 0);
  if (entries.length > INPUT_LIMITS.docxEntries || inflated > INPUT_LIMITS.docxInflatedBytes || inflated / Math.max(file.size, 1) > INPUT_LIMITS.docxInflationRatio || entries.some((entry) => entry.name.includes("..") || entry.name.includes("\\") || entry.name.startsWith("/"))) {
    throw new Error(`${file.name} exceeds the DOCX archive safety limits.`);
  }
  return zip;
}

function docxEntrySize(entry: JSZip.JSZipObject) {
  return Number((entry as any)._data?.uncompressedSize ?? 0);
}

async function readDocxXmlEntry(file: File, entry: JSZip.JSZipObject) {
  const declaredSize = docxEntrySize(entry);
  if (declaredSize > INPUT_LIMITS.docxXmlBytes) {
    throw new Error(`${file.name} contains an XML part larger than the 32 MB DOCX safety limit.`);
  }
  const xml = await withinTime(
    entry.async("text"),
    30_000,
    `${file.name} XML extraction exceeded the 30-second DOCX safety limit.`,
  );
  if (new TextEncoder().encode(xml).byteLength > INPUT_LIMITS.docxXmlBytes) {
    throw new Error(`${file.name} contains an XML part larger than the 32 MB DOCX safety limit.`);
  }
  return xml;
}

export type StatementParagraph = {
  number: number;
  text: string;
};

export type EvidenceRecord = {
  id: string;
  file: File;
  name: string;
  extension: string;
  text: string;
  marker: string | null;
  sha256: string;
  pageCount: number;
  readableText: boolean;
  encrypted: boolean;
  rotationPages: number[];
  annotationPages?: number[];
  unsafePdfActions?: UnsafePdfAction[];
  ocrPages: Array<{ text: string; confidence: number }>;
  ocrStatus: "not-needed" | "completed" | "failed" | "unavailable";
  ocrFailureReason?: string;
  pageSizes?: Array<{
    page: number;
    width: number;
    height: number;
    orientation: "portrait" | "landscape";
    isA4: boolean;
    wouldAddMarginsOnA4: boolean;
    hasAnnotations?: boolean;
  }>;
  workbook?: WorkbookAnalysis;
  sheetSelections?: Array<{ name: string; included: boolean; range: string }>;
  emailAttachments?: EmailAttachmentChild[];
  /** In-memory attachment exhibit; never copied into the project archive or recovery journal. */
  derivedFromEmail?: { parentSha256: string; childIdentity: string };
};

export type ExhibitCandidate = {
  id: string;
  mark: string;
  provisionalNumber: number;
  description: string;
  /** Reviewer-facing alternative names retained locally for searching and audit. */
  aliases?: string[];
  /** Local review rationale; never added to the source witness statement. */
  reviewNote?: string;
  date: string;
  paragraph: number;
  citation: string;
  citationToken?: string;
  citationOrdinal?: number;
  citationCount?: number;
  exhibitInitials?: string;
  exhibitSequence?: number;
  requestedExhibitPageStart?: number;
  requestedExhibitPageEnd?: number;
  citationResolution?: "resolved" | "unresolved" | "none";
  /** Nearby statement text retained for a readable human review context. */
  context?: string;
  /** Structured neighbouring statement paragraphs for the review card. */
  contextParagraphs?: Array<{ paragraph: number; position: "previous" | "following"; text: string }>;
  discoverySignals: string[];
  evidenceId: string | null;
  confidence: number;
  rationale: string;
  included: boolean;
  confirmed: boolean;
  /** How the reviewer recorded the current confirmation. */
  confirmationMethod?: "individual" | "bulk";
  /** Local audit timestamp for the current confirmation. */
  confirmedAt?: string;
  statementName?: string;
  witnessInitials?: string;
  witnessKey?: string;
  statementId?: string;
  pageStart?: number;
  pageEnd?: number;
  alternativeEvidenceIds?: string[];
  sequenceOrder?: number;
  /** Required only when another included citation selects the same file hash. */
  repeatDecision?: "pending" | "same" | "separate";
  /** Deliberately added by the reviewer without a statement citation. */
  manualAddition?: boolean;
  /** Audit timestamp for an intentional uncited addition. */
  manualAddedAt?: string;
  manualWarningAcknowledgedAt?: string;
  emailAttachmentDispositions?: Record<string, EmailChildDisposition>;
  parentEmailProvenance?: {
    parentName: string;
    parentSha256: string;
    childIdentity: string;
    childSha256: string;
  };
};

export type BundleStatementInput = {
  id: string;
  file: File;
  witnessName: string;
  witnessInitials: string;
};

export type AnalysisResult = {
  statementName: string;
  statementHash: string;
  caseTitle: string;
  candidates: ExhibitCandidate[];
  evidence: EvidenceRecord[];
  unreferenced: EvidenceRecord[];
  statementWarnings: string[];
  generatedAt: string;
  statementId?: string;
  witnessName?: string;
  witnessInitials?: string;
};

export type BundleRecord = {
  mark: string;
  description: string;
  fileName: string;
  startPage: number;
  endPage: number;
  statementParagraph: number | null;
  statementReferences: Array<{
    paragraph: number;
    citation: string;
    statementName?: string;
    statementId?: string;
    witnessInitials?: string;
    citationToken?: string;
    citationOrdinal?: number;
    citationCount?: number;
    exhibitInitials?: string;
    exhibitSequence?: number;
    requestedExhibitPageStart?: number;
    requestedExhibitPageEnd?: number;
    citationResolution?: "resolved" | "unresolved" | "none";
    exhibitPageStart?: number;
    exhibitPageEnd?: number;
    exhibitPageLabelStart?: string;
    exhibitPageLabelEnd?: string;
    volumeNumber?: number;
  }>;
  exhibitNumber?: number;
  exhibitPageStart?: number;
  exhibitPageEnd?: number;
  exhibitPageLabelStart?: string;
  exhibitPageLabelEnd?: string;
  /** Administrative volume number within the same AH1 exhibit bundle. */
  volumeNumber?: number;
  sourceHash: string;
  workbookSheet?: { id: string; name: string; path: string; range: string };
  workbookSheets?: Array<{ id: string; name: string; path: string; range: string }>;
  manualAddition?: boolean;
  citationStatus?: "cited" | "not-cited-manual-addition";
  documentDate?: string;
  manualAddedAt?: string;
  manualWarningAcknowledgedAt?: string;
  emailAttachments?: Array<{
    name: string;
    identity: string;
    sha256: string;
    parentSha256: string;
    disposition: EmailChildDisposition | "unresolved";
  }>;
};

export type BuildResult = {
  bytes: Uint8Array;
  fileName: string;
  sha256: string;
  pageCount: number;
  records: BundleRecord[];
  manifest: Record<string, unknown>;
  checks: Array<{ label: string; status: "pass" | "warning" | "blocking"; detail: string }>;
  buildPlan?: BundleBuildPlan;
  volumeZipBytes?: Uint8Array;
  volumeZipFileName?: string;
  volumeZipSha256?: string;
  volumes?: Array<{
    number: number;
    label: string;
    bytes: Uint8Array;
    fileName: string;
    sha256: string;
    pageCount: number;
    records: BundleRecord[];
    manifest: Record<string, unknown>;
    checks: Array<{ label: string; status: "pass" | "warning" | "blocking"; detail: string }>;
  }>;
};

export type BuildOptions = {
  profileId?: string;
  pagination?: PageNumberSettings;
  templates?: TemplateFile[];
  resolutions?: BuildResolution[];
  layout?: BundleLayoutSettings;
  /** Persisted canonical exhibit-group IDs; authoritative over legacy sequenceOrder. */
  canonicalOrder?: string[];
  /** Authoritative schema-8 order and reader-facing index sections. */
  arrangement?: BundleArrangement;
  /** Source IDs default to proportional A4 conversion unless explicitly retained. */
  pageSizeChoices?: Record<string, NonA4PageHandling>;
  /**
   * Desktop-only, read-only Microsoft Excel print export. Production builds
   * deliberately have no simplified spreadsheet fallback because it would
   * change workbook colours, formats, borders and print layout.
   */
  workbookExporter?: (
    file: File,
    sheets: Array<{ name: string; range: string; orientation: "portrait" | "landscape" }>,
  ) => Promise<Array<{ name: string; range: string; orientation?: "portrait" | "landscape"; bytes: Uint8Array }>>;
  /**
   * Deliberately small, local-only progress hook.  It lets the desktop
   * communicate that a long PDF job is alive without exposing document data
   * or altering the deterministic output path.
   */
  onProgress?: (stage: string, detail?: string) => void;
};

async function createVolumeArchive(
  volumes: NonNullable<BuildResult["volumes"]>,
  bundleIdentity: string,
) {
  const zip = new JSZip();
  const stableZipDate = new Date("1980-01-01T00:00:00.000Z");
  for (const volume of volumes) {
    zip.file(volume.fileName, volume.bytes, { date: stableZipDate });
    zip.file(volume.fileName.replace(/\.pdf$/i, "_Manifest.json"), JSON.stringify(volume.manifest, null, 2), { date: stableZipDate });
  }
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return {
    bytes,
    sha256: await sha256(bytes),
    fileName: `Exhibit_Bundle_${bundleIdentity.replace(/\s+/g, "")}_Volumes.zip`,
  };
}

/**
 * Attaches the final substantive audit metadata and, for administrative
 * volumes, regenerates the ZIP from those exact final manifests.  This keeps
 * the embedded audit records byte-for-byte consistent with the downloadable
 * in-memory manifests.
 */
export async function finalizeBuildAudit(
  result: BuildResult,
  inputFingerprint: string,
  rebuildComparison: unknown,
): Promise<BuildResult> {
  const audit = { inputFingerprint, rebuildComparison };
  if (!result.volumes?.length) return { ...result, manifest: { ...result.manifest, ...audit } };
  const volumes = result.volumes.map((volume) => ({
    ...volume,
    manifest: { ...volume.manifest, ...audit },
  }));
  const archive = await createVolumeArchive(volumes, result.buildPlan?.bundleIdentity ?? "Exhibit Bundle");
  const priorOutput = (result.manifest.output ?? {}) as Record<string, unknown>;
  const manifest = {
    ...result.manifest,
    ...audit,
    output: { ...priorOutput, fileName: archive.fileName, sha256: archive.sha256 },
  };
  return {
    ...result,
    bytes: archive.bytes,
    fileName: archive.fileName,
    sha256: archive.sha256,
    manifest,
    volumes,
    volumeZipBytes: archive.bytes,
    volumeZipFileName: archive.fileName,
    volumeZipSha256: archive.sha256,
  };
}

const MONTHS =
  "January|February|March|April|May|June|July|August|September|October|November|December";
const A4_WIDTH = 595.28;
const A4_HEIGHT = 841.89;
const PAGE_FOOTER_TOP = 38;

const STOP_WORDS = new Set([
  "about",
  "after",
  "again",
  "also",
  "been",
  "being",
  "copy",
  "dated",
  "document",
  "exhibit",
  "exhibited",
  "from",
  "have",
  "into",
  "northbridge",
  "meridian",
  "that",
  "their",
  "there",
  "these",
  "this",
  "those",
  "under",
  "were",
  "which",
  "with",
  "would",
]);

export const SAMPLE_STATEMENT = "01_GUIDED_SAMPLE_Witness_Statement.docx";

export const SAMPLE_EVIDENCE = [
  "01_SAMPLE_Agreement.pdf",
  "02_SAMPLE_Invoice.pdf",
  "03_SAMPLE_Project_Report.docx",
  "04_SAMPLE_Claimant_Email.eml",
  "05_SAMPLE_Cost_Workbook.xlsx",
  "06_SAMPLE_Unreferenced_Checklist.pdf",
];

export const SAMPLE_TEMPLATES = [
  { slot: "cover" as const, name: "00_GUIDED_SAMPLE_Cover_Template.pdf" },
  { slot: "index" as const, name: "00_GUIDED_SAMPLE_Index_Template.pdf" },
];

export const SAMPLE_REQUIRED_FILES = [
  SAMPLE_STATEMENT,
  ...SAMPLE_EVIDENCE,
];

function extensionOf(name: string) {
  return name.split(".").pop()?.toLowerCase() ?? "";
}

function formatExhibitMark(number: number) {
  return `AH ${number}`;
}

function normalizeExhibitMark(mark: string) {
  const match = mark.match(/\bAH\s*(\d{1,3})\b/i);
  return match ? `AH${Number(match[1])}` : mark.replace(/\s+/g, "").toUpperCase();
}

function hasExhibitPlaceholder(text: string) {
  return parseStatementCitationTokens(text).length > 0;
}

function normalizeWhitespace(value: string) {
  return value.replace(/\s+/g, " ").trim();
}

function decodeXml(value: string) {
  return value
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&#x([0-9a-f]+);/gi, (_, code) =>
      String.fromCodePoint(Number.parseInt(code, 16)),
    )
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function sanitizePdfText(value: string) {
  return value
    .replace(/[–—]/g, "-")
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/[^\x09\x0A\x0D\x20-\x7E]/g, "");
}

async function sha256(bytes: ArrayBuffer | Uint8Array) {
  const source =
    bytes instanceof Uint8Array
      ? bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
      : bytes;
  const digest = await crypto.subtle.digest("SHA-256", source as ArrayBuffer);
  return Array.from(new Uint8Array(digest))
    .map((part) => part.toString(16).padStart(2, "0"))
    .join("");
}

async function docxXml(file: File) {
  const zip = await safeDocxZip(file);
  const entry = zip.file("word/document.xml");
  if (!entry) throw new Error(`${file.name} is not a valid DOCX file.`);
  return readDocxXmlEntry(file, entry);
}

function parseDocxParagraphs(xml: string) {
  const paragraphs = xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) ?? [];
  return paragraphs
    .map((paragraph) =>
      normalizeWhitespace(
        Array.from(paragraph.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g))
          .map((match) => decodeXml(match[1]))
          .join(""),
      ),
    )
    .filter(Boolean);
}

async function extractDocxText(file: File) {
  const zip = await safeDocxZip(file);
  const contentFiles = Object.keys(zip.files)
    .filter((name) =>
      /^word\/(?:document|header\d+|footer\d+)\.xml$/i.test(name),
    )
    .sort((left, right) =>
      left === "word/document.xml"
        ? -1
        : right === "word/document.xml"
          ? 1
          : left.localeCompare(right),
    );
  const selectedInflatedBytes = contentFiles.reduce((total, name) => {
    const entry = zip.file(name);
    return total + (entry ? docxEntrySize(entry) : 0);
  }, 0);
  if (selectedInflatedBytes > INPUT_LIMITS.docxSelectedXmlBytes) {
    throw new Error(`${file.name} exceeds the 48 MB selected-content DOCX safety limit.`);
  }
  // Expand one XML part at a time so a document with many headers and footers
  // cannot multiply temporary memory use through concurrent decompression.
  const content: string[] = [];
  for (const name of contentFiles) {
    const entry = zip.file(name);
    if (!entry) continue;
    content.push(parseDocxParagraphs(await readDocxXmlEntry(file, entry)).join("\n"));
  }
  return content.filter(Boolean).join("\n");
}

async function extractStatementParagraphs(file: File) {
  const xml = await docxXml(file);
  const raw = parseDocxParagraphs(xml);
  const paragraphs: StatementParagraph[] = [];
  let pendingNumber: number | null = null;
  let numberedListCounter = 0;

  const blocks = xml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) ?? [];
  for (const block of blocks) {
    const text = normalizeWhitespace(Array.from(block.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)).map((match) => decodeXml(match[1])).join(""));
    if (!text) continue;
    const inlineNumberMatch = text.match(/^(\d{1,3})\.\s*(\S[\s\S]*)$/);
    if (inlineNumberMatch) {
      paragraphs.push({ number: Number(inlineNumberMatch[1]), text: inlineNumberMatch[2] });
      pendingNumber = null;
      continue;
    }
    if (/<w:numPr>[\s\S]*?<w:ilvl\b[^>]*w:val="0"[^>]*\/>[\s\S]*?<\/w:numPr>/i.test(block)) {
      numberedListCounter += 1;
      paragraphs.push({ number: numberedListCounter, text });
      pendingNumber = null;
      continue;
    }
    const numberMatch = text.match(/^(\d{1,3})\.$/);
    if (numberMatch) {
      pendingNumber = Number(numberMatch[1]);
      continue;
    }
    if (pendingNumber !== null) {
      paragraphs.push({ number: pendingNumber, text });
      pendingNumber = null;
    }
  }

  return { raw, paragraphs };
}

let pdfExtractionQueue: Promise<void> = Promise.resolve();

type PdfTextInspection = {
  text: string;
  pageTexts: string[];
  pageCount: number;
  rotationPages: number[];
  annotationPages: number[];
  unsafePdfActions: UnsafePdfAction[];
  pageSizes: NonNullable<EvidenceRecord["pageSizes"]>;
};

async function extractPdfTextImmediately(file: File): Promise<PdfTextInspection> {
  const startedAt = Date.now();
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  if (!pdfjs.GlobalWorkerOptions.workerSrc) {
    pdfjs.GlobalWorkerOptions.workerSrc = new URL(
      "../../node_modules/pdfjs-dist/legacy/build/pdf.worker.min.mjs",
      import.meta.url,
    ).toString();
  }
  assertInputFileSize(file, "pdf");
  const sourceBytes = new Uint8Array(await file.arrayBuffer());
  const actionDocument = await withinTime(
    PDFDocument.load(sourceBytes, { updateMetadata: false }),
    30_000,
    `${file.name} active-action inspection exceeded the 30-second PDF safety limit.`,
  );
  const unsafeActions = unsafePdfActions(actionDocument);
  const task = pdfjs.getDocument({
    data: sourceBytes,
    isEvalSupported: false,
    useWorkerFetch: false,
    verbosity: 0,
  });
  const document = await withinTime(
    task.promise,
    30_000,
    `${file.name} did not open within the 30-second PDF safety limit.`,
    () => task.destroy(),
  );
  let destroyed = false;
  const destroyDocument = async () => {
    if (destroyed) return;
    destroyed = true;
    await document.destroy();
  };
  const pdfStep = <T>(operation: Promise<T>, pageNumber: number, action: string) => {
    const remaining = INPUT_LIMITS.pdfTotalMilliseconds - (Date.now() - startedAt);
    if (remaining <= 0) {
      void destroyDocument().catch(() => undefined);
      throw new Error(`${file.name} exceeded the five-minute total PDF safety limit.`);
    }
    const milliseconds = Math.min(INPUT_LIMITS.pdfStepMilliseconds, remaining);
    const totalLimitReached = milliseconds === remaining;
    return withinTime(
      operation,
      milliseconds,
      totalLimitReached
        ? `${file.name} exceeded the five-minute total PDF safety limit.`
        : `${file.name} page ${pageNumber} ${action} exceeded the 10-second PDF safety limit.`,
      destroyDocument,
    );
  };
  try {
    if (document.numPages > INPUT_LIMITS.pdfPages) throw new Error(`${file.name} exceeds the ${INPUT_LIMITS.pdfPages}-page PDF safety limit.`);
    const pages: string[] = [];
    const rotationPages: number[] = [];
    const annotationPages: number[] = [];
    const pageSizes: NonNullable<EvidenceRecord["pageSizes"]> = [];
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
      const page = await pdfStep(document.getPage(pageNumber), pageNumber, "loading");
      try {
        if (page.rotate % 360 !== 0) rotationPages.push(pageNumber);
        const annotations = await pdfStep(page.getAnnotations({ intent: "display" }), pageNumber, "annotation reading");
        if (annotations.length) annotationPages.push(pageNumber);
        const viewport = page.getViewport({ scale: 1 });
        if (viewport.width > INPUT_LIMITS.pdfDimensionPoints || viewport.height > INPUT_LIMITS.pdfDimensionPoints || viewport.width * viewport.height > INPUT_LIMITS.pdfAreaPoints) throw new Error(`${file.name} page ${pageNumber} has unsafe page dimensions.`);
        const orientation = viewport.width > viewport.height ? "landscape" as const : "portrait" as const;
        const targetWidth = orientation === "landscape" ? A4_HEIGHT : A4_WIDTH;
        const targetHeight = orientation === "landscape" ? A4_WIDTH : A4_HEIGHT;
        const isA4 = Math.abs(viewport.width - targetWidth) < 2 && Math.abs(viewport.height - targetHeight) < 2;
        const sourceRatio = viewport.width / Math.max(1, viewport.height);
        const targetRatio = targetWidth / targetHeight;
        pageSizes.push({
          page: pageNumber,
          width: viewport.width,
          height: viewport.height,
          orientation,
          isA4,
          wouldAddMarginsOnA4: !isA4 && Math.abs(sourceRatio - targetRatio) > 0.01,
          hasAnnotations: annotations.length > 0,
        });
        const content = await pdfStep(page.getTextContent(), pageNumber, "text extraction");
        pages.push(normalizeWhitespace(
          content.items
            .map((item) => ("str" in item ? item.str : ""))
            .join(" "),
        ));
      } finally {
        page.cleanup();
      }
    }
    return {
      text: normalizeWhitespace(pages.join("\n")),
      pageTexts: pages,
      pageCount: document.numPages,
      rotationPages,
      annotationPages,
      unsafePdfActions: unsafeActions,
      pageSizes,
    };
  } finally {
    await destroyDocument().catch(() => undefined);
  }
}

function extractPdfText(file: File) {
  const extraction = pdfExtractionQueue.then(
    () => extractPdfTextImmediately(file),
    () => extractPdfTextImmediately(file),
  );
  pdfExtractionQueue = extraction.then(
    () => undefined,
    () => undefined,
  );
  return extraction;
}

async function extractFileText(file: File) {
  const extension = extensionOf(file.name);
  if (extension === "docx") return extractDocxText(file);
  if (extension === "pdf") return (await extractPdfText(file)).text;
  if (extension === "eml") return (await parseBundleEmail(await file.text())).body;
  if (extension === "txt") return await file.text();
  if (extension === "xlsx") return (await analyseXlsx(file)).sheets.map((sheet) => `${sheet.name} ${sheet.cells.map((cell) => cell.value).join(" ")}`).join("\n");
  throw new Error(`Unsupported file type: ${file.name}`);
}

function inferCaseTitle(text: string) {
  if (
    /NORTHBRIDGE RENEWABLES LIMITED/i.test(text) &&
    /MERIDIAN COMPONENTS LIMITED/i.test(text)
  ) {
    return "Northbridge Renewables Limited v Meridian Components Limited";
  }
  return "New matter";
}

function extractDate(text: string) {
  const full = text.match(
    new RegExp(`\\b\\d{1,2}\\s+(?:${MONTHS})(?:\\s+\\d{4})?\\b`, "i"),
  );
  if (!full?.[0]) return "Date not stated";
  return /\d{4}$/.test(full[0]) ? full[0] : `${full[0]} 2026`;
}

function extractDates(text: string) {
  return Array.from(text.matchAll(new RegExp(`\\b\\d{1,2}\\s+(?:${MONTHS})(?:\\s+\\d{4})?\\b`, "gi"))).map((match) => /\d{4}$/.test(match[0]) ? match[0] : `${match[0]} 2026`);
}

function suggestDescription(text: string, date: string) {
  const lower = text.toLowerCase();
  if (lower.includes("sample agreement")) return "Sample agreement";
  if (lower.includes("sample invoice")) return "Sample invoice";
  if (lower.includes("sample project report")) return "Sample project report";
  if (lower.includes("sample claimant email")) return "Sample claimant email";
  if (lower.includes("sample cost workbook")) return "Sample cost workbook";
  if (lower.includes("purchase order")) return "Purchase Order NRL-1047";
  if (lower.includes("termination letter") || lower.includes("terminated"))
    return "Notice of Termination";
  if (lower.includes("quality-warning") || lower.includes("thermal shutdown"))
    return `Quality-warning email chain - ${date}`;
  if (
    lower.includes("component shortage") ||
    lower.includes("processor supply") ||
    (lower.includes("shortage") && lower.includes("email"))
  )
    return `Processor supply email chain - ${date}`;
  if (
    lower.includes("delivery dates") ||
    (lower.includes("email") && lower.includes("delivery"))
  )
    return `Delivery assurance email chain - ${date}`;
  if (lower.includes("progress meeting") || lower.includes("minutes"))
    return "Progress Meeting Minutes";
  if (lower.includes("notice of delay")) return "Formal Notice of Delay";
  if (lower.includes("inspection report"))
    return "Verity Independent Inspection Report";
  if (lower.includes("invoice")) return "Apex Controls Invoice AC-7782";
  if (lower.includes("supply agreement")) return "Executed Supply Agreement";
  if (lower.includes("email")) return `Email chain - ${date}`;
  return `Document referred to at ${date}`;
}

function tokens(text: string) {
  return new Set(
    text
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .split(/\s+/)
      .filter((token) => token.length > 3 && !STOP_WORDS.has(token)),
  );
}

type EvidenceMatchProfile = { sourceTokens: Set<string>; lowerText: string };

function evidenceMatchProfile(evidence: EvidenceRecord): EvidenceMatchProfile {
  return {
    sourceTokens: tokens(`${evidence.name} ${evidence.text.slice(0, 4000)}`),
    lowerText: evidence.text.toLowerCase(),
  };
}

function fallbackScore(candidate: ExhibitCandidate, evidence: EvidenceRecord, profile = evidenceMatchProfile(evidence), target = tokens(`${candidate.description} ${candidate.citation} ${candidate.date}`)) {
  const source = profile.sourceTokens;
  const shared = Array.from(target).filter((token) => source.has(token)).length;
  const overlap = target.size ? shared / target.size : 0;
  const dateMatch =
    candidate.date !== "Date not stated" &&
    profile.lowerText.includes(candidate.date.toLowerCase());
  return Math.min(88, Math.round(overlap * 65 + (dateMatch ? 23 : 0)));
}

export type ParsedStatementCitationToken = {
  raw: string;
  sourceRaw: string;
  index: number;
  endIndex: number;
  ordinal: number;
  exhibitInitials?: string;
  exhibitSequence?: number;
  requestedExhibitPageStart?: number;
  requestedExhibitPageEnd?: number;
  citationResolution: "resolved" | "unresolved" | "none";
  contextText?: string;
};

const DOCUMENT_NOUN = String.raw`(?:purchase\s+orders?|email\s+chains?|appendices|appendix|contracts?|agreements?|emails?|invoices?|letters?|reports?|notices?|minutes|schedules?|spreadsheets?|workbooks?|replies|reply)`;
const DOCUMENT_NOUN_PATTERN = new RegExp(`\\b${DOCUMENT_NOUN}\\b`, "i");
const COORDINATED_DOCUMENT_NOUN = String.raw`(?:${DOCUMENT_NOUN}|ledgers?)`;
const COORDINATED_AND_TO_THE = /\band\s+to\s+the\b/i;
const ATTACH_OBJECT_BLOCK = /\b(?:importance|significance|weight|value|condition|priority|blame|liability)\b/i;

const ROUND_ASIDE_PREFIX = /^(?:para|paras|page|pages|pp|p|clause|item|vol|volume|section|annex|table|fig|figure|sch|no|nos|covid|jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?|q|h|fy)$/i;
const TITLE_ABBREVIATION = String.raw`(?:Dr|Mr|Mrs|Ms|Prof|Sr|Jr|No|Nos|vs|etc|Inc|Ltd|Co|v)`;

function hasDocumentNoun(text: string) {
  DOCUMENT_NOUN_PATTERN.lastIndex = 0;
  return DOCUMENT_NOUN_PATTERN.test(text);
}

function splitBoundedClauses(text: string) {
  return text
    .split(new RegExp(`(?<!(?:\\b${TITLE_ABBREVIATION})\\.)(?<=[.!?])\\s+`))
    .map((part) => part.trim())
    .filter(Boolean);
}

function isRoundExhibitMark(trimmed: string) {
  if (/^(?:exhibit|exhib(?:it)?\s*x{1,4})$/i.test(trimmed)) return true;
  const mark = trimmed.match(/^([A-Z]{2,8})(?:[\s-]*(x{1,4}|\d{1,3}))$/i);
  return Boolean(mark && !ROUND_ASIDE_PREFIX.test(mark[1]));
}

function tokenContext(text: string, start: number, end: number) {
  const before = text.lastIndexOf(";", start - 1);
  const after = text.indexOf(";", end);
  const context = text.slice(before >= 0 ? before + 1 : 0, after >= 0 ? after : text.length).trim();
  return context || text;
}

function wrapCitationToken(open: string, close: string, value: string) {
  return `${open}${value}${close}`;
}

function parseBracketToken(
  value: string,
  sourceRaw: string,
  index: number,
  endIndex: number,
  ordinal: number,
  contextText: string,
  open = "[",
  close = "]",
): ParsedStatementCitationToken | undefined {
  const trimmed = value.trim();
  const raw = wrapCitationToken(open, close, trimmed);
  const page = trimmed.match(/^([A-Z]{1,8})\s*(?:-|\u2013|\u2014)?\s*(\d{1,4})\s*\/\s*(\d{1,5}|xx)\s*(?:-|\u2013|\u2014)?\s*(\d{1,5}|xx)?$/i);
  if (page) {
    const first = page[3].toLowerCase();
    const second = page[4]?.toLowerCase();
    const unresolved = first === "xx" || second === "xx";
    return {
      raw, sourceRaw, index, endIndex, ordinal, contextText,
      exhibitInitials: page[1].toUpperCase(),
      exhibitSequence: Number(page[2]),
      requestedExhibitPageStart: unresolved ? undefined : Number(first),
      requestedExhibitPageEnd: unresolved ? undefined : Number(second ?? first),
      citationResolution: unresolved ? "unresolved" : "resolved",
    };
  }
  const generic = /^(?:exhibit|exhib(?:it)?\s*x{1,4}|[A-Z][A-Z0-9._ ]{0,23}\s*(?:-|[\u2013\u2014]|\s+)\s*(?:x{1,4}|\d{1,3}|[A-Z])|[A-Z]{1,8}(?:[\s-]*x{1,4}|\d{1,3}))$/i;
  if (open === "(" ? !isRoundExhibitMark(trimmed) : !generic.test(trimmed)) return undefined;
  return { raw, sourceRaw, index, endIndex, ordinal, contextText, citationResolution: "none" };
}

function findBracketGroups(text: string) {
  const groups: Array<{ sourceRaw: string; inner: string; index: number; endIndex: number; open: string; close: string }> = [];
  for (const [open, close] of [["[", "]"], ["(", ")"]] as const) {
    const pattern = new RegExp(`${open === "[" ? "\\[" : "\\("}([^\\[\\]()]{1,240})${close === "]" ? "\\]" : "\\)"}`, "g");
    for (const match of text.matchAll(pattern)) {
      groups.push({
        sourceRaw: match[0],
        inner: match[1],
        index: match.index ?? 0,
        endIndex: (match.index ?? 0) + match[0].length,
        open,
        close,
      });
    }
  }
  groups.sort((left, right) => left.index - right.index || left.endIndex - right.endIndex);
  return groups;
}

/**
 * Returns every recognised exhibit token in reading order. Semicolon or comma
 * groups such as [AH-xx; AH-xx] or (BB-xx, BBxx, BB-x) yield independent
 * candidates when every part is a token. A legal page range such as
 * [AH1/12-18] remains one token.
 */
export function parseStatementCitationTokens(text: string): ParsedStatementCitationToken[] {
  const found: ParsedStatementCitationToken[] = [];
  for (const bracket of findBracketGroups(text)) {
    const contextText = tokenContext(text, bracket.index, bracket.endIndex);
    const parts = bracket.inner.split(/[;,]/).map((part) => part.trim()).filter(Boolean);
    const values = parts.length > 1 && parts.every((part) => Boolean(parseBracketToken(part, bracket.sourceRaw, bracket.index, bracket.endIndex, 0, contextText, bracket.open, bracket.close)))
      ? parts
      : [bracket.inner];
    for (const value of values) {
      const parsed = parseBracketToken(value, bracket.sourceRaw, bracket.index, bracket.endIndex, found.length + 1, contextText, bracket.open, bracket.close);
      if (parsed) found.push(parsed);
    }
  }
  return found;
}

/** Backwards-compatible convenience wrapper for callers that need one token. */
export function parseStatementCitationToken(text: string) {
  const token = parseStatementCitationTokens(text)[0];
  if (!token) return undefined;
  return {
    raw: token.raw,
    exhibitInitials: token.exhibitInitials,
    exhibitSequence: token.exhibitSequence,
    requestedExhibitPageStart: token.requestedExhibitPageStart,
    requestedExhibitPageEnd: token.requestedExhibitPageEnd,
    citationResolution: token.citationResolution,
  };
}

type CitationJob = {
  span: string;
  tokens: Array<ParsedStatementCitationToken | undefined>;
  citationCount?: number;
  citationOrdinal?: number;
  preferNounPhrase?: boolean;
  bindDateToSpan?: boolean;
};

function documentNounPhrase(span: string, nounSource: string) {
  const modifier = String.raw`(?:(?!and\b|to\b|of\b|for\b|with\b|from\b|the\b|a\b|an\b)[A-Za-z][A-Za-z'-]*\s+)?`;
  const pattern = new RegExp(`\\b(?:(?:the|a|an)\\s+)?(${modifier}${nounSource})\\b`, "i");
  const match = span.match(pattern);
  return match?.[1]?.replace(/\s+/g, " ").trim() || undefined;
}

function coordinatedDocumentAfterToken(span: string, tokens: ParsedStatementCitationToken[]) {
  const last = tokens[tokens.length - 1];
  if (!last) return undefined;
  const after = span.slice(last.endIndex);
  const andMatch = after.match(COORDINATED_AND_TO_THE);
  if (!andMatch || andMatch.index === undefined) return undefined;
  const clause = after.slice(andMatch.index);
  if (!new RegExp(`\\b${COORDINATED_DOCUMENT_NOUN}\\b`, "i").test(clause)) return undefined;
  const splitAt = last.endIndex + andMatch.index;
  const headSpan = span.slice(0, splitAt).replace(/\s+$/, "");
  const tailSpan = span.slice(splitAt).trim();
  if (!headSpan || !tailSpan) return undefined;
  return { headSpan, tailSpan };
}

function createCandidate(
  provisionalNumber: number,
  paragraph: StatementParagraph,
  context: string,
  discoverySignals: string[],
  reviewContext?: string,
  contextParagraphs?: ExhibitCandidate["contextParagraphs"],
  token?: ParsedStatementCitationToken,
  tokenCount = 1,
  span?: string,
  spanOptions?: { preferNounPhrase?: boolean; bindDateToSpan?: boolean; citationOrdinal?: number },
): ExhibitCandidate {
  const mark = formatExhibitMark(provisionalNumber);
  const candidateContext = spanOptions?.preferNounPhrase
    ? span ?? token?.contextText ?? paragraph.text
    : token?.contextText ?? paragraph.text;
  const directDates = extractDates(candidateContext);
  const directDate = tokenCount > 1 && token?.ordinal && directDates[token.ordinal - 1]
    ? directDates[token.ordinal - 1]
    : extractDate(candidateContext);
  const date =
    spanOptions?.bindDateToSpan || directDate !== "Date not stated" ? directDate : extractDate(context);
  const groupedSampleDescriptions = ["Sample project report", "Sample claimant email", "Sample cost workbook"];
  const nounPhrase = spanOptions?.preferNounPhrase ? documentNounPhrase(candidateContext, COORDINATED_DOCUMENT_NOUN) : undefined;
  const description = tokenCount >= 3 && /sample project report/i.test(candidateContext) && /sample claimant email/i.test(candidateContext) && /sample cost workbook/i.test(candidateContext)
    ? groupedSampleDescriptions[(token?.ordinal ?? 1) - 1] ?? suggestDescription(candidateContext, date)
    : nounPhrase ?? suggestDescription(candidateContext, date);
  return {
    id: `provisional-${provisionalNumber}-paragraph-${paragraph.number}${tokenCount > 1 ? `-reference-${token?.ordinal ?? spanOptions?.citationOrdinal ?? 1}` : ""}`,
    mark,
    provisionalNumber,
    description,
    date,
    paragraph: paragraph.number,
    citation: paragraph.text,
    citationToken: token?.raw,
    citationOrdinal: token?.ordinal ?? spanOptions?.citationOrdinal,
    citationCount: tokenCount,
    exhibitInitials: token?.exhibitInitials,
    exhibitSequence: token?.exhibitSequence,
    requestedExhibitPageStart: token?.requestedExhibitPageStart,
    requestedExhibitPageEnd: token?.requestedExhibitPageEnd,
    citationResolution: token?.citationResolution ?? "none",
    context: reviewContext,
    contextParagraphs,
    discoverySignals,
    evidenceId: null,
    confidence: 0,
    rationale: "No source file matched",
    included: true,
    confirmed: false,
  };
}

function discoverCitationSignals(text: string) {
  const signals: string[] = [];
  if (/\bI refer to\b/i.test(text)) signals.push('"I refer to" language');
  if (/\b(?:is|are)\s+(?:exhibited|at exhibit)\b/i.test(text)) {
    signals.push('"is/are exhibited" language');
  }
  if (
    /\bcopy of\b.{0,220}\b(?:is|are)\s+(?:at\s+)?exhibit(?:ed)?\b/i.test(text)
  ) {
    signals.push('"copy ... at exhibit" language');
  }
  if (hasDocumentNoun(text) && /\b(?:exhibit|exhibited|marked)\b/i.test(text)) {
    signals.push("document and exhibit language");
  }
  if (isGovernedAttachCitation(text)) {
    signals.push("attach or enclose language");
  }
  if (/\bAH\s*(?:\d{1,3}|xx)\b/i.test(text) || hasExhibitPlaceholder(text)) {
    signals.push("statement exhibit placeholder or mark");
  }
  return Array.from(new Set(signals));
}

function isNarrativeExhibitCitation(text: string) {
  if (!hasDocumentNoun(text)) return false;
  return (
    /\bI refer to\b/i.test(text) ||
    /\b(?:is|are)\s+(?:exhibited|at exhibit)\b/i.test(text) ||
    /\b(?:exhibited|marked)\s+(?:together\s+)?(?:to this statement\s+)?(?:as|at)\b/i.test(text) ||
    /\bcopy of\b.{0,220}\b(?:exhibit|exhibited)\b/i.test(text)
  );
}

function isGovernedAttachCitation(text: string) {
  return splitBoundedClauses(text).some((clause) => {
    if (new RegExp(`\\b${DOCUMENT_NOUN}\\b\\s+(?:is|are|was|were)\\s+(?:hereby\\s+)?(?:attached|enclosed|appended)(?:\\s+(?:hereto|herewith|to this statement))?(?:[.,;:]|$)`, "i").test(clause)) {
      return true;
    }
    const active = clause.match(/\b(?:I|we)\s+(?:hereby\s+)?(?:attach|enclose|append)(?:ed|ing)?\b(.*)$/i);
    if (!active) return false;
    const object = active[1] ?? "";
    const nounMatch = object.match(new RegExp(`\\b${DOCUMENT_NOUN}\\b`, "i"));
    if (!nounMatch || nounMatch.index === undefined) return false;
    const beforeNoun = object.slice(0, nounMatch.index);
    if (ATTACH_OBJECT_BLOCK.test(beforeNoun)) return false;
    if (/\b(?:to|for)\s+$/i.test(beforeNoun) && !/\b(?:copy of|set of)\s+$/i.test(beforeNoun)) return false;
    return true;
  });
}

function isExhibitCitation(text: string) {
  if (parseStatementCitationTokens(text).length > 0) return true;
  if (isNarrativeExhibitCitation(text)) return true;
  if (isGovernedAttachCitation(text)) return true;
  return /\bAH\s*xx\b/i.test(text);
}

function citationJobsForParagraph(text: string) {
  const spans = splitBoundedClauses(text);
  const working = spans.length ? spans : [text];
  const paragraphTokens = parseStatementCitationTokens(text);
  const jobs: CitationJob[] = [];
  if (paragraphTokens.length) {
    for (const span of working) {
      const parsedTokens = parseStatementCitationTokens(span);
      if (parsedTokens.length) {
        const extra = coordinatedDocumentAfterToken(span, parsedTokens);
        if (extra) {
          const citationCount = parsedTokens.length + 1;
          const headTokens = parsedTokens.map((token) => ({
            ...token,
            contextText: tokenContext(extra.headSpan, token.index, token.endIndex),
          }));
          jobs.push({
            span: extra.headSpan,
            tokens: headTokens,
            citationCount,
            preferNounPhrase: parsedTokens.length === 1,
            bindDateToSpan: true,
          });
          jobs.push({ span: extra.tailSpan, tokens: [undefined], citationCount, citationOrdinal: citationCount, preferNounPhrase: true, bindDateToSpan: true });
        } else {
          jobs.push({ span, tokens: parsedTokens });
        }
        continue;
      }
      if (isGovernedAttachCitation(span)) jobs.push({ span, tokens: [undefined] });
    }
    if (!jobs.length) jobs.push({ span: text, tokens: paragraphTokens });
    return jobs;
  }
  for (const span of working) {
    if (isNarrativeExhibitCitation(span) || isGovernedAttachCitation(span) || /\bAH\s*xx\b/i.test(span)) {
      jobs.push({ span, tokens: [undefined] });
    }
  }
  if (!jobs.length && isExhibitCitation(text)) jobs.push({ span: text, tokens: [undefined] });
  return jobs;
}

async function toEvidence(file: File, index: number): Promise<EvidenceRecord> {
  assertInputFileSize(file);
  let text: string;
  let pageCount = 0;
  let rotationPages: number[] = [];
  let annotationPages: number[] = [];
  let unsafePdfActions: UnsafePdfAction[] = [];
  let pageSizes: EvidenceRecord["pageSizes"] = [];
  let encrypted = false;
  let ocrPages: Array<{ text: string; confidence: number }> = [];
  let ocrStatus: EvidenceRecord["ocrStatus"] = "not-needed";
  let ocrFailureReason: string | undefined;
  let workbook: WorkbookAnalysis | undefined;
  let parsedEmail: Awaited<ReturnType<typeof parseBundleEmail>> | undefined;
  try {
    if (extensionOf(file.name) === "pdf") {
      const pdf = await extractPdfText(file);
      text = pdf.text;
      pageCount = pdf.pageCount;
      rotationPages = pdf.rotationPages;
      annotationPages = pdf.annotationPages;
      unsafePdfActions = pdf.unsafePdfActions;
      pageSizes = pdf.pageSizes;
      const needsOcr = pdf.pageTexts.map((page) => page.trim().length < 10);
      ocrPages = pdf.pageTexts.map(() => ({ text: "", confidence: 100 }));
      if (needsOcr.some(Boolean) && typeof document !== "undefined") {
        try {
          const recognised = (await ocrPdfLocally(file)).pages;
          ocrPages = recognised.map((page, pageIndex) => needsOcr[pageIndex] ? page : { text: "", confidence: 100 });
          text = pdf.pageTexts.map((page, pageIndex) => needsOcr[pageIndex] ? ocrPages[pageIndex]?.text ?? "" : page).join("\n").trim();
          ocrStatus = needsOcr.every((needed, pageIndex) => !needed || (ocrPages[pageIndex]?.text.length ?? 0) >= 1) ? "completed" : "failed";
        } catch (caught) {
          ocrStatus = "failed";
          ocrFailureReason = caught instanceof Error ? caught.message : "Local OCR stopped for an unknown reason.";
        }
      } else if (needsOcr.some(Boolean)) {
        ocrStatus = "unavailable";
      }
    } else if (extensionOf(file.name) === "xlsx") {
      workbook = await analyseXlsxInWorker(file);
      text = workbook.sheets.map((sheet) => `${sheet.name} ${sheet.cells.map((cell) => cell.value).join(" ")}`).join("\n");
    } else if (extensionOf(file.name) === "eml") {
      parsedEmail = await parseBundleEmail(await file.text());
      text = parsedEmail.body;
    } else {
      text = await extractFileText(file);
    }
  } catch (caught) {
    const message =
      caught instanceof Error ? caught.message : "Unknown document read error";
    if (/password|encrypted|PasswordException/i.test(message)) {
      text = "";
      encrypted = true;
    } else {
      throw new Error(`Could not read "${file.name}": ${message}`);
    }
  }
  const hash = await sha256(await file.arrayBuffer());
  let emailAttachments: EmailAttachmentChild[] | undefined;
  if (parsedEmail) {
    emailAttachments = await rederiveEmailChildren(file, hash);
    for (const child of emailAttachments) {
      if (child.supported) assertInputFileSize(child.file, child.extension);
      if (child.extension === "xlsx") {
        const childWorkbook = await analyseXlsxInWorker(child.file);
        child.workbook = childWorkbook;
        child.sheetSelections = childWorkbook.sheets.map((sheet) => ({ name: sheet.name, included: sheet.state === "visible", range: sheet.range }));
      }
      if (child.extension === "pdf") {
        try {
          child.pageSizes = (await extractPdfText(child.file)).pageSizes;
        } catch {
          // Page-size review is skipped when the child PDF cannot be measured.
        }
      }
    }
  }
  const markerMatch = text.match(/\bEXHIBIT\s+(AH\s*\d{1,3}|N\/A)\b/i);
  const rawMarker = markerMatch?.[1]?.toUpperCase() ?? null;
  return {
    id: `evidence-${index}-${file.name}`,
    file,
    name: file.name,
    extension: extensionOf(file.name),
    text,
    marker:
      rawMarker && rawMarker !== "N/A"
        ? normalizeExhibitMark(rawMarker)
        : rawMarker,
    // Hash only after parsing has released its temporary buffers. Keeping a
    // second full copy here doubled peak memory for large PDFs.
    sha256: hash,
    pageCount,
    readableText: text.trim().length >= 10,
    encrypted,
    rotationPages,
    annotationPages,
    unsafePdfActions,
    ocrPages,
    ocrStatus,
    ocrFailureReason,
    pageSizes,
    workbook,
    sheetSelections: workbook?.sheets.map((sheet) => ({ name: sheet.name, included: sheet.state === "visible", range: sheet.range })),
    emailAttachments,
  };
}

/** Reads additional local evidence with the same safeguards as initial analysis. */
export async function analyseEvidenceFiles(files: File[], startingIndex = 0, onProgress?: (stage: string, detail?: string) => void) {
  const records: EvidenceRecord[] = [];
  const usable = files.filter((file) => !/^~(?:\$|%24)/i.test(file.name));
  if (usable.length > INPUT_LIMITS.evidenceFiles) throw new Error(`Choose no more than ${INPUT_LIMITS.evidenceFiles} evidence files in one project.`);
  for (const [index, file] of usable.entries()) {
    onProgress?.("Reading evidence files", `${index + 1} of ${usable.length}: ${file.name}`);
    records.push(await toEvidence(file, startingIndex + index));
  }
  const expanded = records.reduce((total, record) => total + 1 + (record.emailAttachments?.length ?? 0), 0);
  if (expanded > INPUT_LIMITS.evidenceFiles) {
    throw new Error(`Choose no more than ${INPUT_LIMITS.evidenceFiles} evidence files in one project, including email attachments.`);
  }
  return records;
}

/** Rebuilds in-memory attachment exhibits from the parent EML. Child bytes are never stored in the project archive. */
export async function attachDerivedEmailEvidence(analysis: AnalysisResult, candidates: ExhibitCandidate[]) {
  const extra: EvidenceRecord[] = [];
  const nextCandidates: ExhibitCandidate[] = [];
  for (const candidate of candidates) {
    const provenance = candidate.parentEmailProvenance;
    if (!provenance) {
      nextCandidates.push(candidate);
      continue;
    }
    const parent = analysis.evidence.find((record) => record.sha256 === provenance.parentSha256)
      ?? analysis.evidence.find((record) => record.emailAttachments?.some((child) => child.identity === provenance.childIdentity));
    const child = parent?.emailAttachments?.find((item) => item.identity === provenance.childIdentity && item.sha256 === provenance.childSha256);
    if (!child) {
      nextCandidates.push({ ...candidate, evidenceId: null, confirmed: false, confirmationMethod: undefined, confirmedAt: undefined });
      continue;
    }
    let record = [...analysis.evidence, ...extra].find((item) => item.derivedFromEmail?.childIdentity === child.identity);
    if (!record) {
      record = await toEvidence(child.file, analysis.evidence.length + extra.length);
      record.derivedFromEmail = { parentSha256: provenance.parentSha256, childIdentity: child.identity };
      extra.push(record);
    }
    nextCandidates.push({ ...candidate, evidenceId: record.id });
  }
  if (!extra.length) return { analysis, candidates: nextCandidates };
  return {
    analysis: { ...analysis, evidence: [...analysis.evidence, ...extra] },
    candidates: nextCandidates,
  };
}

export async function analyseFiles(
  statement: File,
  evidenceFiles: File[],
  onProgress?: (stage: string, detail?: string) => void,
): Promise<AnalysisResult> {
  if (extensionOf(statement.name) !== "docx") {
    throw new Error("Exhibit Builder currently extracts citations from DOCX witness statements.");
  }

  onProgress?.("Reading the witness statement", statement.name);
  const { raw, paragraphs } = await extractStatementParagraphs(statement);
  const candidates: ExhibitCandidate[] = [];
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    const jobs = citationJobsForParagraph(paragraph.text);
    if (!jobs.length) continue;
    const nearby = paragraphs
      .slice(Math.max(0, paragraphIndex - 1), paragraphIndex + 2);
    const context = nearby.map((item) => item.text).join(" ");
    const reviewContext = nearby
      .filter((item) => item.number !== paragraph.number)
      .map((item) => `${item.number < paragraph.number ? "Previous" : "Following"} paragraph ${item.number}: ${item.text}`)
      .join(" ");
    const contextParagraphs = nearby
      .filter((item) => item.number !== paragraph.number)
      .map((item) => ({
        paragraph: item.number,
        position: item.number < paragraph.number ? "previous" as const : "following" as const,
        text: item.text,
      }));
    for (const job of jobs) {
      for (const token of job.tokens) {
        candidates.push(
          createCandidate(
            candidates.length + 1,
            paragraph,
            context,
            discoverCitationSignals(job.span),
            reviewContext || undefined,
            contextParagraphs.length ? contextParagraphs : undefined,
            token,
            job.citationCount ?? job.tokens.length,
            job.span,
            {
              preferNounPhrase: job.preferNounPhrase,
              bindDateToSpan: job.bindDateToSpan,
              citationOrdinal: job.citationOrdinal,
            },
          ),
        );
      }
    }
  }

  // OCR owns a PDF.js renderer and a Tesseract worker.  Starting one per
  // evidence file exhausts packaged renderer resources and can stall analysis.
  // Keep parsing deterministic and local by admitting one evidence file at a
  // time; text-native PDFs remain fast and scanned PDFs still receive full OCR.
  const evidence = await analyseEvidenceFiles(evidenceFiles, 0, onProgress);
  const evidenceProfiles = new Map(evidence.map((record) => [record.id, evidenceMatchProfile(record)]));
  const claimed = new Set<string>();
  for (const [candidateIndex, candidate] of candidates.entries()) {
    onProgress?.("Comparing statement references", `${candidateIndex + 1} of ${candidates.length}: paragraph ${candidate.paragraph}`);
    const candidateTarget = tokens(`${candidate.description} ${candidate.citation} ${candidate.date}`);
    if ((candidate.citationCount ?? 1) > 1) {
      const ranked = evidence
        .filter((record) => record.marker !== "N/A")
        .map((record) => ({ record, score: fallbackScore(candidate, record, evidenceProfiles.get(record.id), candidateTarget) }))
        .sort((left, right) => right.score - left.score);
      candidate.confidence = ranked[0]?.score ?? 0;
      candidate.alternativeEvidenceIds = ranked.slice(0, 4).map((item) => item.record.id);
      candidate.rationale = `Reference ${candidate.citationOrdinal ?? 1} of ${candidate.citationCount} in this paragraph requires a reviewer to choose its source file`;
      continue;
    }
    const explicit = evidence.filter(
      (record) =>
        record.marker === normalizeExhibitMark(candidate.mark) &&
        !claimed.has(record.id) && candidate.citationResolution === "none",
    );
    if (explicit.length === 1) {
      candidate.evidenceId = explicit[0].id;
      candidate.confidence = 99;
      candidate.rationale =
        "Existing source label agrees with the provisional number - human confirmation required";
      claimed.add(explicit[0].id);
      continue;
    }

    const ranked = evidence
      .filter((record) => !claimed.has(record.id) && record.marker !== "N/A")
      .map((record) => ({ record, score: fallbackScore(candidate, record, evidenceProfiles.get(record.id), candidateTarget) }))
      .sort((left, right) => right.score - left.score);

    if (ranked[0] && ranked[0].score >= 45) {
      const ambiguous =
        ranked[1] && Math.abs(ranked[0].score - ranked[1].score) < 8;
      candidate.evidenceId = ambiguous ? null : ranked[0].record.id;
      candidate.confidence = ranked[0].score;
      candidate.alternativeEvidenceIds = ranked.slice(1, 4).map((item) => item.record.id);
      candidate.rationale = ambiguous
        ? "Two files have similar dates and wording - human selection required"
        : "Matched using date, title and citation wording";
      if (!ambiguous) {
        claimed.add(ranked[0].record.id);
      }
    }
  }

  const unreferenced = evidence.filter((record) => !claimed.has(record.id));
  const statementBytes = await statement.arrayBuffer();
  const statementText = raw.join(" ");
  const statementWarnings: string[] = [];
  if (/\bAH\s*xx\b/i.test(statementText) || hasExhibitPlaceholder(statementText)) {
    statementWarnings.push(
      "The statement contains exhibit placeholders. Provisional numbers are suggestions only.",
    );
  }
  const residualMarks = Array.from(
    new Set(
      Array.from(statementText.matchAll(/\bAH\s*\d{1,3}\b/gi)).map((match) =>
        normalizeExhibitMark(match[0]),
      ),
    ),
  );
  if (residualMarks.length) {
    statementWarnings.push(
      `Existing numeric references (${residualMarks.join(", ")}) were not used to control provisional numbering.`,
    );
  }

  return {
    statementName: statement.name,
    statementHash: await sha256(statementBytes),
    caseTitle: inferCaseTitle(raw.join(" ")),
    candidates,
    evidence,
    unreferenced,
    statementWarnings,
    generatedAt: new Date().toISOString(),
  };
}

/** Analyses one witness statement without editing the source document. */
export async function analyseBundleStatements(
  statements: BundleStatementInput[],
  evidenceFiles: File[],
  onProgress?: (stage: string, detail?: string) => void,
): Promise<AnalysisResult> {
  if (statements.length !== 1) throw new Error("Add one witness statement for this exhibit bundle.");
  const individual: AnalysisResult[] = [];
  for (const statement of statements) {
    const result = await analyseFiles(statement.file, evidenceFiles, onProgress);
    const configuredInitials = statement.witnessInitials.trim().toUpperCase() || "EX";
    const explicitIdentities = new Map<string, { initials: string; sequence: number }>();
    for (const candidate of result.candidates) {
      if (candidate.citationResolution === "none" || !candidate.exhibitInitials || !candidate.exhibitSequence) continue;
      const initials = candidate.exhibitInitials.replace(/\s+/g, "").toUpperCase();
      const sequence = candidate.exhibitSequence;
      explicitIdentities.set(`${initials}:${sequence}`, { initials, sequence });
    }
    // One explicit full reference such as [LV1/xx] establishes the mark for
    // this single witness exhibit bundle. Incomplete placeholders such as
    // [LV-xx] then inherit it instead of silently creating a second VA1 bundle.
    // Conflicting complete marks remain visible for reviewer resolution.
    const soleExplicitIdentity = explicitIdentities.size === 1 ? [...explicitIdentities.values()][0] : undefined;
    const displayInitials = soleExplicitIdentity?.initials ?? configuredInitials;
    const displaySequence = soleExplicitIdentity?.sequence ?? 1;
    const witnessKey = `${statement.witnessName.trim().toLowerCase()}::${displayInitials}`;
    result.candidates = result.candidates.map((candidate, index) => {
      return {
        ...candidate,
        id: `${statement.id}:${candidate.id}`,
        provisionalNumber: index + 1,
        mark: `${displayInitials} ${index + 1}`,
        // A token's number is the bundle number.  In the absence of a token,
        // suggest this witness's first bundle and fill its pages after build.
        exhibitInitials: soleExplicitIdentity ? displayInitials : candidate.exhibitInitials ?? displayInitials,
        exhibitSequence: soleExplicitIdentity ? displaySequence : candidate.exhibitSequence ?? displaySequence,
        statementName: statement.file.name,
        witnessInitials: displayInitials,
        witnessKey,
        statementId: statement.id,
        sequenceOrder: individual.length * 10000 + index,
      };
    });
    individual.push(result);
  }
  const first = individual[0];
  const sourceStatement = statements[0];
  return {
    ...first,
    statementName: first.statementName,
    // One project accepts one statement, so retain the source file's actual
    // SHA-256 instead of hashing its hexadecimal digest a second time.
    statementHash: first.statementHash,
    candidates: individual.flatMap((item) => item.candidates),
    statementWarnings: individual.flatMap((item) => item.statementWarnings),
    generatedAt: new Date().toISOString(),
    statementId: sourceStatement.id,
    witnessName: sourceStatement.witnessName,
    witnessInitials: first.candidates[0]?.witnessInitials ?? sourceStatement.witnessInitials,
  };
}

/**
 * Applies reviewer-edited witness metadata without reopening either the
 * statement or any evidence file. Source extraction, hashes, OCR results,
 * workbook analysis, matches and approvals remain untouched.
 */
export function applyWitnessDetails(
  analysis: AnalysisResult,
  candidates: ExhibitCandidate[],
  statement: BundleStatementInput,
) {
  const displayInitials = statement.witnessInitials.trim().toUpperCase() || "EX";
  const witnessKey = `${statement.witnessName.trim().toLowerCase()}::${displayInitials}`;
  const updatedCandidates = candidates.map((candidate) => {
    if (candidate.statementId !== statement.id || candidate.manualAddition) return candidate;
    return {
      ...candidate,
      mark: `${displayInitials} ${candidate.provisionalNumber}`,
      exhibitInitials: displayInitials,
      witnessInitials: displayInitials,
      witnessKey,
    };
  });
  const updatedAnalysis = {
    ...analysis,
    candidates: analysis.candidates.map((candidate) => updatedCandidates.find((updated) => updated.id === candidate.id) ?? candidate),
    ...(analysis.statementName === statement.file.name ? { witnessName: statement.witnessName.trim(), witnessInitials: displayInitials } : {}),
  };
  return { analysis: updatedAnalysis, candidates: updatedCandidates };
}

async function fetchSampleFile(name: string) {
  const response = await fetch(`/guided-sample/${encodeURIComponent(name)}`);
  if (!response.ok || /text\/html/i.test(response.headers.get("content-type") ?? "")) {
    throw new Error("The optional guided sample is unavailable. You can still build a bundle from your own files.");
  }
  const blob = await response.blob();
  return new File([blob], name, {
    type: response.headers.get("content-type") ?? blob.type,
  });
}

export async function checkSamplePackAvailability() {
  try {
    const responses = await Promise.all(SAMPLE_REQUIRED_FILES.map((name) => fetch(`/guided-sample/${encodeURIComponent(name)}`, {
      method: "HEAD",
      cache: "no-store",
    })));
    return responses.every((response) => response.ok && !/text\/html/i.test(response.headers.get("content-type") ?? ""));
  } catch {
    return false;
  }
}

export async function loadSamplePack() {
  const [statement, evidence] = await Promise.all([
    fetchSampleFile(SAMPLE_STATEMENT),
    Promise.all(SAMPLE_EVIDENCE.map(fetchSampleFile)),
  ]);
  return {
    statement,
    evidence,
  };
}

function wrapText(text: string, font: PDFFont, size: number, maxWidth: number) {
  const safeText = sanitizePdfText(text).trim();
  if (!safeText) return [];
  const splitOversizeToken = (token: string) => {
    if (font.widthOfTextAtSize(token, size) <= maxWidth) return [token];
    const parts: string[] = [];
    let remaining = token;
    while (remaining && font.widthOfTextAtSize(remaining, size) > maxWidth) {
      let fit = "";
      let preferredBreak = -1;
      for (const character of remaining) {
        const next = fit + character;
        if (font.widthOfTextAtSize(next, size) > maxWidth) break;
        fit = next;
        if (/[_\-/.]/.test(character)) preferredBreak = fit.length;
      }
      if (!fit) fit = remaining[0];
      const breakAt = preferredBreak > 0 && preferredBreak >= Math.floor(fit.length * 0.55) ? preferredBreak : fit.length;
      parts.push(remaining.slice(0, breakAt));
      remaining = remaining.slice(breakAt);
    }
    if (remaining) parts.push(remaining);
    return parts;
  };
  const words = safeText.split(/\s+/).flatMap(splitOversizeToken);
  const lines: string[] = [];
  let line = "";
  for (const word of words) {
    const next = line ? `${line} ${word}` : word;
    if (font.widthOfTextAtSize(next, size) <= maxWidth) {
      line = next;
    } else {
      if (line) lines.push(line);
      line = word;
    }
  }
  if (line) lines.push(line);
  return lines;
}

function drawTextInsideBox(
  page: PDFPage,
  text: string,
  options: {
    font: PDFFont;
    maxSize: number;
    minSize: number;
    x: number;
    y: number;
    maxWidth: number;
    maxHeight: number;
    color?: ReturnType<typeof rgb>;
  },
) {
  for (let size = options.maxSize; size >= options.minSize; size -= 1) {
    const lineHeight = size * 1.28;
    const lines = wrapText(text, options.font, size, options.maxWidth);
    if (lines.length * lineHeight > options.maxHeight && size > options.minSize) continue;
    let y = options.y;
    for (const line of lines) {
      page.drawText(line, { x: options.x, y, size, font: options.font, color: options.color ?? rgb(0.12, 0.15, 0.19) });
      y -= lineHeight;
    }
    return y;
  }
  return options.y;
}

function drawWrapped(
  page: PDFPage,
  text: string,
  options: {
    font: PDFFont;
    size: number;
    x: number;
    y: number;
    maxWidth: number;
    lineHeight: number;
    color?: ReturnType<typeof rgb>;
  },
) {
  const lines = wrapText(text, options.font, options.size, options.maxWidth);
  let y = options.y;
  for (const line of lines) {
    page.drawText(line, {
      x: options.x,
      y,
      size: options.size,
      font: options.font,
      color: options.color ?? rgb(0.12, 0.15, 0.19),
    });
    y -= options.lineHeight;
  }
  return y;
}

function indexDescription(record: BundleRecord) {
  // The index identifies exhibits and their bundle pages.  Statement-reference
  // detail remains in the manifest/audit data and the update suggestions, not
  // in the reader-facing index.
  return record.description;
}

function completeIndexRows(
  indexNodes: readonly BuildPlanIndexNode[],
  recordsById: ReadonlyMap<string, { record: BundleRecord; indexPageLabelStart: string; indexPageLabelEnd: string }>,
  localItemIds: ReadonlySet<string>,
  multiVolume: boolean,
): IndexLayoutRowInput[] {
  const exhibitRow = (itemId: string, precedingGroupBreak = false): IndexLayoutRowInput => {
    const entry = recordsById.get(itemId);
    if (!entry) throw new Error(`The complete index could not find planned exhibit ${itemId}.`);
    const { record, indexPageLabelStart: start, indexPageLabelEnd: end } = entry;
    const range = start === end ? start : `${start}-${end}`;
    return {
      kind: "exhibit",
      id: `exhibit:${itemId}`,
      exhibitLabel: String(record.exhibitNumber ?? "?"),
      description: sanitizePdfText(indexDescription(record)),
      date: sanitizePdfText(record.documentDate?.trim() || "Date not stated"),
      pageLabel: sanitizePdfText(multiVolume ? `Vol. ${record.volumeNumber ?? 1} / ${range}` : range),
      ...(localItemIds.has(itemId) ? { linkTargetId: itemId } : {}),
      ...(precedingGroupBreak ? { precedingGroupBreak: true } : {}),
    };
  };
  const rows: IndexLayoutRowInput[] = [];
  let groupBreakBeforeNextUnheaded = false;
  for (const node of indexNodes) {
    if (node.kind !== "section") {
      rows.push(exhibitRow(node.itemId, groupBreakBeforeNextUnheaded));
      groupBreakBeforeNextUnheaded = false;
      continue;
    }
    if (!node.itemIds.length) continue;
    rows.push({ kind: "section", id: `section:${node.id}`, title: sanitizePdfText(node.title) });
    rows.push(...node.itemIds.map((itemId) => exhibitRow(itemId)));
    groupBreakBeforeNextUnheaded = true;
  }
  return rows;
}

function createAuthoritativeIndexLayout(
  rows: readonly IndexLayoutRowInput[],
  geometry: "built-in" | "custom-template" | IndexLayoutGeometryProfile,
  regular: PDFFont,
  bold: PDFFont,
): IndexLayoutPlan {
  const result = createIndexLayoutPlan({
    rows,
    geometry,
    measureText: (text, size, role) => (role === "description" || role === "date" ? regular : bold).widthOfTextAtSize(text, size),
  });
  if (!result.ok) {
    const suffix = result.error.pageLabel ? ` The page reference is "${result.error.pageLabel}".` : "";
    throw new Error(`Index layout blocked (${result.error.code}): ${result.error.message}${suffix}`);
  }
  return result.plan;
}

function drawAuthoritativeIndexPage(
  page: PDFPage,
  plan: IndexLayoutPlan,
  plannedPageNumber: number,
  regular: PDFFont,
  bold: PDFFont,
) {
  const pageHeight = page.getHeight();
  for (const row of plan.rows.filter((candidate) => candidate.pageNumber === plannedPageNumber)) {
    if (row.kind === "section") {
      if (plan.geometry.id === "built-in") {
        page.drawRectangle({
          x: row.bounds.x,
          y: pageHeight - row.bounds.top - row.bounds.height,
          width: row.bounds.width,
          height: row.bounds.height,
          color: rgb(0.95, 0.95, 0.94),
        });
      }
      for (const line of row.lines) {
        page.drawText(line.text, { x: line.x, y: pageHeight - line.baseline, size: line.fontSize, font: bold, color: rgb(0.12, 0.16, 0.22) });
      }
      continue;
    }
    if (plan.geometry.id === "built-in") {
      page.drawLine({
        start: { x: row.bounds.x, y: pageHeight - row.bounds.top - row.bounds.height },
        end: { x: row.bounds.x + row.bounds.width, y: pageHeight - row.bounds.top - row.bounds.height },
        thickness: 0.35,
        color: rgb(0.78, 0.8, 0.82),
      });
    }
    for (const line of [...row.exhibitLines, ...(row.dateLines ?? []), ...row.descriptionLines, ...row.pageReferenceLines]) {
      const font = line.role === "description" || line.role === "date" ? regular : bold;
      const color = line.role === "exhibit" ? rgb(0.56, 0.18, 0.12) : rgb(0.12, 0.16, 0.22);
      page.drawText(line.text, { x: line.x, y: pageHeight - line.baseline, size: line.fontSize, font, color });
    }
  }
}

export type StatementUpdateSuggestion = {
  paragraph: number | null;
  exhibit: string;
  pageRange: string | null;
  line: string;
  needsReview: boolean;
};

function formatExhibitPageReference(start: number | undefined, end: number | undefined) {
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end)) return null;
  return start === end ? `${start}` : `${start}-${end}`;
}

function isUncitedRecord(record: BundleRecord) {
  return Boolean(record.manualAddition || record.citationStatus === "not-cited-manual-addition");
}

function finishedUncitedPageRange(record: BundleRecord) {
  if (record.exhibitPageLabelStart && record.exhibitPageLabelEnd) {
    return {
      range: record.exhibitPageLabelStart === record.exhibitPageLabelEnd
        ? record.exhibitPageLabelStart
        : `${record.exhibitPageLabelStart}-${record.exhibitPageLabelEnd}`,
      single: record.exhibitPageLabelStart === record.exhibitPageLabelEnd,
    };
  }
  const numeric = formatExhibitPageReference(record.exhibitPageStart, record.exhibitPageEnd);
  if (!numeric) return { range: null, single: true };
  return { range: numeric, single: !numeric.includes("-") };
}

function pageTextEncodesVolume(value: string | null | undefined) {
  return Boolean(value && /Vol\.\s*\d+/i.test(value));
}

function formatUncitedExhibitLine(record: BundleRecord, multiVolume: boolean) {
  const { range, single } = finishedUncitedPageRange(record);
  const pageWord = single ? "page" : "pages";
  const alreadyNamed = pageTextEncodesVolume(range)
    || pageTextEncodesVolume(record.exhibitPageLabelStart)
    || pageTextEncodesVolume(record.exhibitPageLabelEnd);
  const volumePrefix = multiVolume && !alreadyNamed ? `Vol. ${record.volumeNumber ?? 1}, ` : "";
  const location = range ? `${volumePrefix}${pageWord} ${range}` : volumePrefix.replace(/, $/, "");
  if (!location) return `${record.exhibitNumber}. ${record.description}`;
  return `${record.exhibitNumber}. ${record.description} — ${location}`;
}

/** Local-only suggestions; source witness statements are never edited. */
export function buildStatementUpdateSuggestions(records: BundleRecord[]): StatementUpdateSuggestion[] {
  const multiVolume = records.some((record) => (record.volumeNumber ?? 1) > 1);
  const cited = records.flatMap((record) => {
    return record.statementReferences.map((reference) => {
      const initials = reference.exhibitInitials?.replace(/\s+/g, "") ?? "";
      const sequence = reference.exhibitSequence;
      const exhibit = sequence && initials ? `${initials}${sequence}` : "citation";
      const numericPageRange = Number.isSafeInteger(reference.exhibitPageStart) && Number.isSafeInteger(reference.exhibitPageEnd) && reference.exhibitPageStart! >= 1 && reference.exhibitPageEnd! >= reference.exhibitPageStart!
        ? formatExhibitPageReference(reference.exhibitPageStart, reference.exhibitPageEnd)
        : null;
      const pageRange = reference.exhibitPageLabelStart && reference.exhibitPageLabelEnd
        ? reference.exhibitPageLabelStart === reference.exhibitPageLabelEnd
          ? reference.exhibitPageLabelStart
          : `${reference.exhibitPageLabelStart}-${reference.exhibitPageLabelEnd}`
        : numericPageRange;
      const volumePrefix = multiVolume ? `Vol. ${reference.volumeNumber ?? record.volumeNumber ?? 1}/` : "";
      const validParagraph = Number.isInteger(reference.paragraph) && reference.paragraph > 0;
      const hasRequestedRange = Number.isSafeInteger(reference.requestedExhibitPageStart) && Number.isSafeInteger(reference.requestedExhibitPageEnd);
      const requestedRange = hasRequestedRange ? formatExhibitPageReference(reference.requestedExhibitPageStart, reference.requestedExhibitPageEnd) : null;
      const requestedConflict = Boolean(requestedRange && numericPageRange && requestedRange !== numericPageRange);
      const needsReview = requestedConflict || !pageRange || !sequence || !initials || !validParagraph;
      const displayRange = pageRange;
      const line = requestedConflict
        ? `${validParagraph ? `Paragraph ${reference.paragraph}` : "Statement paragraph"} - [${exhibit}/${volumePrefix}${pageRange}] (requested ${requestedRange}; review)`
        : needsReview
          ? `${validParagraph ? `Paragraph ${reference.paragraph}` : "Statement paragraph"} - [${exhibit}/xx]`
          : `Paragraph ${reference.paragraph} - [${exhibit}/${volumePrefix}${displayRange}]`;
      return { paragraph: validParagraph ? reference.paragraph : null, exhibit, pageRange, line, needsReview };
    });
  }).sort((left, right) => (left.paragraph ?? Number.MAX_SAFE_INTEGER) - (right.paragraph ?? Number.MAX_SAFE_INTEGER));
  const uncited = records
    .filter(isUncitedRecord)
    .slice()
    .sort((left, right) => (left.exhibitNumber ?? 0) - (right.exhibitNumber ?? 0))
    .map((record) => {
      const { range } = finishedUncitedPageRange(record);
      return {
        paragraph: null,
        exhibit: String(record.exhibitNumber ?? ""),
        pageRange: range,
        line: formatUncitedExhibitLine(record, multiVolume),
        needsReview: false,
      };
    });
  if (!uncited.length) return cited;
  return [
    ...cited,
    { paragraph: null, exhibit: "", pageRange: null, line: "Uncited exhibits — no statement reference", needsReview: false },
    ...uncited,
  ];
}

async function appendTextEvidence(
  output: PDFDocument,
  record: EvidenceRecord,
  regular: PDFFont,
  bold: PDFFont,
) {
  const width = A4_WIDTH;
  const height = A4_HEIGHT;
  const margin = 54;
  const paragraphs = sanitizePdfText(record.text)
    .split(/\n+/)
    .map(normalizeWhitespace)
    .filter(Boolean);
  let page = output.addPage([width, height]);
  let y = height - 62;
  y = drawWrapped(
    page,
    record.name.replace(/\.[^.]+$/, "").replace(/_/g, " "),
    {
    x: margin,
    y,
    size: 16,
    font: bold,
    maxWidth: width - margin * 2,
    lineHeight: 20,
    color: rgb(0.12, 0.16, 0.22),
    },
  );
  y -= 18;

  for (const paragraph of paragraphs) {
    const lines = wrapText(paragraph, regular, 9.5, width - margin * 2);
    const needed = lines.length * 13 + 12;
    if (y - needed < 54) {
      page = output.addPage([width, height]);
      y = height - 58;
    }
    for (const line of lines) {
      page.drawText(line, {
        x: margin,
        y,
        size: 9.5,
        font: regular,
        color: rgb(0.14, 0.16, 0.19),
      });
      y -= 13;
    }
    y -= 10;
  }
}

function drawEmailPageHeading(
  page: PDFPage,
  account: string,
  bold: PDFFont,
  continuation = false,
) {
  const margin = 46;
  const title = continuation ? `${account} - continued` : account;
  page.drawText(title, {
    x: margin,
    y: A4_HEIGHT - 48,
    size: 10,
    font: bold,
    color: rgb(0.06, 0.07, 0.08),
  });
  page.drawLine({
    start: { x: margin, y: A4_HEIGHT - 58 },
    end: { x: A4_WIDTH - margin, y: A4_HEIGHT - 58 },
    thickness: 1.6,
    color: rgb(0.06, 0.07, 0.08),
  });
  return A4_HEIGHT - 78;
}

function drawEmailHeaderRow(
  page: PDFPage,
  label: string,
  value: string,
  y: number,
  regular: PDFFont,
  bold: PDFFont,
) {
  const labelX = 46;
  const valueX = 108;
  const size = 9;
  const lineHeight = 11.5;
  const lines = wrapText(value || "-", regular, size, A4_WIDTH - valueX - 46);
  page.drawText(`${label}:`, {
    x: labelX,
    y,
    size,
    font: bold,
    color: rgb(0.06, 0.07, 0.08),
  });
  for (const [index, line] of lines.entries()) {
    page.drawText(line, {
      x: valueX,
      y: y - index * lineHeight,
      size,
      font: regular,
      color: rgb(0.06, 0.07, 0.08),
    });
  }
  return y - Math.max(1, lines.length) * lineHeight;
}

async function appendEmailChildEvidence(
  output: PDFDocument,
  child: EmailAttachmentChild,
  regular: PDFFont,
  bold: PDFFont,
  options: BuildOptions,
  resolutions: NonNullable<BuildOptions["resolutions"]>,
) {
  if (child.extension === "pdf") {
    await appendPdfEvidence(output, {
      id: child.identity,
      file: child.file,
      name: child.name,
      extension: "pdf",
      text: "",
      marker: null,
      sha256: child.sha256,
      pageCount: 0,
      readableText: false,
      encrypted: false,
      rotationPages: [],
      ocrPages: [],
      ocrStatus: "not-needed",
    }, undefined, undefined, true, options.pageSizeChoices?.[child.identity] ?? "convert-to-a4");
    return;
  }
  if (child.extension === "xlsx") {
    const record = await toEvidence(child.file, 0);
    record.sheetSelections = child.sheetSelections ?? record.sheetSelections;
    await appendXlsxEvidence(output, record, options.workbookExporter);
    return;
  }
  if (child.extension === "eml") {
    await appendEmailEvidence(output, {
      id: child.identity,
      file: child.file,
      name: child.name,
      extension: "eml",
      text: "",
      marker: null,
      sha256: child.sha256,
      pageCount: 0,
      readableText: true,
      encrypted: false,
      rotationPages: [],
      ocrPages: [],
      ocrStatus: "not-needed",
    }, regular, bold);
    return;
  }
  await appendTextEvidence(output, {
    id: child.identity,
    file: child.file,
    name: child.name,
    extension: child.extension,
    text: await extractFileText(child.file),
    marker: null,
    sha256: child.sha256,
    pageCount: 0,
    readableText: true,
    encrypted: false,
    rotationPages: [],
    ocrPages: [],
    ocrStatus: "not-needed",
  }, regular, bold);
}

async function appendEmailEvidence(
  output: PDFDocument,
  record: EvidenceRecord,
  regular: PDFFont,
  bold: PDFFont,
  printChildren: EmailAttachmentChild[] = [],
  options?: BuildOptions,
) {
  const parsed = await parseBundleEmail(await record.file.text());
  const headers = parsed.headers;
  const from = headers.get("from") ?? "";
  const account =
    from.match(/<([^<>@\s]+@[^<>@\s]+)>/)?.[1] ??
    from.match(/\b[^\s<>@]+@[^\s<>@]+\b/)?.[0] ??
    "Email correspondence";
  const margin = 46;
  const maxWidth = A4_WIDTH - margin * 2;
  const bodySize = 9.5;
  const bodyLineHeight = 12.5;
  const bottom = PAGE_FOOTER_TOP + 16;
  let page = output.addPage([A4_WIDTH, A4_HEIGHT]);
  let y = drawEmailPageHeading(page, account, bold);

  const headerRows = [
    ["From", headers.get("from") ?? ""],
    ["Sent", headers.get("date") ?? ""],
    ["To", headers.get("to") ?? ""],
    ...(headers.get("cc") ? [["Cc", headers.get("cc") ?? ""]] : []),
    ...(headers.get("bcc") ? [["Bcc", headers.get("bcc") ?? ""]] : []),
    ["Subject", headers.get("subject") ?? ""],
  ];
  for (const [label, value] of headerRows) {
    y = drawEmailHeaderRow(page, label, value, y, regular, bold);
  }
  if (parsed.attachments.length) {
    y = drawEmailHeaderRow(page, "Attachments", parsed.attachments.map((attachment) => `${attachment.name} (${attachment.disposition})`).join("; "), y, regular, bold);
  }
  y -= 4;
  page.drawLine({
    start: { x: margin, y },
    end: { x: A4_WIDTH - margin, y },
    thickness: 0.7,
    color: rgb(0.55, 0.57, 0.6),
  });
  y -= 18;

  for (const rawLine of parsed.body.split("\n")) {
    const line = sanitizePdfText(rawLine).trimEnd();
    if (!line.trim()) {
      y -= 8;
      continue;
    }
    if (/^-{12,}$/.test(line.trim())) {
      if (y < bottom + 10) {
        page = output.addPage([A4_WIDTH, A4_HEIGHT]);
        y = drawEmailPageHeading(page, account, bold, true) - 8;
      }
      page.drawLine({
        start: { x: margin, y },
        end: { x: A4_WIDTH - margin, y },
        thickness: 0.5,
        color: rgb(0.66, 0.68, 0.71),
      });
      y -= 10;
      continue;
    }

    const lines = wrapText(line, regular, bodySize, maxWidth);
    const needed = Math.max(1, lines.length) * bodyLineHeight;
    if (y - needed < bottom) {
      page = output.addPage([A4_WIDTH, A4_HEIGHT]);
      y = drawEmailPageHeading(page, account, bold, true) - 8;
    }
    for (const wrappedLine of lines) {
      page.drawText(wrappedLine, {
        x: margin,
        y,
        size: bodySize,
        font: regular,
        color: rgb(0.06, 0.07, 0.08),
      });
      y -= bodyLineHeight;
    }
  }
  if (printChildren.length) {
    for (const child of printChildren) {
      if (!child.supported) throw new Error(`${record.name} attachment ${child.name} cannot be printed. Leave it out or add a supported PDF, DOCX, EML or XLSX file.`);
      await appendEmailChildEvidence(output, child, regular, bold, options ?? {}, options?.resolutions ?? []);
    }
  }
}

async function appendPdfEvidence(
  output: PDFDocument,
  record: EvidenceRecord,
  pageStart?: number,
  pageEnd?: number,
  includeGeneratedOcr = true,
  pageSizeHandling: NonA4PageHandling = "convert-to-a4",
) {
  const sourceBytes = new Uint8Array(await record.file.arrayBuffer());
  const source = await PDFDocument.load(sourceBytes);
  await appendPdfDocument(output, source, includeGeneratedOcr ? record.ocrPages : [], pageStart, pageEnd, pageSizeHandling);
}

type RenderedSheet = { id: string; name: string; path: string; range: string; startPage: number; endPage: number };

async function appendXlsxEvidence(
  output: PDFDocument,
  record: EvidenceRecord,
  exporter: BuildOptions["workbookExporter"],
): Promise<RenderedSheet[]> {
  const workbook = record.workbook ?? await analyseXlsx(record.file);
  const selected = workbook.sheets.filter((sheet) => record.sheetSelections?.find((pick) => pick.name === sheet.name)?.included);
  if (!selected.length) throw new Error(`${record.name} has no selected visible worksheet.`);
  if (!exporter) {
    throw new Error(`${record.name} must be printed through Microsoft Excel to preserve its colours, formatting and print layout. No simplified copy was substituted. Open this project in the Exhibit Builder desktop app and retry.`);
  }
  // A source-defined Print_Area is authoritative. When none exists, allow
  // Microsoft Excel to determine its own printable UsedRange instead of
  // imposing the analyser's preview boundary on the source workbook.
  const exported = await exporter(record.file, selected.map((sheet) => ({ name: sheet.name, range: sheet.printArea ?? "", orientation: sheet.renderPlan.orientation })));
  if (exported.length !== selected.length) throw new Error(`${record.name} did not return every selected worksheet from Microsoft Excel.`);
  const byName = new Map(exported.map((sheet) => [sheet.name, sheet]));
  const rendered: RenderedSheet[] = [];
  for (const sheet of selected) {
    const printed = byName.get(sheet.name);
    if (!printed?.bytes?.length) throw new Error(`${record.name} / ${sheet.name} was not printed by Microsoft Excel.`);
    const startPage = output.getPageCount() + 1;
    const source = await PDFDocument.load(printed.bytes);
    await appendPdfDocument(output, source, [], undefined, undefined, "convert-to-a4");
    rendered.push({
      id: sheet.id,
      name: sheet.name,
      path: sheet.path,
      range: printed.range || sheet.renderPlan.range,
      startPage,
      endPage: output.getPageCount(),
    });
  }
  return rendered;
}

function drawInvisibleOcrText(page: PDFPage, text: string, font: PDFFont) {
  const { width, height } = page.getSize();
  const lines = wrapText(text, font, 5, Math.max(40, width - 80));
  let y = height - 42;
  for (const line of lines) {
    page.drawText(line, { x: 40, y, size: 5, font, color: rgb(1, 1, 1), opacity: 0 });
    y -= 6;
  }
}

async function appendPdfDocument(
  output: PDFDocument,
  source: PDFDocument,
  ocrPages: Array<{ text: string; confidence: number }> = [],
  pageStart?: number,
  pageEnd?: number,
  pageSizeHandling: NonA4PageHandling = "convert-to-a4",
) {
  assertPdfActionsSafe(source, "The selected PDF");
  const first = Math.max(1, pageStart ?? 1);
  const last = Math.min(source.getPageCount(), pageEnd ?? source.getPageCount());
  const sourcePages = source.getPages().slice(first - 1, last);
  const embeddedPages = await output.embedPages(sourcePages);
  const copiedPages = await output.copyPages(source, Array.from({ length: sourcePages.length }, (_, index) => first - 1 + index));
  const ocrFont = ocrPages.length ? await output.embedFont(StandardFonts.Helvetica) : null;
  for (const [index, embeddedPage] of embeddedPages.entries()) {
    const sourcePage = sourcePages[index];
    const rotation = ((sourcePage.getRotation().angle % 360) + 360) % 360;
    const quarterTurn = rotation === 90 || rotation === 270;
    const displayedWidth = quarterTurn ? embeddedPage.height : embeddedPage.width;
    const displayedHeight = quarterTurn ? embeddedPage.width : embeddedPage.height;
    const landscape = displayedWidth > displayedHeight;
    const a4Width = landscape ? A4_HEIGHT : A4_WIDTH;
    const a4Height = landscape ? A4_WIDTH : A4_HEIGHT;
    const alreadyA4 = Math.abs(displayedWidth - a4Width) < 2 && Math.abs(displayedHeight - a4Height) < 2;
    const keepOriginal = pageSizeHandling === "keep-original" && !alreadyA4;
    const hasAnnotations = Boolean(sourcePage.node.Annots()?.size());
    const canCopyDirectly = rotation === 0 && (alreadyA4 || keepOriginal);
    if (hasAnnotations && !canCopyDirectly) {
      const reason = rotation
        ? "has page rotation and PDF annotations"
        : "has PDF annotations that cannot be faithfully repositioned during A4 conversion";
      throw new Error(`Page ${first + index} ${reason}. Preserve fidelity by flattening the page in a trusted PDF application before building${rotation ? "" : ", or choose Keep the original page size"}.`);
    }
    let page: PDFPage;
    if (canCopyDirectly) {
      page = copiedPages[index];
      output.addPage(page);
    } else {
      const targetWidth = keepOriginal ? displayedWidth : a4Width;
      const targetHeight = keepOriginal ? displayedHeight : a4Height;
      const scale = keepOriginal ? 1 : Math.min(targetWidth / displayedWidth, targetHeight / displayedHeight);
      const drawnWidth = displayedWidth * scale;
      const drawnHeight = displayedHeight * scale;
      const offsetX = (targetWidth - drawnWidth) / 2;
      const offsetY = (targetHeight - drawnHeight) / 2;
      page = output.addPage([targetWidth, targetHeight]);
      const x = rotation === 90 || rotation === 180 ? offsetX + drawnWidth : offsetX;
      const y = rotation === 180 || rotation === 270 ? offsetY + drawnHeight : offsetY;
      page.drawPage(embeddedPage, {
        x,
        y,
        width: embeddedPage.width * scale,
        height: embeddedPage.height * scale,
        rotate: degrees(rotation),
      });
    }
    if (ocrFont && ocrPages[first - 1 + index]?.text) drawInvisibleOcrText(page, ocrPages[first - 1 + index].text, ocrFont);
  }
}

async function templatePdfFile(templateFile: TemplateFile) {
  if (templateFile.pdfFile) return templateFile.pdfFile;
  if (/\.pdf$/i.test(templateFile.file.name)) return templateFile.file;
  return convertWordTemplate(templateFile.file);
}

async function readIndexTemplateTextItems(file: File) {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  if (!pdfjs.GlobalWorkerOptions.workerSrc) {
    pdfjs.GlobalWorkerOptions.workerSrc = new URL(
      "../../node_modules/pdfjs-dist/legacy/build/pdf.worker.min.mjs",
      import.meta.url,
    ).toString();
  }
  const task = pdfjs.getDocument({
    data: new Uint8Array(await file.arrayBuffer()),
    isEvalSupported: false,
    useWorkerFetch: false,
    verbosity: 0,
  });
  const document = await task.promise;
  try {
    const page = await document.getPage(1);
    const content = await page.getTextContent();
    return content.items.flatMap((item) => {
      if (!("str" in item) || !item.str.trim()) return [];
      return [{
        str: item.str,
        x: item.transform[4],
        y: item.transform[5],
        width: "width" in item && typeof item.width === "number" ? item.width : 0,
      }];
    });
  } finally {
    await document.destroy();
  }
}

async function resolveCustomIndexGeometry(indexTemplate: TemplateFile): Promise<"custom-template" | IndexLayoutGeometryProfile> {
  const file = await templatePdfFile(indexTemplate);
  const items = await readIndexTemplateTextItems(file);
  const dateColumn = detectIndexTemplateDateColumn(items, A4_HEIGHT);
  if (!dateColumn) return "custom-template";
  return applyDetectedDateColumn(CUSTOM_TEMPLATE_INDEX_GEOMETRY, dateColumn) ?? "custom-template";
}

type TemplateReviewAudit = {
  slot: TemplateFile["slot"];
  sourceSha256: string;
  renderedSha256: string;
  sourceFormat: "pdf" | "docx" | "doc";
  appearanceConfirmed: boolean;
  matterConfirmed: true;
  placeholderConfirmed: boolean;
  textReliability: "reliable" | "limited" | "none";
};

function confirmationMatches(value: unknown, pdfSha256: string) {
  if (!value || typeof value !== "object") return false;
  const confirmation = value as { pdfSha256?: unknown; confirmedAt?: unknown };
  return confirmation.pdfSha256 === pdfSha256 && typeof confirmation.confirmedAt === "string" && confirmation.confirmedAt.trim().length > 0;
}

/** Build-time trust boundary for custom templates. Legacy templateConfirmed is
 * intentionally ignored: it never represented matter-details confirmation. */
async function validateTemplateReview(template: TemplateFile): Promise<TemplateReviewAudit> {
  const sourceFormat = template.sourceFormat ?? (/\.pdf$/i.test(template.file.name) ? "pdf" : /\.docx$/i.test(template.file.name) ? "docx" : "doc");
  const actualSourceSha256 = await sha256(await template.file.arrayBuffer());
  if (actualSourceSha256 !== template.sha256) throw new Error(`${template.file.name} changed after it was selected. Re-select and review the exact template file.`);
  if (sourceFormat !== "pdf" && !template.pdfFile) {
    throw new Error(`Preview ${template.file.name} and retain its exact converted PDF before building.`);
  }
  const renderedFile = await templatePdfFile(template);
  const renderedSha256 = await sha256(await renderedFile.arrayBuffer());
  const declaredRenderedSha256 = template.pdfSha256 ?? (sourceFormat === "pdf" ? template.sha256 : undefined);
  if (!declaredRenderedSha256 || declaredRenderedSha256 !== renderedSha256) {
    throw new Error(`${template.file.name} no longer matches the reviewed PDF artifact. Preview and confirm the exact rendered template again.`);
  }
  const review = template.reviewState;
  if (!review?.matterReview || review.matterReview.pdfSha256 !== renderedSha256) {
    throw new Error(`Review the matter details shown in the exact PDF preview for ${template.file.name} before building.`);
  }
  if (!confirmationMatches(review.matterConfirmation, renderedSha256)) {
    throw new Error(`Confirm the matter number, party names and other identifying details in ${template.file.name} before building.`);
  }
  const appearanceConfirmed = sourceFormat === "pdf" || confirmationMatches(review.appearanceConfirmation, renderedSha256);
  if (!appearanceConfirmed) {
    throw new Error(`Preview and confirm the exact converted PDF appearance for ${template.file.name} before building.`);
  }
  const hasPlaceholders = review.matterReview.placeholders.length > 0;
  const placeholderConfirmed = !hasPlaceholders || confirmationMatches(review.placeholderConfirmation, renderedSha256);
  if (!placeholderConfirmed) {
    throw new Error(`Confirm the possible unfinished placeholders in ${template.file.name} before building.`);
  }
  return {
    slot: template.slot,
    sourceSha256: actualSourceSha256,
    renderedSha256,
    sourceFormat,
    appearanceConfirmed,
    matterConfirmed: true,
    placeholderConfirmed,
    textReliability: review.matterReview.textReliability,
  };
}

async function appendTemplatePages(output: PDFDocument, templateFile: TemplateFile): Promise<TemplatePagePlacement[]> {
  const file = await templatePdfFile(templateFile);
  const template = await PDFDocument.load(await file.arrayBuffer());
  const pages = await output.embedPdf(template, template.getPageIndices());
  const placements: TemplatePagePlacement[] = [];
  for (const [sourceIndex, embeddedPage] of pages.entries()) {
    const scale = Math.min(A4_WIDTH / embeddedPage.width, A4_HEIGHT / embeddedPage.height);
    const width = embeddedPage.width * scale;
    const height = embeddedPage.height * scale;
    const offsetX = (A4_WIDTH - width) / 2;
    const offsetY = (A4_HEIGHT - height) / 2;
    const page = output.addPage([A4_WIDTH, A4_HEIGHT]);
    page.drawPage(embeddedPage, { x: offsetX, y: offsetY, width, height });
    placements.push({
      pageIndex: output.getPageCount() - 1,
      sourcePageNumber: sourceIndex + 1,
      scale,
      offsetX,
      offsetY,
    });
  }
  return placements;
}

async function templatePageCount(templateFile: TemplateFile | undefined) {
  if (!templateFile) return 1;
  const file = await templatePdfFile(templateFile);
  const document = await PDFDocument.load(await file.arrayBuffer());
  return Math.max(1, document.getPageCount());
}

async function templateFirstPageSize(templateFile: TemplateFile | undefined) {
  if (!templateFile) return null;
  const file = await templatePdfFile(templateFile);
  const document = await PDFDocument.load(await file.arrayBuffer());
  const page = document.getPage(0);
  const rotation = ((page.getRotation().angle % 360) + 360) % 360;
  const size = page.getSize();
  return rotation === 90 || rotation === 270 ? { width: size.height, height: size.width } : size;
}

function appendBuiltInDivider(output: PDFDocument, regular: PDFFont, bold: PDFFont, title: string) {
  const page = output.addPage([A4_WIDTH, A4_HEIGHT]);
  page.drawText("EXHIBIT BUNDLE DIVIDER", { x: 48, y: 760, size: 10, font: bold, color: rgb(0.56, 0.18, 0.12) });
  drawTextInsideBox(page, title.replace(/_/g, "_\u200b"), { x: 48, y: 700, maxSize: 22, minSize: 10, font: bold, maxWidth: 495, maxHeight: 520, color: rgb(0.1, 0.14, 0.2) });
}

function appendBuiltInExhibitCover(output: PDFDocument, regular: PDFFont, bold: PDFFont, mark: string, description: string) {
  const page = output.addPage([A4_WIDTH, A4_HEIGHT]);
  page.drawText("EXHIBIT", { x: 48, y: 760, size: 10, font: bold, color: rgb(0.56, 0.18, 0.12) });
  page.drawText(mark, { x: 48, y: 700, size: 24, font: bold, color: rgb(0.1, 0.14, 0.2) });
  drawWrapped(page, description, { x: 48, y: 650, size: 16, font: regular, maxWidth: 495, lineHeight: 22, color: rgb(0.12, 0.16, 0.22) });
}

function roman(value: number) {
  const values: Array<[number, string]> = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
  let remainder = value;
  let result = "";
  for (const [amount, label] of values) {
    while (remainder >= amount) {
      result += label;
      remainder -= amount;
    }
  }
  return result;
}

function templateFor(templates: TemplateFile[] | undefined, slot: TemplateFile["slot"]) {
  return templates?.find((template) => template.slot === slot);
}

function pageNumberPosition(page: PDFPage, pageNumber: number, position: PageNumberSettings["position"], label: string, font: PDFFont, size: number) {
  const { width, height } = page.getSize();
  const textWidth = font.widthOfTextAtSize(label, size);
  const top = position.startsWith("top");
  const outside = position.startsWith("outside");
  const inside = position.startsWith("inside");
  const right = position.endsWith("right");
  const onRight = outside ? pageNumber % 2 === 1 : inside ? pageNumber % 2 === 0 : right;
  const x = position.includes("centre") ? (width - textWidth) / 2 : onRight ? width - 42 - textWidth : 42;
  return { x, y: top ? height - 26 : 16 };
}

function visiblePageLabel(
  pageNumber: number,
  firstContentPage: number,
  settings: PageNumberSettings,
  records: BundleRecord[],
  numberOffset = 0,
  preliminaryOffset = 0,
  countOptionalPagesInReferences = true,
) {
  const visiblePrefix = settings.prefix;
  if (settings.matchPdfPageOrder) {
    const visibleNumber = pageNumber + numberOffset;
    const numeric = settings.padding ? String(visibleNumber).padStart(settings.padding, "0") : String(visibleNumber);
    return `${visiblePrefix}${numeric}${settings.suffix}`;
  }
  if (pageNumber < firstContentPage) {
    if (!settings.countTemplates) return "";
    if (settings.preliminary === "none") return "";
    const preliminaryNumber = pageNumber + preliminaryOffset;
    return settings.preliminary === "roman" ? roman(preliminaryNumber) : String(preliminaryNumber);
  }
  const containingContentRecord = records.find((item) => pageNumber >= item.startPage && pageNumber <= item.endPage);
  const owningRecord = records.find((item, index) => {
    const physicalStart = index === 0 ? firstContentPage : records[index - 1].endPage + 1;
    return pageNumber >= physicalStart && pageNumber <= item.endPage;
  });
  if (!countOptionalPagesInReferences && !containingContentRecord) return "";
  const containingRecord = containingContentRecord ?? owningRecord;
  const contentOrdinal = containingContentRecord
    ? records
      .filter((item) => item.startPage < containingContentRecord.startPage)
      .reduce((total, item) => total + item.endPage - item.startPage + 1, 0) + pageNumber - containingContentRecord.startPage
    : pageNumber - firstContentPage;
  const number = settings.startAt + numberOffset + (countOptionalPagesInReferences ? pageNumber - firstContentPage : contentOrdinal);
  const numeric = settings.padding ? String(number).padStart(settings.padding, "0") : String(number);
  if (settings.scheme === "section") {
    return `${containingRecord?.mark.replace(/\s+/g, "") ?? "B"}-${numeric}`;
  }
  return `${visiblePrefix}${numeric}${settings.suffix}`;
}

function pageLabelSchemeName(settings: PageNumberSettings) {
  if (settings.matchPdfPageOrder || settings.scheme === "bundle") return "Continuous";
  if (settings.scheme === "bates") return "Custom prefix and number";
  return "Exhibit mark and page";
}

function addPageLabels(
  document: PDFDocument,
  firstContentPage: number,
  settings: PageNumberSettings,
  records: BundleRecord[],
  numberOffset = 0,
  preliminaryOffset = 0,
  countOptionalPagesInReferences = true,
) {
  const nums: unknown[] = [];
  const labels = document.getPages().map((_page, index) => {
    const label = visiblePageLabel(index + 1, firstContentPage, settings, records, numberOffset, preliminaryOffset, countOptionalPagesInReferences);
    // Emit an explicit entry even for intentionally unnumbered optional pages.
    // Otherwise the preceding number-tree prefix continues onto those pages.
    nums.push(index, document.context.obj({ P: PDFHexString.fromText(label) }));
    return label;
  });
  document.catalog.set(PDFName.of("PageLabels"), document.context.obj({ Nums: nums } as any));
  return labels;
}

function validatePageLabels(document: PDFDocument, expected: readonly string[]) {
  const treeReference: any = document.catalog.get(PDFName.of("PageLabels"));
  const tree: any = treeReference && document.context.lookup(treeReference);
  const nums: any = tree?.lookup?.(PDFName.of("Nums"));
  if (!nums || nums.size() !== expected.length * 2) throw new Error(`Output validation failed: expected ${expected.length} explicit PDF page labels.`);
  for (let index = 0; index < expected.length; index += 1) {
    const pageIndex: any = document.context.lookup(nums.get(index * 2));
    const labelDictionary: any = document.context.lookup(nums.get(index * 2 + 1));
    const labelValue: any = labelDictionary?.lookup?.(PDFName.of("P"));
    if (pageIndex?.asNumber?.() !== index || labelValue?.decodeText?.() !== expected[index]) {
      throw new Error(`Output validation failed: PDF page label ${index + 1} does not match the visible numbering plan.`);
    }
  }
  return expected.length;
}

type LocalOutlineNode =
  | { kind: "section"; id: string; title: string; items: Array<{ id: string; record: BundleRecord }> }
  | { kind: "exhibit"; id: string; record: BundleRecord };

function localOutlineNodes(
  indexNodes: readonly BuildPlanIndexNode[],
  recordsById: ReadonlyMap<string, BundleRecord>,
): LocalOutlineNode[] {
  return indexNodes.flatMap<LocalOutlineNode>((node) => {
    if (node.kind === "exhibit") {
      const record = recordsById.get(node.itemId);
      return record ? [{ kind: "exhibit", id: node.itemId, record }] : [];
    }
    const items = node.itemIds.flatMap((id) => {
      const record = recordsById.get(id);
      return record ? [{ id, record }] : [];
    });
    // A section is a useful bookmark only in a PDF that physically contains
    // at least one of its exhibits. Empty remote parents would be misleading.
    return items.length ? [{ kind: "section", id: node.id, title: node.title, items }] : [];
  });
}

function exhibitOutlineTitle(record: { exhibitNumber?: number; description: string }) {
  return `${record.exhibitNumber}. ${record.description}`;
}

function addOutlineTree(
  document: PDFDocument,
  nodes: readonly LocalOutlineNode[],
  pages: PDFPage[],
) {
  if (!nodes.length) return;
  const context = document.context;
  const descendantCount = nodes.reduce((total, node) => total + 1 + (node.kind === "section" ? node.items.length : 0), 0);
  const outlines = context.obj({ Type: "Outlines", Count: descendantCount });
  const outlinesRef = context.register(outlines);
  const topItems = nodes.map((node) => {
    const firstRecord = node.kind === "section" ? node.items[0].record : node.record;
    return context.obj({
      Title: PDFHexString.fromText(node.kind === "section" ? node.title : exhibitOutlineTitle(node.record)),
      Parent: outlinesRef,
      Dest: context.obj([pages[firstRecord.startPage - 1].ref, PDFName.of("Fit")]),
    });
  });
  const topRefs = topItems.map((item) => context.register(item));
  topItems.forEach((item, index) => {
    if (index > 0) item.set(PDFName.of("Prev"), topRefs[index - 1]);
    if (index < topRefs.length - 1) item.set(PDFName.of("Next"), topRefs[index + 1]);
    const node = nodes[index];
    if (node.kind !== "section") return;
    const children = node.items.map(({ record }) => context.obj({
      Title: PDFHexString.fromText(exhibitOutlineTitle(record)),
      Parent: topRefs[index],
      Dest: context.obj([pages[record.startPage - 1].ref, PDFName.of("Fit")]),
    }));
    const childRefs = children.map((child) => context.register(child));
    children.forEach((child, childIndex) => {
      if (childIndex > 0) child.set(PDFName.of("Prev"), childRefs[childIndex - 1]);
      if (childIndex < childRefs.length - 1) child.set(PDFName.of("Next"), childRefs[childIndex + 1]);
    });
    item.set(PDFName.of("First"), childRefs[0]);
    item.set(PDFName.of("Last"), childRefs[childRefs.length - 1]);
    item.set(PDFName.of("Count"), context.obj(childRefs.length));
  });
  outlines.set(PDFName.of("First"), topRefs[0]);
  outlines.set(PDFName.of("Last"), topRefs[topRefs.length - 1]);
  document.catalog.set(PDFName.of("Outlines"), outlinesRef);
  document.catalog.set(PDFName.of("PageMode"), PDFName.of("UseOutlines"));
}

function addIndexLink(
  document: PDFDocument,
  indexPage: PDFPage,
  destination: PDFPage,
  rectangle: { x: number; top: number; width: number; height: number },
) {
  const pageHeight = indexPage.getHeight();
  const annotation = document.context.obj({
    Type: "Annot",
    Subtype: "Link",
    Rect: [rectangle.x, pageHeight - rectangle.top - rectangle.height, rectangle.x + rectangle.width, pageHeight - rectangle.top],
    Border: [0, 0, 0],
    A: {
      Type: "Action",
      S: "GoTo",
      D: [destination.ref, PDFName.of("Fit")],
    },
  });
  const annotationRef = document.context.register(annotation);
  const existing = indexPage.node.Annots();
  if (existing) {
    existing.push(annotationRef);
  } else {
    indexPage.node.set(
      PDFName.of("Annots"),
      document.context.obj([annotationRef]),
    );
  }
}

/** Resolve PDF object references after a save/reopen. Link counts alone are
 * insufficient: non-local rows must have no action, while every local action
 * and every hierarchical outline destination must resolve inside this PDF. */
function validateFinalDestinations(
  document: PDFDocument,
  expectedLinks: ReadonlyArray<{ id: string; targetPage: number }>,
  indexStart: number,
  indexCount: number,
  expectedOutlines: readonly LocalOutlineNode[],
) {
  const pages = document.getPages();
  const pageForRef = (ref: unknown) => pages.findIndex((page) => String(page.ref) === String(ref)) + 1;
  const resolveDestination = (value: any): number => {
    const array = value?.asArray?.() ?? value;
    const target = array?.get?.(0) ?? array?.[0];
    return pageForRef(target);
  };
  let links = 0;
  for (let pageNumber = indexStart; pageNumber < indexStart + indexCount; pageNumber++) {
    const annots: any = pages[pageNumber - 1]?.node.Annots();
    for (let index = 0; annots && index < annots.size(); index++) {
      const annotation: any = document.context.lookup(annots.get(index));
      const action: any = annotation.lookup?.(PDFName.of("A"));
      if (String(action?.get?.(PDFName.of("S"))) !== "/GoTo") continue;
      const target = resolveDestination(action.get(PDFName.of("D")));
      const expected = expectedLinks[links]?.targetPage;
      if (!target || target !== expected) throw new Error(`Output validation failed: index GoTo ${links + 1} resolves to page ${target || "none"}, expected ${expected}.`);
      links++;
    }
  }
  if (links !== expectedLinks.length) throw new Error(`Output validation failed: expected ${expectedLinks.length} local index GoTo links, found ${links}.`);
  const outlines: any = document.catalog.get(PDFName.of("Outlines"));
  const root: any = outlines && document.context.lookup(outlines);
  let itemRef: any = root?.get?.(PDFName.of("First"));
  let outlineIndex = 0;
  let bookmarkCount = 0;
  while (itemRef && outlineIndex < expectedOutlines.length) {
    const item: any = document.context.lookup(itemRef);
    const expectedNode = expectedOutlines[outlineIndex];
    const firstRecord = expectedNode.kind === "section" ? expectedNode.items[0].record : expectedNode.record;
    const expectedTitle = expectedNode.kind === "section" ? expectedNode.title : exhibitOutlineTitle(expectedNode.record);
    const actualTitle: any = item.get(PDFName.of("Title"));
    if (actualTitle?.decodeText?.() !== expectedTitle) throw new Error(`Output validation failed: top-level outline ${outlineIndex + 1} title mismatch.`);
    if (String(item.get(PDFName.of("Parent"))) !== String(outlines)) throw new Error(`Output validation failed: top-level outline ${outlineIndex + 1} has the wrong parent.`);
    const target = resolveDestination(item.get(PDFName.of("Dest")));
    if (target !== firstRecord.startPage) throw new Error(`Output validation failed: top-level outline ${outlineIndex + 1} destination mismatch.`);
    bookmarkCount += 1;
    let childRef: any = item.get(PDFName.of("First"));
    let childIndex = 0;
    const expectedChildren = expectedNode.kind === "section" ? expectedNode.items : [];
    const declaredChildren: any = item.get(PDFName.of("Count"));
    if ((declaredChildren?.asNumber?.() ?? 0) !== expectedChildren.length) throw new Error(`Output validation failed: outline ${outlineIndex + 1} has the wrong child count.`);
    while (childRef && childIndex < expectedChildren.length) {
      const child: any = document.context.lookup(childRef);
      const childTitle: any = child.get(PDFName.of("Title"));
      const expectedChildTitle = exhibitOutlineTitle(expectedChildren[childIndex].record);
      if (childTitle?.decodeText?.() !== expectedChildTitle) throw new Error(`Output validation failed: child outline ${outlineIndex + 1}.${childIndex + 1} title mismatch.`);
      const childTarget = resolveDestination(child.get(PDFName.of("Dest")));
      if (childTarget !== expectedChildren[childIndex].record.startPage) throw new Error(`Output validation failed: child outline ${outlineIndex + 1}.${childIndex + 1} destination mismatch.`);
      if (String(child.get(PDFName.of("Parent"))) !== String(itemRef)) throw new Error(`Output validation failed: child outline ${outlineIndex + 1}.${childIndex + 1} has the wrong parent.`);
      bookmarkCount += 1;
      childIndex += 1;
      childRef = child.get(PDFName.of("Next"));
    }
    if (childIndex !== expectedChildren.length || childRef) throw new Error(`Output validation failed: outline ${outlineIndex + 1} has an unexpected child structure.`);
    outlineIndex += 1;
    itemRef = item.get(PDFName.of("Next"));
  }
  if (outlineIndex !== expectedOutlines.length || itemRef) throw new Error(`Output validation failed: expected ${expectedOutlines.length} top-level outlines, found ${outlineIndex}${itemRef ? " plus additional items" : ""}.`);
  return { links, bookmarks: bookmarkCount };
}

async function buildBundleLegacy(
  analysis: AnalysisResult,
  candidates: ExhibitCandidate[],
): Promise<BuildResult> {
  const included = candidates
    .filter((candidate) => candidate.included)
    .sort(
      (left, right) => (left.sequenceOrder ?? left.provisionalNumber) - (right.sequenceOrder ?? right.provisionalNumber),
    );
  if (!included.length) throw new Error("Select at least one exhibit.");
  if (included.some((candidate) => !candidate.evidenceId)) {
    throw new Error("Every included exhibit needs a confirmed source file.");
  }
  if (included.some((candidate) => !candidate.confirmed)) {
    throw new Error(
      "Confirm every included provisional number and source match before building.",
    );
  }
  const normalizedMarks = included.map((candidate) =>
    normalizeExhibitMark(candidate.mark),
  );
  if (new Set(normalizedMarks).size !== normalizedMarks.length) {
    throw new Error("Every included exhibit needs a unique provisional number.");
  }

  const document = await PDFDocument.create();
  document.setTitle(`${analysis.caseTitle} - Exhibit Bundle`);
  document.setSubject("Deterministically generated exhibit bundle");
  document.setProducer("Exhibit Builder");
  document.setCreator("Exhibit Builder");

  const regular = await document.embedFont(StandardFonts.Helvetica);
  const bold = await document.embedFont(StandardFonts.HelveticaBold);
  const indexPage = document.addPage([A4_WIDTH, A4_HEIGHT]);
  const records: BundleRecord[] = [];

  for (const candidate of included) {
    const evidence = analysis.evidence.find(
      (record) => record.id === candidate.evidenceId,
    );
    if (!evidence) continue;
    const startPage = document.getPageCount() + 1;
    if (evidence.extension === "pdf") {
      await appendPdfEvidence(document, evidence);
    } else if (evidence.extension === "eml") {
      const printChildren = emailChildrenForDisposition(evidence.emailAttachments, candidate.emailAttachmentDispositions, "print-with-email");
      await appendEmailEvidence(
        document,
        evidence,
        regular,
        bold,
        printChildren,
      );
    } else {
      await appendTextEvidence(
        document,
        evidence,
        regular,
        bold,
      );
    }
    records.push({
      mark: candidate.mark,
      exhibitNumber: records.length + 1,
      description: candidate.description,
      fileName: evidence.name,
      startPage,
      endPage: document.getPageCount(),
      statementParagraph: candidate.paragraph,
      statementReferences: [{ paragraph: candidate.paragraph, citation: candidate.citation, statementName: candidate.statementName, statementId: candidate.statementId, witnessInitials: candidate.witnessInitials }],
      sourceHash: evidence.sha256,
    });
  }

  const pages = document.getPages();

  indexPage.drawText("EXHIBIT BUNDLE", {
    x: 44,
    y: 778,
    size: 10,
    font: bold,
    color: rgb(0.56, 0.18, 0.12),
  });
  drawWrapped(indexPage, analysis.caseTitle, {
    x: 44,
    y: 742,
    size: 18,
    font: bold,
    maxWidth: 507,
    lineHeight: 23,
    color: rgb(0.1, 0.14, 0.2),
  });
  indexPage.drawText("Index of exhibits", {
    x: 44,
    y: 684,
    size: 11,
    font: regular,
    color: rgb(0.35, 0.38, 0.42),
  });
  indexPage.drawLine({
    start: { x: 44, y: 661 },
    end: { x: 551, y: 661 },
    thickness: 1,
    color: rgb(0.18, 0.23, 0.31),
  });

  let rowY = 625;
  for (const record of records) {
    indexPage.drawText(record.mark, {
      x: 48,
      y: rowY,
      size: 10,
      font: bold,
      color: rgb(0.56, 0.18, 0.12),
    });
    drawWrapped(indexPage, record.description, {
      x: 103,
      y: rowY,
      size: 10,
      font: regular,
      maxWidth: 350,
      lineHeight: 12,
    });
    const range =
      record.startPage === record.endPage
        ? String(record.startPage)
        : `${record.startPage}-${record.endPage}`;
    indexPage.drawText(range, {
      x: 505,
      y: rowY,
      size: 10,
      font: bold,
      color: rgb(0.12, 0.16, 0.22),
    });
    addIndexLink(
      document,
      indexPage,
      pages[record.startPage - 1],
      { x: 45, top: A4_HEIGHT - rowY - 29, width: 505, height: 35 },
    );
    rowY -= 46;
  }

  addOutlineTree(document, records.map((record, index) => ({ kind: "exhibit", id: `legacy-${index}`, record })), pages);
  const bytes = await document.save({ useObjectStreams: false });
  const reopened = await PDFDocument.load(bytes);
  const allPagesA4 = reopened.getPages().every((page) => {
    const { width, height } = page.getSize();
    return (
      Math.abs(width - A4_WIDTH) < 0.02 &&
      Math.abs(height - A4_HEIGHT) < 0.02
    );
  });
  const bundleHash = await sha256(bytes);
  const selectedEvidenceIds = new Set(
    included.map((candidate) => candidate.evidenceId),
  );
  const excludedEvidence = analysis.evidence.filter(
    (record) => !selectedEvidenceIds.has(record.id),
  );
  const manifest = {
    schemaVersion: "1.1",
    caseTitle: analysis.caseTitle,
    createdAt: new Date().toISOString(),
    statement: {
      fileName: analysis.statementName,
      sha256: analysis.statementHash,
      modified: false,
    },
    exhibits: records,
    excludedFiles: excludedEvidence.map((record) => ({
      fileName: record.name,
      sha256: record.sha256,
      reason:
        record.marker === "N/A"
          ? "File explicitly marked N/A and not cited"
          : "No confirmed citation match",
    })),
    output: {
      fileName: "Exhibit_Bundle.pdf",
      sha256: bundleHash,
      pageCount: reopened.getPageCount(),
      pageSize: "A4 portrait",
    },
  };

  return {
    bytes,
    fileName: "Exhibit_Bundle.pdf",
    sha256: bundleHash,
    pageCount: reopened.getPageCount(),
    records,
    manifest,
    checks: [
      {
        label: "All included exhibits matched",
        status: "pass",
        detail: `${records.length} of ${records.length} confirmed`,
      },
      {
        label: "Witness statement unchanged",
        status: "pass",
        detail: "The source statement was read only and was not rewritten",
      },
      {
        label: "Continuous pagination",
        status: "pass",
        detail: `${reopened.getPageCount()} numbered PDF pages`,
      },
      {
        label: "Uniform A4 page size",
        status: allPagesA4 ? "pass" : "warning",
        detail: allPagesA4
          ? "Every bundle page is A4 portrait"
          : "One or more pages are not A4 portrait",
      },
      {
        label: "Index destinations",
        status: "pass",
        detail: `${records.length} internal links created`,
      },
      {
        label: "PDF bookmarks",
        status: "pass",
        detail: `${records.length} exhibit bookmarks created`,
      },
      {
        label: "Unreferenced evidence",
        status: excludedEvidence.length ? "warning" : "pass",
        detail: `${excludedEvidence.length} supplied file excluded`,
      },
    ],
  };
}

/**
 * Two-pass, local-only compositor. The first pass establishes each exhibit's
 * rendered page count. The second creates cover/index pages, then links and
 * outlines using the final destinations. Source PDF pages that do not require
 * geometric conversion are copied so their annotations remain intact.
 */
export async function buildBundle(
  analysis: AnalysisResult,
  candidates: ExhibitCandidate[],
  options: BuildOptions = {},
): Promise<BuildResult> {
  const report = options.onProgress ?? (() => {});
  report("Checking the confirmed exhibits");
  const profile = BUNDLE_PROFILES.find((item) => item.id === options.profileId) ?? BUNDLE_PROFILES[0];
  const pagination = { ...DEFAULT_PAGINATION, ...options.pagination };
  const layout = { ...DEFAULT_BUNDLE_LAYOUT, ...options.layout };
  const pageSizeChoices = options.pageSizeChoices ?? {};
  const resolutions = options.resolutions ?? [];
  const effectiveCandidates = candidates.map((candidate) => {
    if (candidateIsExcluded(resolutions, candidate.id)) return { ...candidate, included: false, confirmed: false };
    if (candidate.evidenceId) {
      const source = analysis.evidence.find((record) => record.id === candidate.evidenceId);
      if (source && sourceIsExcluded(resolutions, source.id, source.sha256)) return { ...candidate, included: false, confirmed: false };
    }
    return candidate;
  });
  const requested = effectiveCandidates.filter((candidate) => candidate.included);
  if (requested.some((candidate) => !candidate.evidenceId)) {
    throw new Error("Every included exhibit needs a confirmed source file.");
  }
  if (requested.some((candidate) => !candidate.confirmed)) {
    throw new Error("Confirm every included provisional number and source match before building.");
  }
  const referenceIdentities = new Set(requested
    .filter((candidate) => !candidate.manualAddition)
    .map((candidate) => `${(candidate.exhibitInitials ?? candidate.witnessInitials ?? "EX").replace(/\s+/g, "").toUpperCase()}${candidate.exhibitSequence ?? 1}`));
  if (referenceIdentities.size > 1) {
    throw new Error(`Statement reference marks conflict (${[...referenceIdentities].join(", ")}). Choose one exhibit mark in Optional settings before building.`);
  }
  // Preflight is deliberately recomputed at build time. UI state can change
  // after a displayed preflight has been calculated, and a caller must not be
  // able to bypass local repeat-source safeguards with a stale check list.
  const preflight = applyBuildResolutions(runPreflight(analysis, effectiveCandidates, profile), resolutions);
  if (hasBlockingPreflight(preflight)) {
    const first = preflight.find((check) => check.severity === "blocking");
    throw new Error(`Preflight blocked build: ${first?.label ?? "review required"}. ${first?.detail ?? ""}`.trim());
  }
  const derivedGroups = deriveExhibitGroups(analysis, effectiveCandidates);
  const arrangement = reconcileBundleArrangement(
    options.arrangement ?? bundleArrangementFromLegacyOrder(options.canonicalOrder ?? []),
    derivedGroups.map((group) => group.id),
  );
  const groups = orderExhibitGroups(derivedGroups, flattenBundleArrangement(arrangement));
  if (!groups.length) throw new Error("Select at least one exhibit.");
  const indexNodes: BuildPlanIndexNode[] = arrangement.nodes.map((node) => node.type === "section"
    ? { kind: "section", id: node.id, title: node.heading, itemIds: node.exhibits.map((exhibit) => exhibit.exhibitId) }
    : { kind: "exhibit", itemId: node.exhibitId });

  const fallbackSlots = templateFallbackSlots(resolutions, options.templates);
  const activeTemplates = (options.templates ?? []).filter((template) => {
    const sourceFormat = template.sourceFormat ?? template.file.name.split(".").pop()?.toLowerCase();
    return sourceFormat === "pdf" || !fallbackSlots.has(template.slot);
  });
  const usedTemplateSlots = new Set<TemplateFile["slot"]>([
    "cover",
    "index",
    ...(layout.includeDividerPages ? ["divider" as const] : []),
    ...(layout.includeExhibitCoverPages ? ["exhibitCover" as const] : []),
  ]);
  const templateReviewAudits: TemplateReviewAudit[] = [];
  for (const template of activeTemplates.filter((candidate) => usedTemplateSlots.has(candidate.slot))) {
    templateReviewAudits.push(await validateTemplateReview(template));
  }
  const templateReviewAuditBySlot = new Map(templateReviewAudits.map((audit) => [audit.slot, audit]));
  const templateManifestRecords = () => activeTemplates.map((template) => {
    const audit = templateReviewAuditBySlot.get(template.slot);
    const included = usedTemplateSlots.has(template.slot);
    const confirmation = template.reviewState?.matterConfirmation;
    const confirmedValues = confirmation && template.pdfSha256 && confirmation.pdfSha256 === template.pdfSha256
      ? matterValuesFromConfirmation(confirmation)
      : undefined;
    return {
      slot: template.slot,
      sourceSha256: audit?.sourceSha256 ?? template.sha256,
      renderedSha256: audit?.renderedSha256 ?? template.pdfSha256 ?? template.sha256,
      included,
      approved: included ? Boolean(audit) : false,
      appearanceConfirmed: audit?.appearanceConfirmed ?? false,
      matterConfirmed: audit?.matterConfirmed ?? false,
      placeholderConfirmed: audit?.placeholderConfirmed ?? false,
      textReliability: audit?.textReliability ?? null,
      confirmedMatterNumbers: confirmedValues?.matterNumbers ?? null,
      confirmedPartyNames: confirmedValues?.partyNames ?? null,
      confirmedForums: confirmedValues?.forums ?? null,
      confirmedMatterTitles: confirmedValues?.matterTitles ?? null,
    };
  });

  report("Preparing the exhibit pages", `${groups.length} confirmed exhibit${groups.length === 1 ? "" : "s"}`);
  const body = await PDFDocument.create();
  const bodyRegular = await body.embedFont(StandardFonts.Helvetica);
  const bodyBold = await body.embedFont(StandardFonts.HelveticaBold);
  const bodyRecords: BundleRecord[] = [];
  // Selecting a template only makes it available. It enters the exhibit
  // bundle only after the reviewer has explicitly opted into that page type.
  const exhibitCover = layout.includeExhibitCoverPages ? templateFor(activeTemplates, "exhibitCover") : undefined;
  const divider = layout.includeDividerPages ? templateFor(activeTemplates, "divider") : undefined;
  // Citation ranges are page positions inside a witness's exhibit bundle, not
  // PDF positions in the final combined file.  Templates and the index do not
  // consume this cursor, and each witness bundle has an independent cursor.
  const bundlePageCursors = new Map<string, number>();
  const allocateBundlePages = (candidate: ExhibitCandidate, contentPageCount: number, optionalPageCount = 0) => {
    const initials = (candidate.exhibitInitials ?? candidate.witnessInitials ?? "EX").replace(/\s+/g, "").toUpperCase();
    const sequence = candidate.exhibitSequence ?? 1;
    const key = `${candidate.witnessKey ?? initials}:${initials}:${sequence}`;
    const start = bundlePageCursors.get(key) ?? 1;
    const optionalPages = layout.countOptionalPagesInReferences ? optionalPageCount : 0;
    const contentStart = start + optionalPages;
    const contentEnd = contentStart + contentPageCount - 1;
    bundlePageCursors.set(key, contentEnd + 1);
    return { initials, sequence, start, end: contentEnd, contentStart, contentEnd };
  };
  let previousStatement: string | undefined;
  for (const [groupIndex, group] of groups.entries()) {
    const { canonical: candidate, evidence } = group;
    const outputMark = group.outputMark;
    report("Rendering exhibits", `${groupIndex + 1} of ${groups.length}: ${evidence.name}`);
    if (evidence.encrypted) throw new Error(`Preflight blocked build: ${evidence.name} is encrypted.`);
    const startPage = body.getPageCount() + 1;
    if (layout.includeDividerPages && candidate.statementName !== previousStatement) {
      if (divider) await appendTemplatePages(body, divider);
      else appendBuiltInDivider(body, bodyRegular, bodyBold, candidate.statementName ?? analysis.statementName);
    }
    if (layout.includeExhibitCoverPages) {
      if (exhibitCover) await appendTemplatePages(body, exhibitCover);
      else appendBuiltInExhibitCover(body, bodyRegular, bodyBold, outputMark, candidate.description);
    }
    const contentStartPage = body.getPageCount() + 1;
    const optionalPageCount = contentStartPage - startPage;
    const segmentItems = group.members
      .map((member) => ({ member, evidence: analysis.evidence.find((record) => record.id === member.evidenceId) }))
      .filter((item): item is { member: ExhibitCandidate; evidence: EvidenceRecord } => Boolean(item.evidence))
      .filter((item, index, items) => items.findIndex((other) => other.evidence.sha256 === item.evidence.sha256) === index);
    if (segmentItems.length > 1 && segmentItems.every((item) => item.evidence.extension === "pdf")) {
      let localCursor = 1;
      const refs: BundleRecord["statementReferences"] = [];
      const segmentRanges = new Map<string, { start: number; end: number }>();
      for (const segment of segmentItems) {
        const before = body.getPageCount();
        await appendPdfEvidence(body, segment.evidence, segment.member.pageStart, segment.member.pageEnd, !shouldSkipOcr(segment.evidence.id, segment.evidence.sha256, resolutions), pageSizeChoices[segment.evidence.id] ?? "convert-to-a4");
        const pages = body.getPageCount() - before;
        const start = localCursor;
        const end = localCursor + pages - 1;
        localCursor = end + 1;
        segmentRanges.set(segment.evidence.sha256, { start, end });
        refs.push({ paragraph: segment.member.paragraph, citation: segment.member.citation, statementName: segment.member.statementName, statementId: segment.member.statementId, witnessInitials: segment.member.witnessInitials, citationToken: segment.member.citationToken, exhibitInitials: segment.member.exhibitInitials ?? candidate.exhibitInitials ?? candidate.witnessInitials, exhibitSequence: segment.member.exhibitSequence ?? candidate.exhibitSequence ?? 1, requestedExhibitPageStart: segment.member.requestedExhibitPageStart, requestedExhibitPageEnd: segment.member.requestedExhibitPageEnd, citationResolution: segment.member.citationResolution ?? "none", exhibitPageStart: start, exhibitPageEnd: end });
      }
      const bundleRange = allocateBundlePages(candidate, localCursor - 1, optionalPageCount);
      const sequence = bundleRange.sequence;
      const allRefs = group.members.map((member) => {
        const memberEvidence = analysis.evidence.find((record) => record.id === member.evidenceId);
        const range = memberEvidence ? segmentRanges.get(memberEvidence.sha256) : undefined;
        return { paragraph: member.paragraph, citation: member.citation, statementName: member.statementName, statementId: member.statementId, witnessInitials: member.witnessInitials, citationToken: member.citationToken, exhibitInitials: member.exhibitInitials ?? bundleRange.initials, exhibitSequence: member.exhibitSequence ?? sequence, requestedExhibitPageStart: member.requestedExhibitPageStart, requestedExhibitPageEnd: member.requestedExhibitPageEnd, citationResolution: member.citationResolution ?? "none", exhibitPageStart: range ? bundleRange.contentStart + range.start - 1 : undefined, exhibitPageEnd: range ? bundleRange.contentStart + range.end - 1 : undefined };
      });
      bodyRecords.push({ mark: outputMark, exhibitNumber: group.exhibitNumber, exhibitPageStart: bundleRange.contentStart, exhibitPageEnd: bundleRange.contentEnd, description: candidate.description, fileName: segmentItems.map((item) => item.evidence.name).join("; "), startPage: contentStartPage, endPage: body.getPageCount(), statementParagraph: candidate.paragraph, statementReferences: allRefs, sourceHash: segmentItems.map((item) => item.evidence.sha256).join(",") });
      previousStatement = candidate.statementName;
      continue;
    }
    if (evidence.extension === "pdf") await appendPdfEvidence(body, evidence, candidate.pageStart, candidate.pageEnd, !shouldSkipOcr(evidence.id, evidence.sha256, resolutions), pageSizeChoices[evidence.id] ?? "convert-to-a4");
    else if (evidence.extension === "xlsx") {
      const sheets = await appendXlsxEvidence(body, evidence, options.workbookExporter);
      const localEnd = body.getPageCount() - contentStartPage + 1;
      const bundleRange = allocateBundlePages(candidate, localEnd, optionalPageCount);
      const refs = group.references.map((reference) => ({ ...reference, exhibitInitials: reference.exhibitInitials ?? bundleRange.initials, exhibitSequence: reference.exhibitSequence ?? bundleRange.sequence, exhibitPageStart: bundleRange.contentStart, exhibitPageEnd: bundleRange.contentEnd, citationResolution: reference.citationResolution ?? "none" }));
      bodyRecords.push({ mark: outputMark, exhibitNumber: group.exhibitNumber, exhibitPageStart: bundleRange.contentStart, exhibitPageEnd: bundleRange.contentEnd, description: candidate.description, fileName: evidence.name, startPage: sheets[0].startPage, endPage: sheets[sheets.length - 1].endPage, statementParagraph: candidate.paragraph, statementReferences: refs, sourceHash: evidence.sha256, workbookSheet: { id: sheets[0].id, name: sheets[0].name, path: sheets[0].path, range: sheets[0].range }, workbookSheets: sheets.map((sheet) => ({ id: sheet.id, name: sheet.name, path: sheet.path, range: sheet.range })) });
      previousStatement = candidate.statementName;
      continue;
    }
    else if (evidence.extension === "eml") {
      const printChildren = emailChildrenForDisposition(evidence.emailAttachments, candidate.emailAttachmentDispositions, "print-with-email");
      await appendEmailEvidence(body, evidence, bodyRegular, bodyBold, printChildren, options);
    }
    else await appendTextEvidence(body, evidence, bodyRegular, bodyBold);
    const localEnd = body.getPageCount() - contentStartPage + 1;
    const bundleRange = allocateBundlePages(candidate, localEnd, optionalPageCount);
    const refs = group.references.map((reference) => ({ ...reference, exhibitInitials: reference.exhibitInitials ?? bundleRange.initials, exhibitSequence: reference.exhibitSequence ?? bundleRange.sequence, exhibitPageStart: bundleRange.contentStart, exhibitPageEnd: bundleRange.contentEnd, citationResolution: reference.citationResolution ?? "none" }));
    bodyRecords.push({
      mark: outputMark,
      exhibitNumber: group.exhibitNumber,
      exhibitPageStart: bundleRange.contentStart,
      exhibitPageEnd: bundleRange.contentEnd,
      description: candidate.description,
      fileName: evidence.name,
      startPage: contentStartPage,
      endPage: body.getPageCount(),
      statementParagraph: candidate.paragraph,
      statementReferences: refs,
      sourceHash: evidence.sha256,
    });
    previousStatement = candidate.statementName;
  }

  bodyRecords.forEach((record, index) => {
    const candidate = groups[index]?.canonical;
    // The reader-facing index number is the final canonical document order.
    // It is distinct from a witness exhibit-bundle mark such as LV1, which
    // can legitimately repeat across several cited source documents.
    record.exhibitNumber = index + 1;
    const manualAddition = Boolean(candidate?.manualAddition);
    record.manualAddition = manualAddition;
    record.citationStatus = manualAddition ? "not-cited-manual-addition" : "cited";
    record.documentDate = candidate?.date;
    record.manualAddedAt = candidate?.manualAddedAt;
    record.manualWarningAcknowledgedAt = candidate?.manualWarningAcknowledgedAt;
    if (manualAddition) record.statementParagraph = null;
    const evidence = groups[index]?.evidence;
    if (evidence?.emailAttachments?.length) {
      record.emailAttachments = evidence.emailAttachments.map((child) => ({
        name: child.name,
        identity: child.identity,
        sha256: child.sha256,
        parentSha256: child.parentSha256,
        disposition: candidate?.emailAttachmentDispositions?.[child.identity] ?? "unresolved",
      }));
    }
  });

  report("Creating the authoritative build plan");
  const cover = templateFor(activeTemplates, "cover");
  const indexTemplate = templateFor(activeTemplates, "index");
  const coverMatterValues = cover && coverWritesMatterText(layout) && cover.reviewState?.matterReview
    ? effectiveMatterValues(cover.reviewState.matterReview, cover.reviewState.matterConfirmation, cover.pdfSha256)
    : undefined;
  const indexMatterValues = indexTemplate?.reviewState?.matterReview
    ? effectiveMatterValues(indexTemplate.reviewState.matterReview, indexTemplate.reviewState.matterConfirmation, indexTemplate.pdfSha256)
    : undefined;
  const bundleTitle = resolvedBundleTitle(analysis.caseTitle, [coverMatterValues, indexMatterValues, layout.builtInMatter]);
  const coverPages = await templatePageCount(cover);
  const indexTemplatePages = await templatePageCount(indexTemplate);
  if (indexTemplate && indexTemplatePages !== 1) {
    throw new Error(`Custom index templates must contain exactly one PDF page; the selected template contains ${indexTemplatePages}.`);
  }
  const indexTemplateSize = await templateFirstPageSize(indexTemplate);
  if (indexTemplateSize && (Math.abs(indexTemplateSize.width - A4_WIDTH) >= 2 || Math.abs(indexTemplateSize.height - A4_HEIGHT) >= 2)) {
    throw new Error("The fixed-layout index background must be one portrait A4 page. Use the guided-sample index as the layout reference.");
  }
  const indexGeometry = indexTemplate ? await resolveCustomIndexGeometry(indexTemplate) : "built-in";
  let previousBodyEnd = 0;
  const planItems = bodyRecords.map((record, recordIndex) => {
    const group = groups[recordIndex];
    const bodyStartPage = previousBodyEnd + 1;
    const bodyEndPage = record.endPage;
    previousBodyEnd = bodyEndPage;
    const optionalPages = Math.max(0, record.startPage - bodyStartPage);
    const contentPages = Math.max(1, record.endPage - record.startPage + 1);
    const initials = record.statementReferences[0]?.exhibitInitials ?? group?.canonical.exhibitInitials ?? group?.canonical.witnessInitials ?? "EX";
    const sequence = record.statementReferences[0]?.exhibitSequence ?? group?.canonical.exhibitSequence ?? 1;
    return {
      id: group?.id ?? `record-${recordIndex}`,
      recordIndex,
      indexNumber: recordIndex + 1,
      witnessKey: group?.canonical.witnessKey ?? initials,
      initials,
      sequence,
      sourceHashes: record.sourceHash.split(",").map((hash) => hash.trim()).filter(Boolean),
      bodyStartPage,
      bodyContentStartPage: record.startPage,
      bodyEndPage,
      physicalPages: Math.max(1, bodyEndPage - bodyStartPage + 1),
      optionalPages,
      contentPages,
      references: record.statementReferences.map((reference, referenceIndex) => ({
        id: `${recordIndex}-${referenceIndex}`,
        relativeStart: Number.isSafeInteger(reference.exhibitPageStart) && Number.isSafeInteger(record.exhibitPageStart) ? reference.exhibitPageStart! - record.exhibitPageStart! : null,
        relativeEnd: Number.isSafeInteger(reference.exhibitPageEnd) && Number.isSafeInteger(record.exhibitPageStart) ? reference.exhibitPageEnd! - record.exhibitPageStart! : null,
      })),
    };
  });
  type PlannedVolumeRecords = {
    volumeNumber: number;
    numberOffset: number;
    preliminaryOffset: number;
    firstContentPage: number;
    records: BundleRecord[];
    recordsById: Map<string, BundleRecord>;
    indexRecordsById: Map<string, { record: BundleRecord; indexPageLabelStart: string; indexPageLabelEnd: string }>;
  };
  const recordsForPlan = (plan: BundleBuildPlan): PlannedVolumeRecords[] => plan.volumes.map((volume) => {
    const earlierVolumes = plan.volumes.slice(0, volume.number - 1);
    const numberOffset = pagination.volumeNumbering === "continuous"
      ? earlierVolumes.reduce((total, earlier) => total + (pagination.matchPdfPageOrder ? earlier.totalPages : earlier.referencePages), 0)
      : 0;
    const preliminaryOffset = pagination.volumeNumbering === "continuous"
      ? earlierVolumes.reduce((total, earlier) => total + earlier.coverPages + earlier.indexPages, 0)
      : 0;
    const firstContentPage = volume.coverPages + volume.indexPages + 1;
    const baseRecords = volume.items.map((item) => {
      const source = bodyRecords[item.recordIndex];
      return {
        ...source,
        startPage: item.physicalContentStartPage,
        endPage: item.physicalEndPage,
        exhibitPageStart: item.legalStartPage,
        exhibitPageEnd: item.legalEndPage,
        volumeNumber: volume.number,
        statementReferences: source.statementReferences.map((reference, referenceIndex) => ({
          ...reference,
          exhibitPageStart: item.references[referenceIndex]?.legalStartPage ?? undefined,
          exhibitPageEnd: item.references[referenceIndex]?.legalEndPage ?? undefined,
          volumeNumber: volume.number,
        })),
      } satisfies BundleRecord;
    });
    const indexRecordsById = new Map<string, { record: BundleRecord; indexPageLabelStart: string; indexPageLabelEnd: string }>();
    const records = baseRecords.map((record, recordIndex) => {
      const item = volume.items[recordIndex];
      const exhibitPageLabelStart = visiblePageLabel(record.startPage, firstContentPage, pagination, baseRecords, numberOffset, 0, layout.countOptionalPagesInReferences);
      const exhibitPageLabelEnd = visiblePageLabel(record.endPage, firstContentPage, pagination, baseRecords, numberOffset, 0, layout.countOptionalPagesInReferences);
      const resolved = {
        ...record,
        exhibitPageLabelStart,
        exhibitPageLabelEnd,
        statementReferences: record.statementReferences.map((reference, referenceIndex) => {
          const plannedReference = item.references[referenceIndex];
          const physicalStart = plannedReference.relativeStart === null ? null : item.physicalContentStartPage + plannedReference.relativeStart;
          const physicalEnd = plannedReference.relativeEnd === null ? null : item.physicalContentStartPage + plannedReference.relativeEnd;
          return {
            ...reference,
            exhibitPageLabelStart: physicalStart === null ? undefined : visiblePageLabel(physicalStart, firstContentPage, pagination, baseRecords, numberOffset, 0, layout.countOptionalPagesInReferences),
            exhibitPageLabelEnd: physicalEnd === null ? undefined : visiblePageLabel(physicalEnd, firstContentPage, pagination, baseRecords, numberOffset, 0, layout.countOptionalPagesInReferences),
          };
        }),
      } satisfies BundleRecord;
      indexRecordsById.set(item.id, {
        record: resolved,
        indexPageLabelStart: exhibitPageLabelStart,
        indexPageLabelEnd: exhibitPageLabelEnd,
      });
      return resolved;
    });
    return {
      volumeNumber: volume.number,
      numberOffset,
      preliminaryOffset,
      firstContentPage,
      records,
      recordsById: new Map(volume.items.map((item, index) => [item.id, records[index]])),
      indexRecordsById,
    };
  });

  let completeIndexPages = 1;
  let buildPlan!: BundleBuildPlan;
  let plannedVolumeRecords: PlannedVolumeRecords[] = [];
  let convergedIndexLayout!: IndexLayoutPlan;
  const convergenceStates = new Set<string>();
  for (let iteration = 0; iteration < 16; iteration += 1) {
    buildPlan = createBuildPlan(planItems, {
      pageLimit: layout.volumePageLimit,
      coverPages,
      completeIndexPages,
      indexNodes,
      includeDividerPages: layout.includeDividerPages,
      includeExhibitCoverPages: layout.includeExhibitCoverPages,
      countOptionalPagesInReferences: layout.countOptionalPagesInReferences,
      matchPdfPageOrder: pagination.matchPdfPageOrder,
      volumeNumbering: pagination.volumeNumbering,
      startAt: pagination.startAt,
    });
    plannedVolumeRecords = recordsForPlan(buildPlan);
    const allIndexRecords = new Map(plannedVolumeRecords.flatMap((planned) => [...planned.indexRecordsById]));
    const firstLocalIds = new Set(buildPlan.volumes[0].items.map((item) => item.id));
    const rows = completeIndexRows(buildPlan.indexNodes, allIndexRecords, firstLocalIds, buildPlan.multiVolume);
    convergedIndexLayout = createAuthoritativeIndexLayout(rows, indexGeometry, bodyRegular, bodyBold);
    const state = `${completeIndexPages}->${convergedIndexLayout.pageCount}:${buildPlan.volumes.map((volume) => volume.items.map((item) => item.id).join(",")).join("|")}:${rows.filter((row): row is Extract<IndexLayoutRowInput, { kind: "exhibit" }> => row.kind === "exhibit").map((row) => row.pageLabel).join("|")}`;
    if (convergedIndexLayout.pageCount === completeIndexPages) break;
    if (convergenceStates.has(state)) throw new Error("Index and volume planning did not converge to a stable page count.");
    convergenceStates.add(state);
    completeIndexPages = convergedIndexLayout.pageCount;
    if (iteration === 15) throw new Error("Index and volume planning exceeded the 16-pass convergence safety limit.");
  }
  if (buildPlan.volumes.some((volume) => volume.indexPages !== convergedIndexLayout.pageCount)) {
    throw new Error("Index and volume planning ended with inconsistent complete-index page counts.");
  }
  if (buildPlan.multiVolume) report("Creating administrative volumes", `${buildPlan.volumes.length} separate PDFs for the same ${buildPlan.bundleIdentity} exhibit bundle`);

  const selectedIds = new Set(requested.map((candidate) => candidate.evidenceId).filter(Boolean));
  const selectedHashes = new Set(analysis.evidence.filter((record) => selectedIds.has(record.id)).map((record) => record.sha256));
  const excluded = analysis.evidence.filter((record) => !selectedIds.has(record.id));
  const effectiveIncludedIds = new Set(effectiveCandidates.filter((candidate) => candidate.included).map((candidate) => candidate.id));
  const exclusionResolutionFor = (candidate: ExhibitCandidate, source?: EvidenceRecord) => resolutions.find((resolution) =>
    (resolution.action === "exclude-candidate" && resolution.candidateId === candidate.id) ||
    (resolution.action === "exclude-source" && source && resolution.sourceId === source.id && resolution.sourceSha256 === source.sha256),
  );
  const omittedCitations = candidates
    .filter((candidate) => !candidate.manualAddition && !effectiveIncludedIds.has(candidate.id) && (candidate.included || exclusionResolutionFor(candidate, candidate.evidenceId ? analysis.evidence.find((record) => record.id === candidate.evidenceId) : undefined)))
    .map((candidate) => {
      const source = candidate.evidenceId ? analysis.evidence.find((record) => record.id === candidate.evidenceId) : undefined;
      const decision = exclusionResolutionFor(candidate, source);
      return { paragraph: candidate.paragraph, citation: candidate.citation, mark: candidate.mark, description: candidate.description, sourceFile: source?.name ?? null, sourceSha256: source?.sha256 ?? null, reason: decision?.note ?? (decision?.action === "exclude-source" ? "Reviewer excluded this source from the exhibit bundle." : decision?.action === "exclude-candidate" ? "Reviewer excluded this cited item from the exhibit bundle." : "This cited item was not included in the exhibit bundle."), decisionAction: decision?.action ?? null, decidedAt: decision?.approvedAt ?? null };
    });
  const technicalExceptions = resolutions
    .filter((resolution) => resolution.action === "proceed-without-ocr")
    .filter((resolution) => preflight.some((check) => check.policy === "exception-eligible" && check.severity === "warning" && findResolution([resolution], check)));
  const templateFallbacks = resolutions.filter((resolution) => resolution.action === "use-built-in-template" && (resolution.templateSlots ?? []).some((slot) => fallbackSlots.has(slot)));
  const excludedFiles = excluded.map((record) => {
    const sourceResolution = resolutions.find((resolution) => resolution.action === "exclude-source" && resolution.sourceId === record.id && resolution.sourceSha256 === record.sha256);
    const candidateResolution = candidates.filter((candidate) => candidate.evidenceId === record.id).map((candidate) => exclusionResolutionFor(candidate, record)).find(Boolean);
    return { fileName: record.name, sha256: record.sha256, reason: sourceResolution?.note ?? candidateResolution?.note ?? (selectedHashes.has(record.sha256) ? "Unselected duplicate physical copy of a selected source" : "No confirmed citation match") };
  });
  const completeIndexRecordMap = new Map(plannedVolumeRecords.flatMap((planned) => [...planned.indexRecordsById]));

  const finalizeVolume = async (volume: PlannedBuildVolume) => {
    const planned = plannedVolumeRecords[volume.number - 1];
    if (!planned || planned.volumeNumber !== volume.number) throw new Error(`Build plan lost the record map for ${volume.label}.`);
    const { numberOffset, preliminaryOffset, firstContentPage, records, recordsById } = planned;
    const visibleVolumeLabel = `Volume ${volume.number} of ${buildPlan.volumes.length}`;
    const finalDocument = await PDFDocument.create();
    finalDocument.setTitle(`${bundleTitle} - ${volume.label}`);
    finalDocument.setSubject("Locally generated, searchable exhibit bundle");
    finalDocument.setProducer("Exhibit Builder");
    finalDocument.setCreator("Exhibit Builder");
    const regular = await finalDocument.embedFont(StandardFonts.Helvetica);
    const bold = await finalDocument.embedFont(StandardFonts.HelveticaBold);
    if (cover) {
      const coverPlacement = await appendTemplatePages(finalDocument, cover);
      if (coverWritesMatterText(layout)) {
        applyTemplateMatterPatches(finalDocument.getPages(), coverPlacement, cover.reviewState?.matterReview, cover.reviewState?.matterConfirmation?.patches, { regular, bold });
      }
      if (buildPlan.multiVolume && coverPrintsVolumeLabel(layout)) finalDocument.getPages()[0].drawText(visibleVolumeLabel, { x: 44, y: 26, size: 9, font: bold, color: rgb(0.2, 0.22, 0.26) });
    } else {
      const page = finalDocument.addPage([A4_WIDTH, A4_HEIGHT]);
      const panel = coverMatterValues ?? indexMatterValues ?? layout.builtInMatter;
      page.drawText("EXHIBIT BUNDLE", { x: 48, y: 760, size: 11, font: bold, color: rgb(0.56, 0.18, 0.12) });
      drawWrapped(page, bundleTitle, { x: 48, y: 700, size: 22, font: bold, maxWidth: 495, lineHeight: 28, color: rgb(0.1, 0.14, 0.2) });
      let panelY = 640;
      const partyLine = (panel?.partyNames ?? []).map((item) => item.trim()).filter(Boolean).join("  v  ");
      if (partyLine && partyLine !== bundleTitle) {
        drawWrapped(page, partyLine, { x: 48, y: panelY, size: 13, font: regular, maxWidth: 495, lineHeight: 18, color: rgb(0.12, 0.16, 0.22) });
        panelY -= 40;
      }
      const numberLine = (panel?.matterNumbers ?? []).map((item) => item.trim()).filter(Boolean).join("  ");
      if (numberLine) {
        page.drawText(numberLine, { x: 48, y: panelY, size: 12, font: regular, color: rgb(0.12, 0.16, 0.22) });
        panelY -= 22;
      }
      const forumLine = (panel?.forums ?? []).map((item) => item.trim()).filter(Boolean).join("  ");
      if (forumLine) {
        drawWrapped(page, forumLine, { x: 48, y: panelY, size: 12, font: regular, maxWidth: 495, lineHeight: 16, color: rgb(0.12, 0.16, 0.22) });
        panelY -= 24;
      }
      if (buildPlan.multiVolume) page.drawText(visibleVolumeLabel, { x: 48, y: panelY === 640 ? 620 : panelY, size: 16, font: regular, color: rgb(0.12, 0.16, 0.22) });
    }
    const coverPageCount = finalDocument.getPageCount();
    if (coverPageCount !== volume.coverPages) throw new Error(`Build plan drifted while creating ${volume.label}: cover page count changed.`);

    const localItemIds = new Set(volume.items.map((item) => item.id));
    const indexRows = completeIndexRows(buildPlan.indexNodes, completeIndexRecordMap, localItemIds, buildPlan.multiVolume);
    const indexLayout = createAuthoritativeIndexLayout(indexRows, indexGeometry, regular, bold);
    if (indexLayout.pageCount !== volume.indexPages) throw new Error(`Build plan drifted while creating ${volume.label}: complete index page count changed.`);
    const indexPages: PDFPage[] = [];
    const indexPlacements: TemplatePagePlacement[] = [];
    for (let index = 0; index < volume.indexPages; index += 1) {
      const before = finalDocument.getPageCount();
      if (indexTemplate) {
        indexPlacements.push(...await appendTemplatePages(finalDocument, indexTemplate));
        indexPages.push(finalDocument.getPages()[before]);
      } else indexPages.push(finalDocument.addPage([A4_WIDTH, A4_HEIGHT]));
    }
    if (indexTemplate) {
      applyTemplateMatterPatches(finalDocument.getPages(), indexPlacements, indexTemplate.reviewState?.matterReview, indexTemplate.reviewState?.matterConfirmation?.patches, { regular, bold });
    }
    if (finalDocument.getPageCount() + 1 !== firstContentPage) throw new Error(`Build plan drifted while creating ${volume.label}: preliminary page count changed.`);
    for (const item of volume.items) {
      const indices = Array.from({ length: item.bodyEndPage - item.bodyStartPage + 1 }, (_, offset) => item.bodyStartPage - 1 + offset);
      const copied = await finalDocument.copyPages(body, indices);
      copied.forEach((page) => finalDocument.addPage(page));
    }
    const pages = finalDocument.getPages();
    for (let pageIndex = 0; pageIndex < indexPages.length; pageIndex += 1) {
      const page = indexPages[pageIndex];
      if (!indexTemplate) {
        page.drawText("EXHIBIT BUNDLE", { x: 44, y: 778, size: 10, font: bold, color: rgb(0.56, 0.18, 0.12) });
        drawWrapped(page, bundleTitle, { x: 44, y: 742, size: 18, font: bold, maxWidth: 507, lineHeight: 23, color: rgb(0.1, 0.14, 0.2) });
        const indexHeading = buildPlan.multiVolume ? `Complete index of exhibits - ${visibleVolumeLabel}` : "Index of exhibits";
        page.drawText(`${indexHeading}${indexPages.length > 1 ? ` (${pageIndex + 1} of ${indexPages.length})` : ""}`, { x: 44, y: 684, size: 11, font: regular, color: rgb(0.35, 0.38, 0.42) });
        page.drawLine({ start: { x: 44, y: 661 }, end: { x: 551, y: 661 }, thickness: 1, color: rgb(0.18, 0.23, 0.31) });
        const headerY = 646;
        const cols = indexLayout.geometry.columns;
        page.drawText("No.", { x: cols.exhibit.x + 4, y: headerY, size: 8, font: bold, color: rgb(0.35, 0.38, 0.42) });
        if (cols.date) page.drawText("Date", { x: cols.date.x + 4, y: headerY, size: 8, font: bold, color: rgb(0.35, 0.38, 0.42) });
        page.drawText("Description", { x: cols.description.x + 4, y: headerY, size: 8, font: bold, color: rgb(0.35, 0.38, 0.42) });
        const pageHeader = "Page";
        page.drawText(pageHeader, {
          x: cols.pageReference.x + cols.pageReference.width - 4 - bold.widthOfTextAtSize(pageHeader, 8),
          y: headerY,
          size: 8,
          font: bold,
          color: rgb(0.35, 0.38, 0.42),
        });
      }
      drawAuthoritativeIndexPage(page, indexLayout, pageIndex + 1, regular, bold);
      for (const row of indexLayout.rows.filter((candidate): candidate is PlannedIndexExhibit => candidate.kind === "exhibit" && candidate.pageNumber === pageIndex + 1 && Boolean(candidate.linkTargetId))) {
        const destination = row.linkTargetId ? recordsById.get(row.linkTargetId) : undefined;
        if (!destination || !row.linkRectangle) throw new Error(`Index layout created an invalid local link for ${row.id}.`);
        addIndexLink(finalDocument, page, pages[destination.startPage - 1], row.linkRectangle);
      }
    }
    const outlineNodes = localOutlineNodes(buildPlan.indexNodes, recordsById);
    addOutlineTree(finalDocument, outlineNodes, pages);
    const expectedPageLabels = addPageLabels(finalDocument, firstContentPage, pagination, records, numberOffset, preliminaryOffset, layout.countOptionalPagesInReferences);
    const skipExactCoverNumbers = Boolean(cover && !coverPrintsPageNumber(layout));
    pages.forEach((page, index) => {
      if (skipExactCoverNumbers && index < coverPageCount) return;
      const pageNumber = index + 1;
      const label = expectedPageLabels[index];
      if (!label) return;
      const availableWidth = Math.max(0, page.getWidth() - 84);
      const labelWidth = regular.widthOfTextAtSize(label, pagination.fontSize);
      if (labelWidth > availableWidth) {
        throw new Error(`Page label "${label}" cannot fit safely within the page-number area at ${pagination.fontSize} pt. Shorten the prefix or suffix.`);
      }
      const position = pageNumberPosition(page, pageNumber, pagination.position, label, regular, pagination.fontSize);
      page.drawText(label, { ...position, size: pagination.fontSize, font: regular, color: rgb(0.25, 0.27, 0.3) });
    });
    report("Applying page numbers, bookmarks and links", volume.label);
    const bytes = await finalDocument.save({ useObjectStreams: false });
    report("Reopening and validating the finished PDF", volume.label);
    const reopened = await PDFDocument.load(bytes);
    if (reopened.getPageCount() !== volume.totalPages) throw new Error(`Build plan drifted while creating ${volume.label}: expected ${volume.totalPages} pages and produced ${reopened.getPageCount()}.`);
    const orientation = reopened.getPages().reduce((counts, page) => {
      const size = page.getSize();
      const rotation = ((page.getRotation().angle % 360) + 360) % 360;
      const { width, height } = rotation === 90 || rotation === 270 ? { width: size.height, height: size.width } : size;
      if (Math.abs(width - A4_WIDTH) < 0.02 && Math.abs(height - A4_HEIGHT) < 0.02) counts.portrait += 1;
      else if (Math.abs(width - A4_HEIGHT) < 0.02 && Math.abs(height - A4_WIDTH) < 0.02) counts.landscape += 1;
      else counts.nonA4 += 1;
      return counts;
    }, { portrait: 0, landscape: 0, nonA4: 0 });
    const allPagesA4 = orientation.nonA4 === 0;
    const expectedIndexLinks = indexLayout.rows.flatMap((row) => {
      if (row.kind !== "exhibit" || !row.linkTargetId) return [];
      const record = recordsById.get(row.linkTargetId);
      if (!record) throw new Error(`Output validation could not find local index destination ${row.linkTargetId}.`);
      return [{ id: row.linkTargetId, targetPage: record.startPage }];
    });
    const navigation = validateFinalDestinations(reopened, expectedIndexLinks, coverPageCount + 1, volume.indexPages, outlineNodes);
    const hasOutlines = Boolean(reopened.catalog.get(PDFName.of("Outlines")));
    const verifiedPageLabels = validatePageLabels(reopened, expectedPageLabels);
    if (!hasOutlines) throw new Error(`Output validation failed for ${volume.label}: index links or bookmarks could not be verified.`);
    const hash = await sha256(bytes);
    const checks = [
      ...preflight.map((check) => ({ label: check.label, status: check.severity, detail: check.detail })),
      ...(technicalExceptions.length ? [{ label: "Approved technical exceptions", status: "warning" as const, detail: `${technicalExceptions.length} source exception${technicalExceptions.length === 1 ? "" : "s"} approved; see the manifest.` }] : []),
      ...(omittedCitations.length ? [{ label: "Cited material omitted by reviewer", status: "warning" as const, detail: `${omittedCitations.length} cited reference${omittedCitations.length === 1 ? "" : "s"} intentionally excluded.` }] : []),
      ...(templateFallbacks.length ? [{ label: "Built-in template fallback", status: "warning" as const, detail: "One or more selected custom templates were replaced by the approved built-in layout." }] : []),
      ...(volume.oversize ? [{ label: "Oversize exhibit volume", status: "warning" as const, detail: `${volume.label} is ${volume.totalPages} pages and exceeds the ${layout.volumePageLimit}-page advisory limit because an exhibit is kept intact.` }] : []),
      { label: "Complete repeated index", status: "pass" as const, detail: `${buildPlan.canonicalOrder.length} exhibits across ${indexLayout.pageCount} index page${indexLayout.pageCount === 1 ? "" : "s"}${buildPlan.multiVolume ? "; every row identifies its owning volume." : "."}` },
      { label: "Hyperlinked index", status: "pass" as const, detail: `${navigation.links} local destinations verified after reopening; ${buildPlan.canonicalOrder.length - navigation.links} non-local row${buildPlan.canonicalOrder.length - navigation.links === 1 ? "" : "s"} intentionally left without invalid links.` },
      { label: "PDF bookmarks", status: "pass" as const, detail: `${navigation.bookmarks} hierarchical bookmark${navigation.bookmarks === 1 ? "" : "s"} verified after reopening` },
      { label: "PDF page labels", status: "pass" as const, detail: `${verifiedPageLabels} explicit ${pageLabelSchemeName(pagination).toLowerCase()} labels verified after reopening` },
      { label: "A4 page treatment", status: allPagesA4 ? "pass" as const : "warning" as const, detail: allPagesA4 ? `Every page is A4 (${orientation.portrait} portrait, ${orientation.landscape} landscape).` : `${orientation.nonA4} page(s) retain their original non-A4 size by reviewer choice.` },
    ];
    const manifest = {
      schemaVersion: "3.0",
      caseTitle: bundleTitle,
      createdAt: new Date().toISOString(),
      pagination,
      layout,
      pageSizeChoices: Object.fromEntries(Object.entries(pageSizeChoices).filter(([, choice]) => choice === "keep-original")),
      buildPlan,
      indexLayout: { schemaVersion: indexLayout.schemaVersion, geometry: indexLayout.geometry.id, pageCount: indexLayout.pageCount, rowCount: indexLayout.rows.length },
      volume: { number: volume.number, label: volume.label, pageLimit: layout.volumePageLimit, oversize: volume.oversize },
      statement: { fileName: analysis.statementName, sha256: analysis.statementHash, modified: false },
      exhibits: records,
      manualAdditions: records.filter((record) => record.manualAddition).map((record) => ({ exhibitNumber: record.exhibitNumber, description: record.description, documentDate: record.documentDate, sourceFile: record.fileName, sourceSha256: record.sourceHash, addedAt: record.manualAddedAt, warningAcknowledgedAt: record.manualWarningAcknowledgedAt, citationStatus: record.citationStatus })),
      templates: templateManifestRecords(),
      technicalExceptions: technicalExceptions.map((resolution) => ({ ...resolution, originalSeverity: "blocking", resultingStatus: "warning" })),
      templateFallbacks: templateFallbacks.map((resolution) => ({ ...resolution, resultingLayout: "built-in" })),
      omittedCitations,
      excludedFiles,
      output: { fileName: volume.fileName, sha256: hash, pageCount: reopened.getPageCount(), pageSize: allPagesA4 ? "A4" : "mixed", orientation },
    };
    return { number: volume.number, label: volume.label, bytes, fileName: volume.fileName, sha256: hash, pageCount: reopened.getPageCount(), records, manifest, checks };
  };

  const volumes = [] as NonNullable<BuildResult["volumes"]>;
  for (const volume of buildPlan.volumes) volumes.push(await finalizeVolume(volume));
  const allRecords = volumes.flatMap((volume) => volume.records);
  if (!buildPlan.multiVolume) {
    report("Bundle complete");
    return { ...volumes[0], buildPlan };
  }
  const archive = await createVolumeArchive(volumes, buildPlan.bundleIdentity);
  const volumeZipBytes = archive.bytes;
  const volumeZipSha256 = archive.sha256;
  const volumeZipFileName = archive.fileName;
  const combinedManifest = {
    schemaVersion: "3.0",
    caseTitle: bundleTitle,
    createdAt: new Date().toISOString(),
    pagination,
    layout,
    buildPlan,
    indexLayout: { schemaVersion: convergedIndexLayout.schemaVersion, geometry: convergedIndexLayout.geometry.id, pageCount: convergedIndexLayout.pageCount, rowCount: convergedIndexLayout.rows.length },
    templates: templateManifestRecords(),
    statement: { fileName: analysis.statementName, sha256: analysis.statementHash, modified: false },
    exhibits: allRecords,
    manualAdditions: allRecords.filter((record) => record.manualAddition).map((record) => ({ exhibitNumber: record.exhibitNumber, description: record.description, documentDate: record.documentDate, sourceFile: record.fileName, sourceSha256: record.sourceHash, addedAt: record.manualAddedAt, warningAcknowledgedAt: record.manualWarningAcknowledgedAt, citationStatus: record.citationStatus })),
    volumes: volumes.map((volume) => ({ number: volume.number, label: volume.label, fileName: volume.fileName, sha256: volume.sha256, pageCount: volume.pageCount, checks: volume.checks })),
    technicalExceptions,
    templateFallbacks,
    omittedCitations,
    excludedFiles,
    output: { fileName: volumeZipFileName, sha256: volumeZipSha256, pageCount: volumes.reduce((total, volume) => total + volume.pageCount, 0), volumeCount: volumes.length },
  };
  report("Bundle complete");
  return {
    bytes: volumeZipBytes,
    fileName: volumeZipFileName,
    sha256: volumeZipSha256,
    pageCount: volumes.reduce((total, volume) => total + volume.pageCount, 0),
    records: allRecords,
    manifest: combinedManifest,
    checks: volumes.flatMap((volume) => volume.checks.map((check) => ({ ...check, label: `${volume.label}: ${check.label}` }))),
    volumes,
    buildPlan,
    volumeZipBytes,
    volumeZipFileName,
    volumeZipSha256,
  };
}

export function downloadBytes(
  bytes: Uint8Array,
  fileName: string,
  type: string,
): Promise<{ saved: boolean; filePath?: string }> {
  if (window.bundleBuilderDesktop) {
    return window.bundleBuilderDesktop.saveFile(bytes, fileName, type);
  }
  const blob = new Blob([bytes as BlobPart], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  return Promise.resolve({ saved: true });
}

export function downloadJson(value: unknown, fileName: string) {
  const bytes = new TextEncoder().encode(JSON.stringify(value, null, 2));
  void downloadBytes(bytes, fileName, "application/json");
}
