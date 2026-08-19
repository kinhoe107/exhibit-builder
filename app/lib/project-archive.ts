import JSZip from "jszip";
import { bundleArrangementFromLegacyOrder, validateBundleArrangement } from "./bundle-arrangement.ts";
import { DEFAULT_BUNDLE_LAYOUT, DEFAULT_PAGINATION, type ProjectSnapshot, type ProjectSource, type StoredTemplateReview, type TemplateDiscrepancyConfirmation } from "./bundle-types.ts";

const PROJECT_FILE = "bundle-project.json";
const SOURCES_DIR = "sources";
// A saved project contains copies of locally selected source files. The same
// bounded limit is enforced before saving and opening so a valid save can
// always be reopened by this application.
const ARCHIVE_LIMITS = {
  bytes: 192 * 1024 * 1024,
  entries: 1_000,
  inflated: 256 * 1024 * 1024,
  source: 128 * 1024 * 1024,
  projectJson: 4 * 1024 * 1024,
  inflationRatio: 50,
} as const;
const HASH = /^[a-f0-9]{64}$/i;
const SOURCE_ROLES = new Set<ProjectSource["role"]>(["statement", "evidence", "template", "template-rendered"]);
const TEMPLATE_SLOTS = new Set(["cover", "exhibitCover", "index", "divider"]);
const TEMPLATE_FORMATS = new Set(["pdf", "docx", "doc"]);
const FINDING_KINDS = new Set(["matter-number", "party-name", "forum", "matter-title", "placeholder"]);

type StoredSource = {
  id: string;
  role: ProjectSource["role"];
  name: string;
  sha256: string;
  path: string;
};

type StoredProject = ProjectSnapshot & { sources: StoredSource[] };

function safeName(name: string) {
  return name.replace(/[^a-z0-9._ -]/gi, "_").replace(/\.{2,}/g, "_");
}

function unsafeArchivePath(value: string) {
  if (!value || value.includes("\\") || value.startsWith("/") || /^[a-z]:/i.test(value)) return true;
  return value.split("/").some((segment) => !segment || segment === "." || segment === "..");
}

function plainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value) && Object.getPrototypeOf(value) === Object.prototype;
}

function onlyKeys(value: Record<string, unknown>, allowed: readonly string[]) {
  const keys = new Set(allowed);
  return Object.keys(value).every((key) => keys.has(key));
}

function boundedText(value: unknown, maximum: number) {
  return typeof value === "string" && value.length <= maximum && !value.includes("\0");
}

function validConfirmation(value: unknown, pdfSha256: string) {
  if (!plainObject(value) || !onlyKeys(value, ["pdfSha256", "confirmedAt"])) return false;
  return value.pdfSha256 === pdfSha256 && typeof value.confirmedAt === "string" && boundedText(value.confirmedAt, 64) && !Number.isNaN(Date.parse(value.confirmedAt));
}

function validMatterValueList(value: unknown) {
  return Array.isArray(value) && value.length <= 20 && value.every((item) => boundedText(item, 240));
}

function validMatterGeometry(value: unknown, pageCount: number) {
  if (!plainObject(value) || !onlyKeys(value, ["pageNumber", "x", "y", "width", "height", "fontSize", "color"])) return false;
  if (!Number.isSafeInteger(value.pageNumber) || value.pageNumber < 1 || value.pageNumber > pageCount) return false;
  if (![value.x, value.y, value.width, value.height, value.fontSize].every((item) => typeof item === "number" && Number.isFinite(item))) return false;
  if (value.color === undefined) return true;
  return plainObject(value.color) && onlyKeys(value.color, ["r", "g", "b"]) && [value.color.r, value.color.g, value.color.b].every((item) => typeof item === "number" && Number.isFinite(item));
}

function validMatterPatch(value: unknown) {
  return plainObject(value) && onlyKeys(value, ["findingId", "value"]) && boundedText(value.findingId, 240) && boundedText(value.value, 240);
}

function validMatterConfirmation(value: unknown, pdfSha256: string) {
  if (!plainObject(value) || !onlyKeys(value, ["pdfSha256", "confirmedAt", "matterNumbers", "partyNames", "forums", "matterTitles", "patches"])) return false;
  if (value.pdfSha256 !== pdfSha256 || typeof value.confirmedAt !== "string" || !boundedText(value.confirmedAt, 64) || Number.isNaN(Date.parse(value.confirmedAt))) return false;
  if (value.patches !== undefined && (!Array.isArray(value.patches) || value.patches.length > 100 || value.patches.some((patch) => !validMatterPatch(patch)))) return false;
  return (["matterNumbers", "partyNames", "forums", "matterTitles"] as const).every((property) => value[property] === undefined || validMatterValueList(value[property]));
}

function validFinding(value: unknown, expectedKind: string, pageCount: number) {
  if (!plainObject(value) || !onlyKeys(value, ["id", "kind", "value", "normalizedValue", "pageNumbers", "evidence", "geometry", "unverified"])) {
    if (!plainObject(value) || !onlyKeys(value, ["kind", "value", "normalizedValue", "pageNumbers", "evidence", "unverified"])) return false;
  }
  if (value.kind !== expectedKind || !FINDING_KINDS.has(expectedKind)) return false;
  if (value.id !== undefined && !boundedText(value.id, 240)) return false;
  if (value.geometry !== undefined && !validMatterGeometry(value.geometry, pageCount)) return false;
  return boundedText(value.value, 1_024) && boundedText(value.normalizedValue, 1_024)
    && Array.isArray(value.pageNumbers) && value.pageNumbers.length <= 25 && value.pageNumbers.every((page) => Number.isSafeInteger(page) && page >= 1 && page <= pageCount)
    && Array.isArray(value.evidence) && value.evidence.length <= 25 && value.evidence.every((item) => boundedText(item, 2_048))
    && value.unverified === true;
}

function validMatterReview(value: unknown, pdfSha256: string) {
  if (!plainObject(value) || !onlyKeys(value, ["sourceName", "pdfSha256", "exactByteLength", "pageCount", "extractedCharacterCount", "textReliability", "requiresVisualConfirmation", "notice", "matterNumbers", "partyNames", "forums", "matterTitles", "placeholders"])) return false;
  const pageCount = typeof value.pageCount === "number" ? value.pageCount : Number.NaN;
  const exactByteLength = typeof value.exactByteLength === "number" ? value.exactByteLength : Number.NaN;
  const extractedCharacterCount = typeof value.extractedCharacterCount === "number" ? value.extractedCharacterCount : Number.NaN;
  if (!boundedText(value.sourceName, 512) || value.pdfSha256 !== pdfSha256 || !Number.isSafeInteger(exactByteLength) || exactByteLength < 1 || exactByteLength > 25 * 1024 * 1024 || !Number.isSafeInteger(pageCount) || pageCount < 1 || pageCount > 25 || !Number.isSafeInteger(extractedCharacterCount) || extractedCharacterCount < 0 || extractedCharacterCount > 250_000 || !["reliable", "limited", "none"].includes(String(value.textReliability)) || typeof value.requiresVisualConfirmation !== "boolean" || !boundedText(value.notice, 2_048)) return false;
  for (const [property, kind] of [["matterNumbers", "matter-number"], ["partyNames", "party-name"], ["forums", "forum"], ["matterTitles", "matter-title"], ["placeholders", "placeholder"]] as const) {
    const findings = value[property];
    if (!Array.isArray(findings) || findings.length > 100 || findings.some((finding) => !validFinding(finding, kind, pageCount))) return false;
  }
  return true;
}

function validStoredTemplateReview(value: unknown): value is StoredTemplateReview {
  if (!plainObject(value) || !onlyKeys(value, ["slot", "sourceId", "renderedSourceId", "sourceFormat", "sourceSha256", "pdfSha256", "reviewState"])) return false;
  if (typeof value.slot !== "string" || !TEMPLATE_SLOTS.has(value.slot) || !boundedText(value.sourceId, 256) || (value.renderedSourceId !== undefined && !boundedText(value.renderedSourceId, 256)) || typeof value.sourceFormat !== "string" || !TEMPLATE_FORMATS.has(value.sourceFormat) || typeof value.sourceSha256 !== "string" || !HASH.test(value.sourceSha256) || typeof value.pdfSha256 !== "string" || !HASH.test(value.pdfSha256) || !plainObject(value.reviewState) || !onlyKeys(value.reviewState, ["matterReview", "appearanceConfirmation", "matterConfirmation", "placeholderConfirmation"]) || !validMatterReview(value.reviewState.matterReview, value.pdfSha256)) return false;
  return value.reviewState.appearanceConfirmation === undefined || validConfirmation(value.reviewState.appearanceConfirmation, value.pdfSha256)
    ? (value.reviewState.matterConfirmation === undefined || validMatterConfirmation(value.reviewState.matterConfirmation, value.pdfSha256))
      && (value.reviewState.placeholderConfirmation === undefined || validConfirmation(value.reviewState.placeholderConfirmation, value.pdfSha256))
    : false;
}

function validTemplateDiscrepancyConfirmation(value: unknown): value is TemplateDiscrepancyConfirmation {
  return plainObject(value) && onlyKeys(value, ["fingerprint", "confirmedAt"])
    && boundedText(value.fingerprint, 100_000) && typeof value.confirmedAt === "string" && boundedText(value.confirmedAt, 64) && !Number.isNaN(Date.parse(value.confirmedAt));
}

function entrySize(entry: JSZip.JSZipObject) {
  return Number((entry as any)._data?.uncompressedSize ?? 0);
}

function validStoredSource(value: unknown): value is StoredSource {
  if (!value || typeof value !== "object") return false;
  const source = value as Partial<StoredSource>;
  return typeof source.id === "string" && source.id.length > 0 && source.id.length <= 256
    && typeof source.name === "string" && source.name.length > 0 && source.name.length <= 512 && !source.name.includes("\0")
    && typeof source.sha256 === "string" && HASH.test(source.sha256)
    && typeof source.path === "string" && source.path.startsWith(`${SOURCES_DIR}/`) && !unsafeArchivePath(source.path)
    && typeof source.role === "string" && SOURCE_ROLES.has(source.role as ProjectSource["role"]);
}

async function digest(bytes: ArrayBuffer) {
  const value = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(value)).map((part) => part.toString(16).padStart(2, "0")).join("");
}

export async function createProjectArchive(
  snapshot: ProjectSnapshot,
  sources: ProjectSource[],
) {
  const sourceBytes = sources.reduce((total, source) => total + source.file.size, 0);
  if (sources.length > ARCHIVE_LIMITS.entries - 1 || sourceBytes > ARCHIVE_LIMITS.inflated || sources.some((source) => !source.file.size || source.file.size > ARCHIVE_LIMITS.source || !HASH.test(source.sha256))) {
    throw new Error("This exhibit project is too large to save safely as one local project file.");
  }
  const zip = new JSZip();
  const storedSources: StoredSource[] = [];
  const usedIds = new Set<string>();
  const usedPaths = new Set<string>();
  for (const source of sources) {
    if (!source.id || usedIds.has(source.id) || !SOURCE_ROLES.has(source.role)) throw new Error("This exhibit project contains an invalid or duplicate source descriptor.");
    const path = `${SOURCES_DIR}/${safeName(source.id)}-${safeName(source.name)}`;
    if (unsafeArchivePath(path) || usedPaths.has(path)) throw new Error("This exhibit project contains duplicate or unsafe source paths.");
    usedIds.add(source.id);
    usedPaths.add(path);
    zip.file(path, await source.file.arrayBuffer());
    storedSources.push({
      id: source.id,
      role: source.role,
      name: source.name,
      sha256: source.sha256,
      path,
    });
  }
  const arrangement = snapshot.schemaVersion === 8
    ? validateBundleArrangement(snapshot.arrangement)
    : bundleArrangementFromLegacyOrder(snapshot.finalOrder);
  // finalOrder is accepted only as a schemas 2-7 migration input. New archives
  // have one authoritative order representation and cannot drift between two.
  const { finalOrder: _legacyFinalOrder, ...currentSnapshot } = snapshot;
  zip.file(PROJECT_FILE, JSON.stringify({ ...currentSnapshot, schemaVersion: 8, arrangement, sources: storedSources }, null, 2));
  const archive = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  if (archive.byteLength > ARCHIVE_LIMITS.bytes) {
    throw new Error("This exhibit project is too large to save safely as one local project file.");
  }
  return archive;
}

export async function openProjectArchive(file: File): Promise<{
  snapshot: ProjectSnapshot;
  sources: ProjectSource[];
}> {
  if (!file.size || file.size > ARCHIVE_LIMITS.bytes) throw new Error("Exhibit project is empty or exceeds the 192 MB archive safety limit.");
  let zip: JSZip;
  try {
    let timer: ReturnType<typeof setTimeout> | undefined;
    zip = await Promise.race([
      JSZip.loadAsync(await file.arrayBuffer(), { createFolders: false }),
      new Promise<never>((_resolve, reject) => { timer = setTimeout(() => reject(new Error("timeout")), 30_000); }),
    ]).finally(() => { if (timer) clearTimeout(timer); });
  } catch (caught) {
    if (caught instanceof Error && caught.message === "timeout") throw new Error("Exhibit project archive did not open within the 30-second safety limit.");
    throw new Error("Exhibit project archive is malformed.");
  }
  const entries = Object.values(zip.files);
  const inflated = entries.reduce((sum, entry) => sum + entrySize(entry), 0);
  if (entries.length > ARCHIVE_LIMITS.entries || inflated > ARCHIVE_LIMITS.inflated || inflated / Math.max(file.size, 1) > ARCHIVE_LIMITS.inflationRatio || entries.some((entry) => !entry.dir && (entrySize(entry) > ARCHIVE_LIMITS.source || unsafeArchivePath((entry as any).unsafeOriginalName ?? entry.name) || unsafeArchivePath(entry.name) || (!entry.name.startsWith(`${SOURCES_DIR}/`) && entry.name !== PROJECT_FILE)))) throw new Error("Exhibit project archive has unsafe paths or limits.");
  const projectEntry = zip.file(PROJECT_FILE);
  if (!projectEntry) throw new Error("This is not an Exhibit Builder project file.");
  if (entrySize(projectEntry) > ARCHIVE_LIMITS.projectJson) throw new Error("Exhibit project metadata exceeds the 4 MB safety limit.");
  let stored: StoredProject;
  try { stored = JSON.parse(await projectEntry.async("text")) as StoredProject; } catch { throw new Error("Exhibit project metadata is not valid JSON."); }
  if ((stored.schemaVersion !== 2 && stored.schemaVersion !== 3 && stored.schemaVersion !== 4 && stored.schemaVersion !== 5 && stored.schemaVersion !== 6 && stored.schemaVersion !== 7 && stored.schemaVersion !== 8) || !Array.isArray(stored.sources)) {
    throw new Error("This exhibit project uses an unsupported format.");
  }
  if (stored.schemaVersion === 8 && (stored.templateReviews !== undefined && (!Array.isArray(stored.templateReviews) || stored.templateReviews.length > 4 || stored.templateReviews.some((review) => !validStoredTemplateReview(review))))) throw new Error("Exhibit project contains invalid template-review metadata.");
  if (stored.schemaVersion === 8 && stored.templateDiscrepancyConfirmation !== undefined && !validTemplateDiscrepancyConfirmation(stored.templateDiscrepancyConfirmation)) throw new Error("Exhibit project contains invalid template-discrepancy confirmation metadata.");
  let arrangement;
  try {
    arrangement = stored.schemaVersion === 8
      ? validateBundleArrangement(stored.arrangement)
      : bundleArrangementFromLegacyOrder(stored.finalOrder);
  } catch {
    throw new Error("Exhibit project contains an invalid bundle arrangement.");
  }
  if (stored.sources.length > ARCHIVE_LIMITS.entries - 1 || stored.sources.some((source) => !validStoredSource(source))) throw new Error("Exhibit project contains an invalid source descriptor.");
  const ids = new Set<string>();
  const paths = new Set<string>();
  for (const source of stored.sources) {
    if (ids.has(source.id) || paths.has(source.path)) throw new Error("Exhibit project contains duplicate source IDs or paths.");
    ids.add(source.id);
    paths.add(source.path);
  }
  const sources: ProjectSource[] = [];
  // Expand and verify one source at a time. This prevents many compressed
  // entries from allocating their full contents concurrently.
  for (const source of stored.sources) {
    const entry = zip.file(source.path);
    if (!entry) throw new Error(`Exhibit project is missing ${source.name}.`);
    if (entrySize(entry) > ARCHIVE_LIMITS.source) throw new Error(`${source.name} exceeds the 128 MB source safety limit.`);
    const bytes = await entry.async("arraybuffer");
    if (await digest(bytes) !== source.sha256) {
      throw new Error(`Exhibit project integrity check failed for ${source.name}.`);
    }
    sources.push({
      id: source.id,
      role: source.role,
      name: source.name,
      sha256: source.sha256,
      file: new File([bytes], source.name),
    });
  }
  const { sources: _sources, ...snapshot } = stored;
  const { finalOrder: _legacyFinalOrder, ...currentSnapshot } = snapshot;
  return { snapshot: { ...currentSnapshot, schemaVersion: 8, arrangement, sheetSelections: snapshot.sheetSelections ?? [], resolutions: snapshot.resolutions ?? [], pagination: { ...DEFAULT_PAGINATION, ...(snapshot.pagination ?? {}) }, layout: { ...DEFAULT_BUNDLE_LAYOUT, ...(snapshot.layout ?? {}) } } as ProjectSnapshot, sources };
}
