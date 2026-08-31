const { createHash, randomUUID } = require("node:crypto");
const { open, mkdir, readFile, rename, rm, stat } = require("node:fs/promises");
const path = require("node:path");

const JOURNAL_NAME = "current-recovery.json";
const BACKUP_NAME = "previous-recovery.json";
const MAX_JSON_BYTES = 8 * 1024 * 1024;
const MAX_SOURCES = 5_000;
const MAX_CANDIDATES = 20_000;
const MAX_ORDER = 20_000;
const MAX_RESOLUTIONS = 20_000;
const HASH = /^[a-f0-9]{64}$/i;
const ROLES = new Set(["statement", "evidence", "template", "template-rendered", "project"]);
const TOP_LEVEL_KEYS = new Set(["project", "candidates", "arrangement", "finalOrder", "layout", "pagination", "resolutions", "statements", "templates", "templateReviews", "templateDiscrepancyConfirmation", "sources", "fingerprint", "pageSizeChoices"]);
const TEMPLATE_SLOTS = new Set(["cover", "exhibitCover", "index", "divider"]);
const TEMPLATE_FORMATS = new Set(["pdf", "docx", "doc"]);
const FINDING_KINDS = new Set(["matter-number", "party-name", "forum", "matter-title", "placeholder"]);

function plainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value) && Object.getPrototypeOf(value) === Object.prototype;
}

function boundedString(value, label, maximum = 4_096) {
  if (typeof value !== "string" || value.length > maximum) throw new Error(`${label} is invalid or too long.`);
  return value;
}

function validateSource(source) {
  if (!plainObject(source)) throw new Error("Recovery source descriptor is invalid.");
  const id = boundedString(source.id, "Recovery source ID", 512);
  const role = boundedString(source.role, "Recovery source role", 32);
  const name = boundedString(source.name, "Recovery source name", 1_024);
  const sourcePath = boundedString(source.path, "Recovery source path", 32_768);
  const sha256 = boundedString(source.sha256, "Recovery source hash", 64);
  if (!id || !ROLES.has(role) || !name || !path.isAbsolute(sourcePath) || !HASH.test(sha256)) throw new Error("Recovery source descriptor failed validation.");
  if (!Number.isSafeInteger(source.size) || source.size < 0 || source.size > 2 * 1024 * 1024 * 1024) throw new Error("Recovery source size is invalid.");
  return { id, role, name, path: sourcePath, sha256: sha256.toLowerCase(), size: source.size };
}

function hasOnlyKeys(value, allowed) {
  return Object.keys(value).every((key) => allowed.has(key));
}

function validDate(value) {
  return typeof value === "string" && value.length <= 64 && !Number.isNaN(Date.parse(value));
}

function validateTemplateConfirmation(value, pdfSha256) {
  if (!plainObject(value) || !hasOnlyKeys(value, new Set(["pdfSha256", "confirmedAt"])) || value.pdfSha256 !== pdfSha256 || !validDate(value.confirmedAt)) throw new Error("Recovery template confirmation is invalid.");
  return { pdfSha256, confirmedAt: value.confirmedAt };
}

function validateMatterConfirmation(value, pdfSha256) {
  if (!plainObject(value) || !hasOnlyKeys(value, new Set(["pdfSha256", "confirmedAt", "matterNumbers", "partyNames", "forums", "matterTitles", "patches"])) || value.pdfSha256 !== pdfSha256 || !validDate(value.confirmedAt)) throw new Error("Recovery template confirmation is invalid.");
  const confirmation = { pdfSha256, confirmedAt: value.confirmedAt };
  for (const property of ["matterNumbers", "partyNames", "forums", "matterTitles"]) {
    if (value[property] === undefined) continue;
    if (!Array.isArray(value[property]) || value[property].length > 20) throw new Error("Recovery template confirmation is invalid.");
    confirmation[property] = value[property].map((item) => boundedString(item, "Recovery confirmed matter detail", 240));
  }
  if (value.patches !== undefined) {
    if (!Array.isArray(value.patches) || value.patches.length > 100) throw new Error("Recovery template confirmation is invalid.");
    confirmation.patches = value.patches.map((patch) => {
      if (!plainObject(patch) || !hasOnlyKeys(patch, new Set(["findingId", "value"]))) throw new Error("Recovery template confirmation is invalid.");
      return { findingId: boundedString(patch.findingId, "Recovery matter patch id", 240), value: boundedString(patch.value, "Recovery matter patch", 240) };
    });
  }
  return confirmation;
}

function validateMatterGeometry(value, pageCount) {
  if (!plainObject(value) || !hasOnlyKeys(value, new Set(["pageNumber", "x", "y", "width", "height", "fontSize", "color"])) || !Number.isSafeInteger(value.pageNumber) || value.pageNumber < 1 || value.pageNumber > pageCount) throw new Error("Recovery template finding geometry is invalid.");
  if (![value.x, value.y, value.width, value.height, value.fontSize].every((item) => typeof item === "number" && Number.isFinite(item))) throw new Error("Recovery template finding geometry is invalid.");
  if (value.color === undefined) return { pageNumber: value.pageNumber, x: value.x, y: value.y, width: value.width, height: value.height, fontSize: value.fontSize };
  if (!plainObject(value.color) || !hasOnlyKeys(value.color, new Set(["r", "g", "b"])) || ![value.color.r, value.color.g, value.color.b].every((item) => typeof item === "number" && Number.isFinite(item))) throw new Error("Recovery template finding geometry is invalid.");
  return { pageNumber: value.pageNumber, x: value.x, y: value.y, width: value.width, height: value.height, fontSize: value.fontSize, color: { r: value.color.r, g: value.color.g, b: value.color.b } };
}

function validateTemplateFinding(value, expectedKind, pageCount) {
  const allowed = new Set(["id", "kind", "value", "normalizedValue", "pageNumbers", "evidence", "geometry", "unverified"]);
  if (!plainObject(value) || !hasOnlyKeys(value, allowed) || value.kind !== expectedKind || !FINDING_KINDS.has(expectedKind) || value.unverified !== true) throw new Error("Recovery template finding is invalid.");
  const findingValue = boundedString(value.value, "Recovery template finding", 1_024);
  const normalizedValue = boundedString(value.normalizedValue, "Recovery normalized template finding", 1_024);
  if (!Array.isArray(value.pageNumbers) || value.pageNumbers.length > 25 || value.pageNumbers.some((page) => !Number.isSafeInteger(page) || page < 1 || page > pageCount) || !Array.isArray(value.evidence) || value.evidence.length > 25) throw new Error("Recovery template finding evidence is invalid.");
  const finding = { kind: expectedKind, value: findingValue, normalizedValue, pageNumbers: value.pageNumbers, evidence: value.evidence.map((item) => boundedString(item, "Recovery template finding evidence", 2_048)), unverified: true };
  if (value.id !== undefined) finding.id = boundedString(value.id, "Recovery template finding id", 240);
  if (value.geometry !== undefined) finding.geometry = validateMatterGeometry(value.geometry, pageCount);
  return finding;
}

function validateTemplateReview(value) {
  if (!plainObject(value) || !hasOnlyKeys(value, new Set(["slot", "sourceId", "renderedSourceId", "sourceFormat", "sourceSha256", "pdfSha256", "reviewState"]))) throw new Error("Recovery template-review metadata is invalid.");
  const slot = boundedString(value.slot, "Recovery template slot", 32);
  const sourceId = boundedString(value.sourceId, "Recovery template source ID", 256);
  const renderedSourceId = value.renderedSourceId === undefined ? undefined : boundedString(value.renderedSourceId, "Recovery rendered-template source ID", 256);
  const sourceFormat = boundedString(value.sourceFormat, "Recovery template format", 8);
  const sourceSha256 = boundedString(value.sourceSha256, "Recovery template source hash", 64);
  const pdfSha256 = boundedString(value.pdfSha256, "Recovery template PDF hash", 64);
  if (!TEMPLATE_SLOTS.has(slot) || !TEMPLATE_FORMATS.has(sourceFormat) || !HASH.test(sourceSha256) || !HASH.test(pdfSha256) || !plainObject(value.reviewState) || !hasOnlyKeys(value.reviewState, new Set(["matterReview", "appearanceConfirmation", "matterConfirmation", "placeholderConfirmation"]))) throw new Error("Recovery template-review identity is invalid.");
  const matter = value.reviewState.matterReview;
  if (!plainObject(matter) || !hasOnlyKeys(matter, new Set(["sourceName", "pdfSha256", "exactByteLength", "pageCount", "extractedCharacterCount", "textReliability", "requiresVisualConfirmation", "notice", "matterNumbers", "partyNames", "forums", "matterTitles", "placeholders"])) || matter.pdfSha256 !== pdfSha256 || !Number.isSafeInteger(matter.exactByteLength) || matter.exactByteLength < 1 || matter.exactByteLength > 25 * 1024 * 1024 || !Number.isSafeInteger(matter.pageCount) || matter.pageCount < 1 || matter.pageCount > 25 || !Number.isSafeInteger(matter.extractedCharacterCount) || matter.extractedCharacterCount < 0 || matter.extractedCharacterCount > 250_000 || !["reliable", "limited", "none"].includes(matter.textReliability) || typeof matter.requiresVisualConfirmation !== "boolean") throw new Error("Recovery template matter review is invalid.");
  const matterReview = { sourceName: boundedString(matter.sourceName, "Recovery template source name", 512), pdfSha256, exactByteLength: matter.exactByteLength, pageCount: matter.pageCount, extractedCharacterCount: matter.extractedCharacterCount, textReliability: matter.textReliability, requiresVisualConfirmation: matter.requiresVisualConfirmation, notice: boundedString(matter.notice, "Recovery template review notice", 2_048) };
  for (const [property, kind] of [["matterNumbers", "matter-number"], ["partyNames", "party-name"], ["forums", "forum"], ["matterTitles", "matter-title"], ["placeholders", "placeholder"]]) {
    if (!Array.isArray(matter[property]) || matter[property].length > 100) throw new Error("Recovery template findings are invalid.");
    matterReview[property] = matter[property].map((finding) => validateTemplateFinding(finding, kind, matter.pageCount));
  }
  const reviewState = { matterReview };
  if (value.reviewState.appearanceConfirmation !== undefined) reviewState.appearanceConfirmation = validateTemplateConfirmation(value.reviewState.appearanceConfirmation, pdfSha256);
  if (value.reviewState.matterConfirmation !== undefined) reviewState.matterConfirmation = validateMatterConfirmation(value.reviewState.matterConfirmation, pdfSha256);
  if (value.reviewState.placeholderConfirmation !== undefined) reviewState.placeholderConfirmation = validateTemplateConfirmation(value.reviewState.placeholderConfirmation, pdfSha256);
  return { slot, sourceId, ...(renderedSourceId ? { renderedSourceId } : {}), sourceFormat, sourceSha256: sourceSha256.toLowerCase(), pdfSha256: pdfSha256.toLowerCase(), reviewState };
}

function validateArrangement(value) {
  if (!plainObject(value) || value.version !== 1 || !Array.isArray(value.nodes) || !hasOnlyKeys(value, new Set(["version", "nodes"]))) {
    throw new Error("Recovery arrangement is invalid.");
  }
  if (value.nodes.length > MAX_ORDER) throw new Error("Recovery arrangement contains too many nodes.");
  const exhibitIds = new Set();
  const sectionIds = new Set();
  let totalNodes = value.nodes.length;
  const validateExhibit = (node) => {
    if (!plainObject(node) || node.type !== "exhibit" || !hasOnlyKeys(node, new Set(["type", "exhibitId"]))) throw new Error("Recovery arrangement exhibit is invalid.");
    const exhibitId = boundedString(node.exhibitId, "Recovery arrangement exhibit ID", 1_024);
    if (!exhibitId || exhibitId.includes("\0") || exhibitIds.has(exhibitId)) throw new Error("Recovery arrangement contains an invalid or duplicate exhibit ID.");
    exhibitIds.add(exhibitId);
    return { type: "exhibit", exhibitId };
  };
  const nodes = value.nodes.map((node) => {
    if (!plainObject(node)) throw new Error("Recovery arrangement node is invalid.");
    if (node.type === "exhibit") return validateExhibit(node);
    if (node.type !== "section" || !hasOnlyKeys(node, new Set(["type", "id", "heading", "exhibits"])) || !Array.isArray(node.exhibits)) throw new Error("Recovery arrangement section is invalid.");
    const id = boundedString(node.id, "Recovery arrangement section ID", 1_024);
    const heading = boundedString(node.heading, "Recovery arrangement section heading", 512);
    if (!id || id.includes("\0") || sectionIds.has(id) || !heading.trim() || heading.includes("\0")) throw new Error("Recovery arrangement contains an invalid or duplicate section.");
    sectionIds.add(id);
    totalNodes += node.exhibits.length;
    if (totalNodes > MAX_ORDER) throw new Error("Recovery arrangement contains too many nodes.");
    return { type: "section", id, heading, exhibits: node.exhibits.map(validateExhibit) };
  });
  return { version: 1, nodes };
}

function validateRecoveryPayload(payload) {
  if (!plainObject(payload)) throw new Error("Recovery payload must be an object.");
  for (const key of Object.keys(payload)) if (!TOP_LEVEL_KEYS.has(key)) throw new Error(`Recovery payload contains unsupported field ${key}.`);
  const encoded = JSON.stringify(payload);
  if (Buffer.byteLength(encoded, "utf8") > MAX_JSON_BYTES) throw new Error("Recovery payload exceeds the journal safety limit.");
  const sources = Array.isArray(payload.sources) ? payload.sources.map(validateSource) : [];
  if (!sources.length || sources.length > MAX_SOURCES) throw new Error("Recovery payload has an invalid number of sources.");
  const sourceIds = new Set();
  for (const source of sources) {
    if (sourceIds.has(source.id)) throw new Error("Recovery payload contains duplicate source IDs.");
    sourceIds.add(source.id);
  }
  const candidates = Array.isArray(payload.candidates) ? payload.candidates : [];
  if (payload.finalOrder !== undefined && !Array.isArray(payload.finalOrder)) throw new Error("Recovery legacy final order is invalid.");
  const finalOrder = Array.isArray(payload.finalOrder) ? payload.finalOrder : [];
  if (payload.arrangement !== undefined && payload.finalOrder !== undefined) throw new Error("Recovery payload cannot contain both arrangement and legacy final order.");
  const arrangement = payload.arrangement === undefined ? undefined : validateArrangement(payload.arrangement);
  const resolutions = Array.isArray(payload.resolutions) ? payload.resolutions : [];
  if (candidates.length > MAX_CANDIDATES || finalOrder.length > MAX_ORDER || resolutions.length > MAX_RESOLUTIONS) throw new Error("Recovery payload contains too many review records.");
  const orderIds = new Set();
  for (const id of finalOrder) {
    boundedString(id, "Recovery order ID", 1_024);
    if (!id || id.includes("\0") || orderIds.has(id)) throw new Error("Recovery payload contains invalid or duplicate order IDs.");
    orderIds.add(id);
  }
  const rawPageSizeChoices = payload.pageSizeChoices ?? {};
  if (!plainObject(rawPageSizeChoices) || Object.keys(rawPageSizeChoices).length > MAX_SOURCES) throw new Error("Recovery page-size choices are invalid.");
  const pageSizeChoices = {};
  for (const [id, choice] of Object.entries(rawPageSizeChoices)) {
    boundedString(id, "Recovery page-size source ID", 512);
    if (choice !== "convert-to-a4" && choice !== "keep-original") throw new Error("Recovery page-size choice is invalid.");
    pageSizeChoices[id] = choice;
  }
  const templateReviews = payload.templateReviews === undefined ? undefined : payload.templateReviews;
  if (templateReviews !== undefined && (!Array.isArray(templateReviews) || templateReviews.length > 4)) throw new Error("Recovery template-review metadata is invalid.");
  const validatedTemplateReviews = templateReviews?.map(validateTemplateReview);
  let templateDiscrepancyConfirmation;
  if (payload.templateDiscrepancyConfirmation !== undefined) {
    const confirmation = payload.templateDiscrepancyConfirmation;
    if (!plainObject(confirmation) || !hasOnlyKeys(confirmation, new Set(["fingerprint", "confirmedAt"]))) throw new Error("Recovery template-discrepancy confirmation is invalid.");
    templateDiscrepancyConfirmation = { fingerprint: boundedString(confirmation.fingerprint, "Recovery template-discrepancy fingerprint", 100_000), confirmedAt: confirmation.confirmedAt };
    if (!validDate(templateDiscrepancyConfirmation.confirmedAt)) throw new Error("Recovery template-discrepancy confirmation date is invalid.");
  }
  const validated = { ...payload, sources, candidates, resolutions, pageSizeChoices, ...(validatedTemplateReviews ? { templateReviews: validatedTemplateReviews } : {}), ...(templateDiscrepancyConfirmation ? { templateDiscrepancyConfirmation } : {}) };
  if (arrangement) {
    validated.arrangement = arrangement;
    delete validated.finalOrder;
  } else {
    validated.finalOrder = finalOrder;
  }
  return validated;
}

function validateRecoveryId(value) {
  if (typeof value !== "string" || !/^[a-f0-9-]{36}$/i.test(value)) throw new Error("Recovery ID is invalid.");
}

async function hashFile(filePath) {
  const bytes = await readFile(filePath);
  return { bytes, sha256: createHash("sha256").update(bytes).digest("hex") };
}

class RecoveryStore {
  constructor(root) {
    if (!path.isAbsolute(root)) throw new Error("Recovery root must be absolute.");
    this.root = root;
    this.journalPath = path.join(root, JOURNAL_NAME);
    this.backupPath = path.join(root, BACKUP_NAME);
    this.writeQueue = Promise.resolve();
  }

  async readJournal() {
    let bytes;
    try {
      bytes = await readFile(this.journalPath);
    } catch (error) {
      if (error?.code !== "ENOENT") throw error;
      try { bytes = await readFile(this.backupPath); } catch (backupError) { if (backupError?.code === "ENOENT") return null; throw backupError; }
    }
    if (bytes.byteLength > MAX_JSON_BYTES) throw new Error("Recovery journal exceeds the safety limit.");
    const stored = JSON.parse(bytes.toString("utf8"));
    validateRecoveryId(stored.recoveryId);
    if (!Number.isSafeInteger(stored.revision) || stored.revision < 0 || typeof stored.dirty !== "boolean") throw new Error("Recovery journal metadata is invalid.");
    stored.payload = validateRecoveryPayload(stored.payload);
    return stored;
  }

  async status() {
    try {
      const journal = await this.readJournal();
      return journal && journal.dirty ? { available: true, stored: true, recoveryId: journal.recoveryId, revision: journal.revision, projectName: journal.payload.project?.name ?? "Recovered exhibit project" } : { available: false, stored: Boolean(journal) };
    } catch (error) {
      let stored = false;
      for (const candidate of [this.journalPath, this.backupPath]) {
        try { await stat(candidate); stored = true; break; } catch {}
      }
      return { available: false, stored, corrupt: stored, issue: stored ? "The automatic recovery journal is damaged and cannot be restored. You can delete it from Local recovery data." : undefined };
    }
  }

  async begin() {
    return { recoveryId: randomUUID(), revision: 0 };
  }

  async replaceJournal(temporaryPath) {
    await rm(this.backupPath, { force: true });
    try { await rename(this.journalPath, this.backupPath); } catch (error) { if (error?.code !== "ENOENT") throw error; }
    try {
      await rename(temporaryPath, this.journalPath);
      await rm(this.backupPath, { force: true });
    } catch (error) {
      try { await rename(this.backupPath, this.journalPath); } catch {}
      await rm(temporaryPath, { force: true });
      throw error;
    }
  }

  write(recoveryId, revision, payload) {
    const pending = this.writeQueue.then(() => this.writeNow(recoveryId, revision, payload));
    this.writeQueue = pending.catch(() => {});
    return pending;
  }

  async writeNow(recoveryId, revision, payload) {
    validateRecoveryId(recoveryId);
    if (!Number.isSafeInteger(revision) || revision < 1) throw new Error("Recovery revision is invalid.");
    const validPayload = validateRecoveryPayload(payload);
    const current = await this.readJournal();
    if (current?.recoveryId === recoveryId && revision <= current.revision) throw new Error("Recovery revision is stale.");
    const stored = { schemaVersion: 1, recoveryId, revision, dirty: true, updatedAt: new Date().toISOString(), payload: validPayload, savedArchive: current?.recoveryId === recoveryId ? current.savedArchive ?? null : null };
    await mkdir(this.root, { recursive: true });
    const temporaryPath = path.join(this.root, `${JOURNAL_NAME}.${process.pid}.${randomUUID()}.tmp`);
    const handle = await open(temporaryPath, "wx");
    try {
      await handle.writeFile(JSON.stringify(stored));
      await handle.sync();
    } finally {
      await handle.close();
    }
    await this.replaceJournal(temporaryPath);
    return { saved: true, recoveryId, revision };
  }

  async load(recoveryId) {
    validateRecoveryId(recoveryId);
    const journal = await this.readJournal();
    if (!journal || journal.recoveryId !== recoveryId) throw new Error("Recovery journal is unavailable.");
    return journal;
  }

  async markClean(recoveryId, revision, savedArchive = null) {
    validateRecoveryId(recoveryId);
    const journal = await this.readJournal();
    if (!journal || journal.recoveryId !== recoveryId) return { cleaned: false };
    if (!Number.isSafeInteger(revision) || revision < journal.revision) return { cleaned: false };
    if (savedArchive !== null) {
      if (!plainObject(savedArchive) || !path.isAbsolute(savedArchive.path) || !HASH.test(savedArchive.sha256)) throw new Error("Saved project archive descriptor is invalid.");
      savedArchive = { path: savedArchive.path, sha256: savedArchive.sha256.toLowerCase() };
    }
    await this.write(recoveryId, revision + 1, journal.payload);
    const updated = await this.readJournal();
    updated.dirty = false;
    updated.revision = revision + 1;
    updated.savedArchive = savedArchive ?? updated.savedArchive ?? null;
    const temporaryPath = path.join(this.root, `${JOURNAL_NAME}.${process.pid}.${randomUUID()}.tmp`);
    const handle = await open(temporaryPath, "wx");
    try { await handle.writeFile(JSON.stringify(updated)); await handle.sync(); } finally { await handle.close(); }
    await this.replaceJournal(temporaryPath);
    return { cleaned: true, revision: updated.revision };
  }

  async discard(recoveryId) {
    validateRecoveryId(recoveryId);
    const journal = await this.readJournal();
    if (!journal || journal.recoveryId !== recoveryId) return { discarded: false };
    await rm(this.journalPath, { force: true });
    await rm(this.backupPath, { force: true });
    return { discarded: true };
  }

  clearAll() {
    const pending = this.writeQueue.then(async () => {
      await rm(this.root, { recursive: true, force: true, maxRetries: 2, retryDelay: 100 });
      return { cleared: true };
    });
    this.writeQueue = pending.catch(() => {});
    return pending;
  }

  async readSource(recoveryId, sourceId) {
    validateRecoveryId(recoveryId);
    boundedString(sourceId, "Recovery source ID", 512);
    const journal = await this.load(recoveryId);
    const source = journal.payload.sources.find((item) => item.id === sourceId);
    if (!source) throw new Error("Recovery source is not part of this journal.");
    const details = await stat(source.path);
    if (!details.isFile() || details.size !== source.size) throw new Error(`${source.name} is missing or has changed since the recovery journal was written.`);
    const read = await hashFile(source.path);
    if (read.sha256 !== source.sha256) throw new Error(`${source.name} has changed since the recovery journal was written.`);
    return { id: source.id, role: source.role, name: source.name, sha256: source.sha256, bytes: new Uint8Array(read.bytes) };
  }
}

module.exports = { RecoveryStore, validateRecoveryPayload };
