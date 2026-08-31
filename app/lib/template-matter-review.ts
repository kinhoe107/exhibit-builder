export type TemplateMatterFindingKind =
  | "matter-number"
  | "party-name"
  | "forum"
  | "matter-title"
  | "placeholder";

export type TemplateMatterGeometry = {
  pageNumber: number;
  x: number;
  y: number;
  width: number;
  height: number;
  fontSize: number;
  color?: { r: number; g: number; b: number };
};

export type TemplateMatterFinding = {
  id: string;
  kind: TemplateMatterFindingKind;
  value: string;
  normalizedValue: string;
  pageNumbers: number[];
  evidence: string[];
  geometry?: TemplateMatterGeometry;
  /** Extraction is only a review aid. It is never independent legal verification. */
  unverified: true;
};

export type TemplateMatterReview = {
  sourceName: string;
  pdfSha256: string;
  exactByteLength: number;
  pageCount: number;
  extractedCharacterCount: number;
  textReliability: "reliable" | "limited" | "none";
  requiresVisualConfirmation: boolean;
  notice: string;
  matterNumbers: TemplateMatterFinding[];
  partyNames: TemplateMatterFinding[];
  forums: TemplateMatterFinding[];
  matterTitles: TemplateMatterFinding[];
  placeholders: TemplateMatterFinding[];
};

export type TemplateMatterValues = {
  matterNumbers: string[];
  partyNames: string[];
  forums: string[];
  matterTitles: string[];
};

export type TemplateMatterPatch = {
  findingId: string;
  value: string;
};

export type MatterOccurrenceDraft = {
  findingId: string;
  kind: Exclude<TemplateMatterFindingKind, "placeholder">;
  originalValue: string;
  value: string;
};

export type MatterDraft = {
  occurrences: MatterOccurrenceDraft[];
};

export type MatterListDraft = {
  matterNumbers: string;
  partyNames: string;
  forums: string;
  matterTitles: string;
};

export type TemplateReviewReference = {
  templateId: string;
  role: string;
  sourceName: string;
  review: TemplateMatterReview;
  /** Reviewer-corrected list used for comparison when confirmation is bound to this PDF. */
  confirmedValues?: TemplateMatterValues;
};

export type TemplateMatterDiscrepancy = {
  field: "matter-number" | "party-name" | "forum" | "matter-title";
  message: string;
  evidence: Array<{
    templateId: string;
    role: string;
    sourceName: string;
    pdfSha256: string;
    values: string[];
    normalizedValues: string[];
  }>;
  /** A discrepancy is evidence for the user to review, not a legal conclusion. */
  unverified: true;
};

export const EMPTY_MATTER_VALUES: TemplateMatterValues = {
  matterNumbers: [],
  partyNames: [],
  forums: [],
  matterTitles: [],
};

const REVIEW_LIMITS = {
  pdfBytes: 25 * 1024 * 1024,
  pages: 25,
  extractedCharacters: 250_000,
  openMilliseconds: 30_000,
  pageMilliseconds: 10_000,
  totalMilliseconds: 2 * 60 * 1000,
} as const;

type PdfInput = Blob | ArrayBuffer | Uint8Array;

type TextItemGeom = {
  str: string;
  x: number;
  y: number;
  width: number;
  height: number;
  fontSize: number;
};

type TextLine = {
  pageNumber: number;
  text: string;
  items: TextItemGeom[];
  geometry: TemplateMatterGeometry;
};

function normalizeWhitespace(value: string) {
  return value.replace(/\s+/g, " ").trim();
}

function normalizeMatterValue(value: string) {
  return normalizeWhitespace(value)
    .normalize("NFKC")
    .replace(/[\u2010-\u2015]/g, "-")
    .replace(/\s*([/,:&()-])\s*/g, "$1")
    .replace(/\b(?:THE|LIMITED|LTD\.?|PLC|LLP|INC\.?|CORPORATION|CORP\.?)\b/gi, (word) => word.toUpperCase())
    .toLocaleUpperCase("en-GB");
}

function cleanFindingValue(value: string) {
  return normalizeWhitespace(value)
    .replace(/^[\s:;,.-]+/, "")
    .replace(/[\s;,.-]+$/, "");
}

function toBytes(input: PdfInput) {
  if (input instanceof Uint8Array) return Promise.resolve(new Uint8Array(input));
  if (input instanceof ArrayBuffer) return Promise.resolve(new Uint8Array(input.slice(0)));
  return input.arrayBuffer().then((buffer) => new Uint8Array(buffer));
}

async function sha256(bytes: Uint8Array) {
  const digest = await crypto.subtle.digest("SHA-256", Uint8Array.from(bytes).buffer);
  return Array.from(new Uint8Array(digest), (part) => part.toString(16).padStart(2, "0")).join("");
}

function withTimeout<T>(operation: Promise<T>, milliseconds: number, message: string, onTimeout?: () => void) {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return Promise.race([
    operation,
    new Promise<never>((_resolve, reject) => {
      timer = setTimeout(() => {
        onTimeout?.();
        reject(new Error(message));
      }, milliseconds);
    }),
  ]).finally(() => {
    if (timer) clearTimeout(timer);
  });
}

function itemFontSize(transform: number[], height: number) {
  const fromTransform = Math.abs(typeof transform[3] === "number" ? transform[3] : 0) || Math.hypot(Number(transform[0]) || 0, Number(transform[1]) || 0);
  return fromTransform || (Number.isFinite(height) && height > 0 ? height : 12);
}

function geometryFromItems(pageNumber: number, items: TextItemGeom[]): TemplateMatterGeometry {
  const left = Math.min(...items.map((item) => item.x));
  const baseline = Math.min(...items.map((item) => item.y));
  const right = Math.max(...items.map((item) => item.x + item.width));
  const fontSize = Math.max(...items.map((item) => item.fontSize));
  const height = Math.max(...items.map((item) => item.height || item.fontSize), fontSize);
  return {
    pageNumber,
    x: left,
    y: baseline,
    width: Math.max(1, right - left),
    height: Math.max(1, height),
    fontSize,
  };
}

function geometryForValue(line: TextLine, rawValue: string): TemplateMatterGeometry {
  const needle = normalizeWhitespace(rawValue);
  if (!needle || !line.items.length) return line.geometry;
  const joined = line.items.map((item) => item.str).join("");
  const start = joined.toLocaleUpperCase("en-GB").indexOf(needle.replace(/\s+/g, "").toLocaleUpperCase("en-GB"));
  if (start < 0) return line.geometry;
  let cursor = 0;
  const matched: TextItemGeom[] = [];
  for (const item of line.items) {
    const next = cursor + item.str.length;
    if (next > start && cursor < start + needle.replace(/\s+/g, "").length) matched.push(item);
    cursor = next;
  }
  return matched.length ? geometryFromItems(line.pageNumber, matched) : line.geometry;
}

export function matterFindingId(kind: TemplateMatterFindingKind, geometry: TemplateMatterGeometry | undefined, normalizedValue: string) {
  const place = geometry
    ? `${geometry.pageNumber}:${geometry.x.toFixed(2)}:${geometry.y.toFixed(2)}`
    : "none";
  return `${kind}:${place}:${normalizedValue}`.slice(0, 240);
}

function pageTextLines(items: Array<unknown>, pageNumber: number) {
  const lines: TextLine[] = [];
  let currentItems: TextItemGeom[] = [];
  let previousY: number | null = null;
  const flush = () => {
    const text = normalizeWhitespace(currentItems.map((item) => item.str).join(" "));
    if (text && currentItems.length) {
      lines.push({
        pageNumber,
        text,
        items: currentItems,
        geometry: geometryFromItems(pageNumber, currentItems),
      });
    }
    currentItems = [];
  };

  for (const item of items) {
    if (!item || typeof item !== "object" || !("str" in item) || typeof item.str !== "string") continue;
    const transform = "transform" in item && Array.isArray(item.transform)
      ? item.transform.map((part) => typeof part === "number" ? part : 0)
      : [];
    const y: number | null = typeof transform[5] === "number" ? transform[5] : previousY;
    const x = typeof transform[4] === "number" ? transform[4] : 0;
    const width = "width" in item && typeof item.width === "number" ? item.width : 0;
    const height = "height" in item && typeof item.height === "number" ? item.height : 0;
    if (previousY !== null && y !== null && Math.abs(y - previousY) > 1.5) flush();
    currentItems.push({
      str: item.str,
      x,
      y: y ?? 0,
      width,
      height,
      fontSize: itemFontSize(transform, height),
    });
    if ("hasEOL" in item && item.hasEOL === true) flush();
    previousY = y;
  }
  flush();
  return lines;
}

function addFinding(
  findings: Map<string, TemplateMatterFinding>,
  kind: TemplateMatterFindingKind,
  rawValue: string,
  line: TextLine,
) {
  const value = cleanFindingValue(rawValue);
  if (!value || value.length > 240) return;
  const normalizedValue = normalizeMatterValue(value);
  if (!normalizedValue) return;
  const geometry = geometryForValue(line, value);
  const id = matterFindingId(kind, geometry, normalizedValue);
  const existing = findings.get(id);
  if (existing) {
    if (!existing.pageNumbers.includes(line.pageNumber)) existing.pageNumbers.push(line.pageNumber);
    if (!existing.evidence.includes(line.text)) existing.evidence.push(line.text);
    return;
  }
  findings.set(id, {
    id,
    kind,
    value,
    normalizedValue,
    pageNumbers: [line.pageNumber],
    evidence: [line.text],
    geometry,
    unverified: true,
  });
}

function isFieldLabelLine(text: string) {
  return /^(?:icc\s+)?(?:case|matter|claim|arbitration|reference|document|bundle|place of arbitration|hearing)\b/i.test(text)
    || /\b(?:case|matter|claim)\s*(?:number|no\.?|#)\b/i.test(text)
    || /editable|synthetic test|replace case details/i.test(text);
}

function isConjunctionLine(text: string) {
  return /^(?:and|v\.?|vs\.?|versus)$/i.test(text);
}

function extractFindings(lines: TextLine[]) {
  const findings = new Map<string, TemplateMatterFinding>();
  const roleLine = /^(.{2,180}?)\s*\(?\b(claimant|respondent|applicant|defendant|plaintiff|petitioner|appellant)\b\)?\s*$/i;
  const labelledNumber = /\b(?:case|matter|claim|arbitration|reference|ref\.?|proceedings?)\s*(?:number|no\.?|#)\s*[:.-]?\s*([A-Z0-9][A-Z0-9 ./_-]{2,80})/i;
  const forum = /\b(?:IN THE\s+.{0,100}?(?:COURT|TRIBUNAL)|(?:HIGH|SUPREME|CIRCUIT|DISTRICT|COMMERCIAL|CROWN|COUNTY)\s+COURT|COURT OF APPEAL|ARBITRAL TRIBUNAL|INTERNATIONAL (?:COURT|CHAMBER|CENTRE)|LONDON COURT OF INTERNATIONAL ARBITRATION|LCIA\s+ARBITRATION|ICC\s+(?:COURT|ARBITRATION)|SINGAPORE INTERNATIONAL ARBITRATION CENTRE|SINGAPORE INTERNATION ARBITRATION CENTRE)\b/i;
  const title = /^(.{2,110}?\s+(?:v\.?|vs\.?|versus)\s+.{2,110})$/i;
  const matterOf = /^(IN THE MATTER OF\s+.{3,200})$/i;
  const placeholderPatterns = [
    /\[[^\]\r\n]{2,100}\]/g,
    /\{\{[^}\r\n]{2,100}\}\}/g,
    /<<[^>\r\n]{2,100}>>/g,
    /\b(?:INSERT|ENTER|ADD)\s+(?:CASE|MATTER|CLAIM|PARTY|PARTIES|TRIBUNAL|COURT|TITLE|NUMBER|NO\.)[^\r\n]{0,80}/gi,
    /\b(?:CASE|MATTER|CLAIM)\s*(?:NUMBER|NO\.?)\s*[:.-]?\s*_{3,}/gi,
  ];

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const numberMatch = line.text.match(labelledNumber);
    if (numberMatch?.[1]) addFinding(findings, "matter-number", numberMatch[1], line);

    const roleMatch = line.text.match(roleLine);
    if (roleMatch?.[1]) addFinding(findings, "party-name", roleMatch[1], line);

    const titleMatch = line.text.match(title);
    if (titleMatch?.[1]) {
      addFinding(findings, "matter-title", titleMatch[1], line);
      const parties = titleMatch[1].split(/\s+(?:v\.?|vs\.?|versus)\s+/i);
      for (const party of parties) addFinding(findings, "party-name", party, line);
    }

    const matterMatch = line.text.match(matterOf);
    if (matterMatch?.[1]) addFinding(findings, "matter-title", matterMatch[1], line);

    const forumMatch = line.text.match(forum);
    if (forumMatch?.[0]) addFinding(findings, "forum", /^\s*IN THE\b/i.test(line.text) ? line.text : forumMatch[0], line);

    if (/\bBETWEEN\s*:?\s*$/i.test(line.text) || /\bARBITRATION BETWEEN\b/i.test(line.text)) {
      for (let offset = 1; offset <= 8 && index + offset < lines.length; offset += 1) {
        const candidate = lines[index + offset];
        if (isConjunctionLine(candidate.text)) continue;
        if (isFieldLabelLine(candidate.text)) break;
        const candidateRole = candidate.text.match(roleLine);
        if (candidateRole?.[1]) {
          addFinding(findings, "party-name", candidateRole[1], candidate);
          continue;
        }
        if (candidate.text.length >= 3 && candidate.text.length <= 180) {
          addFinding(findings, "party-name", candidate.text, candidate);
        }
      }
    }

    for (const pattern of placeholderPatterns) {
      pattern.lastIndex = 0;
      for (const match of line.text.matchAll(pattern)) addFinding(findings, "placeholder", match[0], line);
    }
  }

  const all = [...findings.values()].map((finding) => ({
    ...finding,
    pageNumbers: [...finding.pageNumbers].sort((left, right) => left - right),
    evidence: [...finding.evidence].sort((left, right) => left.localeCompare(right, "en-GB")),
  }));
  all.sort((left, right) => left.kind.localeCompare(right.kind) || left.normalizedValue.localeCompare(right.normalizedValue, "en-GB") || (left.geometry?.y ?? 0) - (right.geometry?.y ?? 0));
  return all;
}

function findingsOfKind(findings: TemplateMatterFinding[], kind: TemplateMatterFindingKind) {
  return findings.filter((finding) => finding.kind === kind);
}

export function writableMatterFindings(review: TemplateMatterReview): Array<TemplateMatterFinding & { kind: Exclude<TemplateMatterFindingKind, "placeholder"> }> {
  return [...review.matterNumbers, ...review.partyNames, ...review.forums, ...review.matterTitles].map((finding) => ({
    ...finding,
    kind: finding.kind as Exclude<TemplateMatterFindingKind, "placeholder">,
  }));
}

/**
 * Reads possible identifying details from the exact PDF artifact supplied.
 * Results are deliberately labelled unverified and never replace visual review.
 */
export async function reviewTemplateMatterPdf(
  input: PdfInput,
  sourceName = "Template.pdf",
): Promise<TemplateMatterReview> {
  const bytes = await toBytes(input);
  if (!bytes.byteLength || bytes.byteLength > REVIEW_LIMITS.pdfBytes) {
    throw new Error("The template PDF is empty or exceeds the 25 MB review safety limit.");
  }
  if (String.fromCharCode(...bytes.subarray(0, 5)) !== "%PDF-") {
    throw new Error("The rendered template artifact is not a readable PDF.");
  }
  const [pdfSha256, pdfjs] = await Promise.all([
    sha256(bytes),
    import("pdfjs-dist/legacy/build/pdf.mjs"),
  ]);
  if (!pdfjs.GlobalWorkerOptions.workerSrc) {
    pdfjs.GlobalWorkerOptions.workerSrc = new URL(
      "../../node_modules/pdfjs-dist/legacy/build/pdf.worker.min.mjs",
      import.meta.url,
    ).toString();
  }
  const startedAt = Date.now();
  const task = pdfjs.getDocument({
    data: new Uint8Array(bytes),
    isEvalSupported: false,
    useWorkerFetch: false,
    verbosity: 0,
  });
  const document = await withTimeout(
    task.promise,
    REVIEW_LIMITS.openMilliseconds,
    "The template PDF did not open within the 30-second review safety limit.",
    () => { void task.destroy().catch(() => undefined); },
  );
  try {
    if (document.numPages > REVIEW_LIMITS.pages) {
      throw new Error(`The template PDF exceeds the ${REVIEW_LIMITS.pages}-page review safety limit.`);
    }
    const lines: TextLine[] = [];
    let extractedCharacterCount = 0;
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
      const remaining = REVIEW_LIMITS.totalMilliseconds - (Date.now() - startedAt);
      if (remaining <= 0) throw new Error("Template matter review exceeded the two-minute safety limit.");
      const page = await withTimeout(
        document.getPage(pageNumber),
        Math.min(REVIEW_LIMITS.pageMilliseconds, remaining),
        `Template page ${pageNumber} loading exceeded the review safety limit.`,
      );
      try {
        const content = await withTimeout(
          page.getTextContent(),
          Math.min(REVIEW_LIMITS.pageMilliseconds, remaining),
          `Template page ${pageNumber} text extraction exceeded the review safety limit.`,
        );
        const pageLines = pageTextLines(content.items, pageNumber);
        extractedCharacterCount += pageLines.reduce((total, line) => total + line.text.length, 0);
        if (extractedCharacterCount > REVIEW_LIMITS.extractedCharacters) {
          throw new Error("The template PDF exceeds the extracted-text review safety limit.");
        }
        lines.push(...pageLines);
      } finally {
        page.cleanup();
      }
    }
    const findings = extractFindings(lines);
    const nonWhitespaceText = lines.map((line) => line.text).join(" ").replace(/\s/g, "");
    const alphanumeric = (nonWhitespaceText.match(/[\p{L}\p{N}]/gu) ?? []).length;
    const textReliability = alphanumeric >= 24
      ? "reliable" as const
      : alphanumeric > 0
        ? "limited" as const
        : "none" as const;
    const requiresVisualConfirmation = textReliability !== "reliable";
    const notice = requiresVisualConfirmation
      ? "Matter details could not be read reliably from this template. Visually check the exact PDF preview and confirm the matter number, party names and other identifying details."
      : "Possible matter details were extracted from the exact PDF artifact for review. They are unverified and must be checked against the visible template.";
    return {
      sourceName,
      pdfSha256,
      exactByteLength: bytes.byteLength,
      pageCount: document.numPages,
      extractedCharacterCount,
      textReliability,
      requiresVisualConfirmation,
      notice,
      matterNumbers: findingsOfKind(findings, "matter-number"),
      partyNames: findingsOfKind(findings, "party-name"),
      forums: findingsOfKind(findings, "forum"),
      matterTitles: findingsOfKind(findings, "matter-title"),
      placeholders: findingsOfKind(findings, "placeholder"),
    };
  } finally {
    await document.destroy().catch(() => undefined);
  }
}

const DISCREPANCY_FIELDS = [
  ["matter-number", "matterNumbers", "matter or case numbers"],
  ["party-name", "partyNames", "party names"],
  ["forum", "forums", "forum or tribunal details"],
  ["matter-title", "matterTitles", "matter titles"],
] as const;

const MATTER_VALUE_LIMITS = { items: 20, characters: 240 } as const;

export function parseMatterValueList(value: string): string[] {
  const seen = new Set<string>();
  const items: string[] = [];
  for (const part of value.split(/[\n;]+/)) {
    const cleaned = cleanFindingValue(part);
    if (!cleaned || cleaned.length > MATTER_VALUE_LIMITS.characters) continue;
    const key = normalizeMatterValue(cleaned);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    items.push(cleaned);
    if (items.length >= MATTER_VALUE_LIMITS.items) break;
  }
  return items;
}

export function extractedMatterValues(review: TemplateMatterReview): TemplateMatterValues {
  return {
    matterNumbers: review.matterNumbers.map((finding) => finding.value),
    partyNames: review.partyNames.map((finding) => finding.value),
    forums: review.forums.map((finding) => finding.value),
    matterTitles: review.matterTitles.map((finding) => finding.value),
  };
}

export function matterValuesFromConfirmation(confirmation?: Partial<TemplateMatterValues> & { patches?: TemplateMatterPatch[] }): TemplateMatterValues | undefined {
  if (!confirmation) return undefined;
  if (confirmation.matterNumbers === undefined && confirmation.partyNames === undefined && confirmation.forums === undefined && confirmation.matterTitles === undefined) return undefined;
  return {
    matterNumbers: [...(confirmation.matterNumbers ?? [])],
    partyNames: [...(confirmation.partyNames ?? [])],
    forums: [...(confirmation.forums ?? [])],
    matterTitles: [...(confirmation.matterTitles ?? [])],
  };
}

export function effectiveMatterValues(review: TemplateMatterReview, confirmation?: Partial<TemplateMatterValues> & { pdfSha256?: string; patches?: TemplateMatterPatch[] }, pdfSha256 = review.pdfSha256): TemplateMatterValues {
  const confirmed = confirmation?.pdfSha256 === pdfSha256 ? matterValuesFromConfirmation(confirmation) : undefined;
  return confirmed ?? extractedMatterValues(review);
}

export function matterListDraftFromValues(values: TemplateMatterValues): MatterListDraft {
  return {
    matterNumbers: values.matterNumbers.join("\n"),
    partyNames: values.partyNames.join("\n"),
    forums: values.forums.join("\n"),
    matterTitles: values.matterTitles.join("\n"),
  };
}

export function parseMatterListDraft(draft: MatterListDraft): TemplateMatterValues {
  return {
    matterNumbers: parseMatterValueList(draft.matterNumbers),
    partyNames: parseMatterValueList(draft.partyNames),
    forums: parseMatterValueList(draft.forums),
    matterTitles: parseMatterValueList(draft.matterTitles),
  };
}

export function matterDraftFromReview(review: TemplateMatterReview, confirmation?: Partial<TemplateMatterValues> & { patches?: TemplateMatterPatch[]; pdfSha256?: string }): MatterDraft {
  const boundPatches = confirmation?.pdfSha256 === review.pdfSha256 ? confirmation?.patches : undefined;
  const patchById = new Map((boundPatches ?? []).map((patch) => [patch.findingId, patch.value]));
  const occurrences: MatterOccurrenceDraft[] = writableMatterFindings(review).map((finding) => {
    const findingId = finding.id || matterFindingId(finding.kind, finding.geometry, finding.normalizedValue);
    return {
      findingId,
      kind: finding.kind,
      originalValue: finding.value,
      value: patchById.has(findingId) ? patchById.get(findingId) ?? "" : finding.value,
    };
  });
  return { occurrences };
}

export function parseMatterDraft(draft: MatterDraft): { values: TemplateMatterValues; patches: TemplateMatterPatch[] } {
  const values: TemplateMatterValues = { ...EMPTY_MATTER_VALUES, matterNumbers: [], partyNames: [], forums: [], matterTitles: [] };
  const patches: TemplateMatterPatch[] = [];
  for (const occurrence of draft.occurrences) {
    const cleaned = cleanFindingValue(occurrence.value);
    patches.push({ findingId: occurrence.findingId, value: cleaned });
    if (!cleaned) continue;
    if (occurrence.kind === "matter-number") values.matterNumbers.push(cleaned);
    if (occurrence.kind === "party-name") values.partyNames.push(cleaned);
    if (occurrence.kind === "forum") values.forums.push(cleaned);
    if (occurrence.kind === "matter-title") values.matterTitles.push(cleaned);
  }
  return { values, patches };
}

export function matterValuesEqual(left: TemplateMatterValues, right: TemplateMatterValues): boolean {
  return (["matterNumbers", "partyNames", "forums", "matterTitles"] as const).every((property) => {
    const leftNorm = [...new Set(left[property].map(normalizeMatterValue).filter(Boolean))].sort((a, b) => a.localeCompare(b, "en-GB"));
    const rightNorm = [...new Set(right[property].map(normalizeMatterValue).filter(Boolean))].sort((a, b) => a.localeCompare(b, "en-GB"));
    return leftNorm.join("\u001f") === rightNorm.join("\u001f");
  });
}

export function patchesDifferFromReview(review: TemplateMatterReview, patches: TemplateMatterPatch[] | undefined) {
  if (!patches?.length) return false;
  const originals = new Map(writableMatterFindings(review).map((finding) => [finding.id, finding.value]));
  return patches.some((patch) => (originals.get(patch.findingId) ?? "") !== patch.value);
}

export function resolvedBundleTitle(fallback: string, sources: Array<TemplateMatterValues | undefined>) {
  for (const source of sources) {
    if (!source) continue;
    const title = source.matterTitles.find((item) => item.trim());
    if (title) return title.trim();
    const parties = source.partyNames.map((item) => item.trim()).filter(Boolean);
    if (parties.length >= 2) return `${parties[0]} v ${parties[1]}`;
    if (parties[0]) return parties[0];
  }
  return fallback;
}

function comparisonValues(reference: TemplateReviewReference, property: "matterNumbers" | "partyNames" | "forums" | "matterTitles") {
  if (reference.confirmedValues) {
    const values = [...reference.confirmedValues[property]].sort((left, right) => left.localeCompare(right, "en-GB"));
    return {
      values,
      normalizedValues: [...new Set(values.map(normalizeMatterValue).filter(Boolean))].sort((left, right) => left.localeCompare(right, "en-GB")),
    };
  }
  const findings = reference.review[property];
  if (!findings.length) return null;
  return {
    values: findings.map((finding) => finding.value).sort((left, right) => left.localeCompare(right, "en-GB")),
    normalizedValues: [...new Set(findings.map((finding) => finding.normalizedValue))].sort((left, right) => left.localeCompare(right, "en-GB")),
  };
}

/** Returns stable, source-hash-bound evidence where reviewed templates differ. */
export function compareTemplateMatterReviews(references: TemplateReviewReference[]): TemplateMatterDiscrepancy[] {
  const ordered = [...references].sort((left, right) =>
    left.role.localeCompare(right.role, "en-GB") ||
    left.templateId.localeCompare(right.templateId, "en-GB") ||
    left.sourceName.localeCompare(right.sourceName, "en-GB"));
  const discrepancies: TemplateMatterDiscrepancy[] = [];
  for (const [field, property, label] of DISCREPANCY_FIELDS) {
    const evidence = ordered.flatMap((reference) => {
      const values = comparisonValues(reference, property);
      if (!values) return [];
      return [{
        templateId: reference.templateId,
        role: reference.role,
        sourceName: reference.sourceName,
        pdfSha256: reference.review.pdfSha256,
        values: values.values,
        normalizedValues: values.normalizedValues,
      }];
    });
    if (evidence.length < 2) continue;
    const distinct = new Set(evidence.map((item) => item.normalizedValues.join("\u001f")));
    if (distinct.size < 2) continue;
    discrepancies.push({
      field,
      message: `The selected templates show different possible ${label}. Review the exact previews; Exhibit Builder has not decided which value is correct.`,
      evidence,
      unverified: true,
    });
  }
  return discrepancies;
}
