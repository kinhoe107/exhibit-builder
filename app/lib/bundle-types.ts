import type { BundleArrangement } from "./bundle-arrangement.ts";
import type { TemplateMatterPatch, TemplateMatterReview, TemplateMatterValues } from "./template-matter-review.ts";

export type Severity = "pass" | "warning" | "blocking";
export type PreflightPolicy = "hard-legal" | "hard-technical" | "exception-eligible" | "warning";
export type PreflightCode = string;

export type PageNumberSettings = {
  /** Print the physical PDF page number on every page, including cover/index. */
  matchPdfPageOrder: boolean;
  /** Continue visible numbers across separate volume PDFs unless restart is explicitly accepted. */
  volumeNumbering: "continuous" | "restart";
  scheme: "bundle" | "bates" | "section";
  prefix: string;
  suffix: string;
  startAt: number;
  padding: number;
  preliminary: "arabic" | "roman" | "none";
  countTemplates: boolean;
  position: "bottom-left" | "bottom-centre" | "bottom-right" | "top-left" | "top-centre" | "top-right" | "inside-bottom" | "outside-bottom";
  fontSize: number;
  /**
   * Retained for saved projects. Stamps, the index page column and statement
   * suggestions always use the same visible prefix as the finished PDF.
   */
  includePrefixInIndex: boolean;
};

export type NonA4PageHandling = "convert-to-a4" | "keep-original";
export type CoverInsertion = "fit-a4" | "exact";

/**
 * Final assembly choices are deliberately separate from template selection.
 * A template is only a candidate asset; it cannot enter the PDF until the
 * reviewer explicitly includes that page type.
 */
export type BundleLayoutSettings = {
  includeDividerPages: boolean;
  includeExhibitCoverPages: boolean;
  /** When false, optional pages never advance the legal [AH1/x] cursor. */
  countOptionalPagesInReferences: boolean;
  /** Zero keeps a single PDF. A positive value creates administrative volumes. */
  volumePageLimit: number;
  /**
   * How a selected custom cover is finished. Both choices still fit the cover
   * to A4 without cropping. Exact mode hides amend fields and does not rewrite
   * cover text. The index is never exact: exhibit rows must still be written
   * into it, and matter-text corrections are still printed there.
   */
  coverInsertion: CoverInsertion;
  /** Exact covers receive a printed page number only when the reviewer opts in. */
  exactCoverPageNumber: boolean;
  /** Exact covers receive a volume label only when the reviewer opts in. */
  exactCoverVolumeLabel: boolean;
  /** Used on generated cover/index pages when no custom template supplies those details. */
  builtInMatter: TemplateMatterValues;
};

export const DEFAULT_BUNDLE_LAYOUT: BundleLayoutSettings = {
  includeDividerPages: false,
  includeExhibitCoverPages: false,
  countOptionalPagesInReferences: false,
  volumePageLimit: 0,
  coverInsertion: "fit-a4",
  exactCoverPageNumber: false,
  exactCoverVolumeLabel: false,
  builtInMatter: { matterNumbers: [], partyNames: [], forums: [], matterTitles: [] },
};

export function coverPrintsPageNumber(layout: BundleLayoutSettings) {
  return layout.coverInsertion !== "exact" || layout.exactCoverPageNumber;
}

export function coverPrintsVolumeLabel(layout: BundleLayoutSettings) {
  return layout.coverInsertion !== "exact" || layout.exactCoverVolumeLabel;
}

export function coverWritesMatterText(layout: BundleLayoutSettings) {
  return layout.coverInsertion !== "exact";
}

export type BundleProfile = {
  id: string;
  name: string;
  description: string;
  requireOcr: boolean;
  pageSize: "a4";
  requireBookmarks: boolean;
  requireLinkedIndex: boolean;
  requireConfirmedExhibits: boolean;
  numbering: "witness-sequence" | "bundle-sequence";
};

export const BUNDLE_PROFILES: BundleProfile[] = [
  {
    id: "exhibit-neutral",
    name: "Neutral",
    description: "A4, searchable output, linked index and exhibit bookmarks.",
    requireOcr: true,
    pageSize: "a4",
    requireBookmarks: true,
    requireLinkedIndex: true,
    requireConfirmedExhibits: true,
    numbering: "witness-sequence",
  },
  {
    id: "review-draft",
    name: "Automated review draft",
    description: "Test-only profile that records unavailable OCR as a warning while exercising the remaining build pipeline.",
    requireOcr: false,
    pageSize: "a4",
    requireBookmarks: true,
    requireLinkedIndex: true,
    requireConfirmedExhibits: true,
    numbering: "witness-sequence",
  },
];

export const DEFAULT_PAGINATION: PageNumberSettings = {
  matchPdfPageOrder: true,
  volumeNumbering: "continuous",
  scheme: "bundle",
  prefix: "",
  suffix: "",
  startAt: 1,
  padding: 0,
  preliminary: "arabic",
  countTemplates: true,
  position: "bottom-centre",
  fontSize: 8,
  includePrefixInIndex: true,
};

const APPLY_GATED_PAGINATION_KEYS = ["prefix", "suffix", "padding", "startAt"] as const;

/** Apply a control change to the complete visible draft, not stale committed state. */
export function updatePaginationDraft(currentDraft: PageNumberSettings, change: Partial<PageNumberSettings>): PageNumberSettings {
  return { ...currentDraft, ...change };
}

/**
 * Instant commits merge into the last applied pagination so an unapplied
 * prefix, suffix, padding or start-at cannot leak. Apply (empty change)
 * commits the visible draft. Apply-gated fields in an instant change still
 * update committed pagination (and the warning object) but are not written
 * into the draft until Apply or accept.
 */
export function commitPaginationChange(
  committed: PageNumberSettings,
  draft: PageNumberSettings,
  change: Partial<PageNumberSettings>,
): { pagination: PageNumberSettings; draft: PageNumberSettings } {
  if (Object.keys(change).length === 0) {
    const applied = { ...draft };
    return { pagination: applied, draft: applied };
  }
  const applyGated = new Set<string>(APPLY_GATED_PAGINATION_KEYS);
  const draftChange = Object.fromEntries(Object.entries(change).filter(([key]) => !applyGated.has(key)));
  return {
    pagination: { ...committed, ...change },
    draft: { ...draft, ...draftChange },
  };
}

/** Cancel a numbering warning: restore instant radios from committed state, keep unapplied gated fields. */
export function paginationDraftAfterWarningCancel(
  committed: PageNumberSettings,
  draft: PageNumberSettings,
): PageNumberSettings {
  return {
    ...committed,
    prefix: draft.prefix,
    suffix: draft.suffix,
    padding: draft.padding,
    startAt: draft.startAt,
  };
}

/** Accept a numbering warning: take gated values from the change, keep other unapplied gated fields. */
export function paginationDraftAfterWarningAccept(
  confirmation: PageNumberSettings,
  draft: PageNumberSettings,
  change: Partial<PageNumberSettings>,
): PageNumberSettings {
  const preserved = Object.fromEntries(
    APPLY_GATED_PAGINATION_KEYS.filter((key) => !(key in change)).map((key) => [key, draft[key]]),
  );
  return { ...confirmation, ...preserved };
}

export type TemplateSlot = "cover" | "exhibitCover" | "index" | "divider";

export type TemplateReviewConfirmation = {
  pdfSha256: string;
  confirmedAt: string;
};

/** Matter confirmation may also store the reviewer-corrected identifying list and per-occurrence patches. */
export type TemplateMatterConfirmation = TemplateReviewConfirmation & Partial<TemplateMatterValues> & {
  patches?: TemplateMatterPatch[];
};

export type TemplateReviewState = {
  /** Possible identifying details read from the exact rendered PDF. */
  matterReview?: TemplateMatterReview;
  /** Required only when a Word source was converted to PDF. */
  appearanceConfirmation?: TemplateReviewConfirmation;
  /** Required for every selected custom template, including a supplied PDF. */
  matterConfirmation?: TemplateMatterConfirmation;
  /** Required when the exact reviewed artifact contains possible placeholders. */
  placeholderConfirmation?: TemplateReviewConfirmation;
};

export type TemplateDiscrepancyConfirmation = {
  /** Stable identity of the exact reviewed template set and discrepancies. */
  fingerprint: string;
  confirmedAt: string;
};

export type TemplateFile = {
  slot: TemplateSlot;
  file: File;
  sha256: string;
  /** Original Word/PDF source retained for project persistence and audit. */
  sourceFormat?: "pdf" | "docx" | "doc";
  /** Local PDF normalisation cached for the current desktop session. */
  pdfFile?: File;
  pdfSha256?: string;
  reviewState?: TemplateReviewState;
  /** @deprecated Project schemas before 8 used one ambiguous Word-only flag. */
  templateConfirmed?: boolean;
};

export type StoredTemplateReview = {
  slot: TemplateSlot;
  sourceId: string;
  renderedSourceId?: string;
  sourceFormat: "pdf" | "docx" | "doc";
  sourceSha256: string;
  pdfSha256: string;
  reviewState: TemplateReviewState;
};

/**
 * A reviewer-approved resolution for a build gate.  Resolutions are tied to
 * the source hash where possible so replacing a document cannot silently
 * inherit a decision made for an earlier file.
 */
export type BuildResolutionAction =
  | "proceed-without-ocr"
  | "exclude-source"
  | "exclude-candidate"
  | "use-built-in-template";

export type BuildResolution = {
  blockerId: string;
  checkCode?: PreflightCode;
  profileId?: string;
  action: BuildResolutionAction;
  sourceId?: string;
  sourceSha256?: string;
  candidateId?: string;
  fileName?: string;
  templateSlots?: TemplateSlot[];
  templateHashes?: Partial<Record<TemplateSlot, string>>;
  approvedAt: string;
  note?: string;
  visualReviewConfirmed?: boolean;
};

export type ProjectSource = {
  id: string;
  role: "statement" | "evidence" | "template" | "template-rendered";
  name: string;
  sha256: string;
  file: File;
};
export type SheetSelection = { sourceSha256: string; sheetId: string; sheetPath: string; sheetName: string; included: boolean; range: string; renderPlanHash?: string };

export type ProjectSnapshot = {
  schemaVersion: 2 | 3 | 4 | 5 | 6 | 7 | 8;
  name: string;
  createdAt: string;
  updatedAt: string;
  profileId: string;
  pagination: PageNumberSettings;
  witnessSettings: Record<string, { initials: string; nextNumber: number }>;
  candidates: unknown[];
  analysis: unknown;
  sheetSelections?: SheetSelection[];
  resolutions?: BuildResolution[];
  layout?: BundleLayoutSettings;
  /** Authoritative order and index-section structure from project schema 8. */
  arrangement?: BundleArrangement;
  /** Hash-bound review state and rendered-artifact identity for custom templates. */
  templateReviews?: StoredTemplateReview[];
  /** Required when selected templates contain conflicting possible matter details. */
  templateDiscrepancyConfirmation?: TemplateDiscrepancyConfirmation;
  /** @deprecated Read only while migrating project schemas 2-7. */
  finalOrder?: string[];
  lastBuildSnapshot?: unknown;
  /** Per-source treatment for PDF pages that are not already A4. */
  pageSizeChoices?: Record<string, NonA4PageHandling>;
};

export type PreflightCheck = {
  id: string;
  code?: PreflightCode;
  policy?: PreflightPolicy;
  severity: Severity;
  label: string;
  detail: string;
  fileName?: string;
  page?: number;
  sourceId?: string;
  sourceSha256?: string;
  /** Stable related evidence identities for advisory relationships. */
  relatedSourceIds?: string[];
  candidateId?: string;
};
