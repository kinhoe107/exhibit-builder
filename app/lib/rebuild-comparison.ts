import type { AnalysisResult, ExhibitCandidate } from "./bundle-engine.ts";
import { bundleArrangementFromLegacyOrder, flattenBundleArrangement, validateBundleArrangement, type BundleArrangement } from "./bundle-arrangement.ts";
import type { BundleLayoutSettings, BuildResolution, NonA4PageHandling, PageNumberSettings, TemplateFile } from "./bundle-types.ts";

export type BuildFingerprintInput = {
  analysis: AnalysisResult;
  candidates: ExhibitCandidate[];
  arrangement?: BundleArrangement;
  /** @deprecated Transitional input for callers not yet upgraded to schema 8. */
  canonicalOrder?: string[];
  templates: TemplateFile[];
  layout: BundleLayoutSettings;
  pagination: PageNumberSettings;
  pageSizeChoices?: Record<string, NonA4PageHandling>;
  resolutions: BuildResolution[];
  volumePlan?: unknown;
  statementSuggestions?: string[];
};

export type SubstantiveBuildSnapshot = {
  fingerprint: string;
  canonical: Record<string, unknown>;
};

export type RebuildDifference = {
  changed: boolean;
  categories: string[];
  summary: string;
};

function stable(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(stable);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(Object.entries(value as Record<string, unknown>)
    .filter(([, nested]) => nested !== undefined)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([key, nested]) => [key, stable(nested)]));
}

async function digest(value: unknown) {
  const bytes = new TextEncoder().encode(JSON.stringify(stable(value)));
  const hash = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(hash)).map((part) => part.toString(16).padStart(2, "0")).join("");
}

export async function createSubstantiveBuildSnapshot(input: BuildFingerprintInput): Promise<SubstantiveBuildSnapshot> {
  const arrangement = input.arrangement
    ? validateBundleArrangement(input.arrangement)
    : bundleArrangementFromLegacyOrder(input.canonicalOrder ?? []);
  const evidenceById = new Map(input.analysis.evidence.map((record) => [record.id, record]));
  const usedTemplateSlots = new Set<TemplateFile["slot"]>([
    "cover",
    "index",
    ...(input.layout.includeDividerPages ? ["divider" as const] : []),
    ...(input.layout.includeExhibitCoverPages ? ["exhibitCover" as const] : []),
  ]);
  const included = input.candidates.filter((candidate) => candidate.included).map((candidate) => {
    const evidence = candidate.evidenceId ? evidenceById.get(candidate.evidenceId) : undefined;
    return {
      id: candidate.id,
      evidenceId: candidate.evidenceId,
      sourceSha256: evidence?.sha256 ?? null,
      pageStart: candidate.pageStart ?? null,
      pageEnd: candidate.pageEnd ?? null,
      description: candidate.description,
      date: candidate.date,
      manualAddition: candidate.manualAddition ?? false,
      manualWarningAcknowledgedAt: candidate.manualWarningAcknowledgedAt ?? null,
      repeatDecision: candidate.repeatDecision ?? null,
      confirmed: candidate.confirmed,
      sheets: evidence?.workbook?.sheets.map((sheet) => ({
        id: sheet.id,
        included: evidence.sheetSelections?.find((selection) => selection.name === sheet.name)?.included ?? false,
        range: evidence.sheetSelections?.find((selection) => selection.name === sheet.name)?.range ?? sheet.renderPlan.range,
        renderPlanHash: sheet.renderPlan.planHash,
      })) ?? [],
    };
  });
  const canonical = {
    statements: input.analysis.statementHash.split(",").map((hash) => hash.trim()).filter(Boolean),
    arrangement,
    order: flattenBundleArrangement(arrangement),
    exhibits: included,
    templates: input.templates
      .filter((template) => usedTemplateSlots.has(template.slot))
      .map((template) => ({
        slot: template.slot,
        sha256: template.sha256,
        pdfSha256: template.pdfSha256 ?? null,
        matterReviewPdfSha256: template.reviewState?.matterReview?.pdfSha256 ?? null,
        matterConfirmedPdfSha256: template.reviewState?.matterConfirmation?.pdfSha256 ?? null,
        confirmedMatterNumbers: template.reviewState?.matterConfirmation?.matterNumbers ?? null,
        confirmedPartyNames: template.reviewState?.matterConfirmation?.partyNames ?? null,
        confirmedForums: template.reviewState?.matterConfirmation?.forums ?? null,
        confirmedMatterTitles: template.reviewState?.matterConfirmation?.matterTitles ?? null,
        confirmedMatterPatches: template.reviewState?.matterConfirmation?.patches ?? null,
        appearanceConfirmedPdfSha256: template.reviewState?.appearanceConfirmation?.pdfSha256 ?? null,
        placeholderConfirmedPdfSha256: template.reviewState?.placeholderConfirmation?.pdfSha256 ?? null,
      }))
      .sort((left, right) => left.slot.localeCompare(right.slot)),
    layout: input.layout,
    pagination: input.pagination,
    pageSizeChoices: input.pageSizeChoices ?? {},
    resolutions: input.resolutions.map((resolution) => ({
      blockerId: resolution.blockerId,
      action: resolution.action,
      sourceId: resolution.sourceId ?? null,
      sourceSha256: resolution.sourceSha256 ?? null,
      candidateId: resolution.candidateId ?? null,
      templateSlots: resolution.templateSlots ?? [],
      templateHashes: resolution.templateHashes ?? {},
      note: resolution.note ?? "",
      visualReviewConfirmed: resolution.visualReviewConfirmed ?? false,
    })),
    volumePlan: input.volumePlan ?? null,
    statementSuggestions: input.statementSuggestions ?? [],
  };
  return { canonical, fingerprint: await digest(canonical) };
}

function differs(left: Record<string, unknown>, right: Record<string, unknown>, key: string) {
  return JSON.stringify(stable(left[key])) !== JSON.stringify(stable(right[key]));
}

function comparableCanonical(canonical: Record<string, unknown>) {
  return {
    ...canonical,
    arrangement: canonical.arrangement ?? bundleArrangementFromLegacyOrder(Array.isArray(canonical.order) ? canonical.order : []),
  };
}

export function compareSubstantiveBuilds(previous: SubstantiveBuildSnapshot | null, current: SubstantiveBuildSnapshot): RebuildDifference {
  if (!previous) return { changed: true, categories: ["This is the first version of this bundle."], summary: "Bundle created." };
  if (previous.fingerprint === current.fingerprint) return { changed: false, categories: [], summary: "No substantive change" };
  const previousComparable = comparableCanonical(previous.canonical);
  const currentComparable = comparableCanonical(current.canonical);
  // Adding schema-8 structure around an unchanged legacy flat order is a data
  // migration, not a reason to tell the user that their bundle changed.
  if (JSON.stringify(stable(previousComparable)) === JSON.stringify(stable(currentComparable))) return { changed: false, categories: [], summary: "No substantive change" };
  const categories: string[] = [];
  const orderChanged = differs(previous.canonical, current.canonical, "order");
  if (orderChanged) categories.push("Order changed");
  if (!orderChanged && differs(previousComparable, currentComparable, "arrangement")) categories.push("Index sections changed");
  if (differs(previous.canonical, current.canonical, "statements")) categories.push("Statement source changed");
  const previousExhibits = new Map(((previous.canonical.exhibits as Array<Record<string, unknown>>) ?? []).map((item) => [item.id, item]));
  const currentExhibits = new Map(((current.canonical.exhibits as Array<Record<string, unknown>>) ?? []).map((item) => [item.id, item]));
  const exhibitIds = new Set([...previousExhibits.keys(), ...currentExhibits.keys()]);
  const exhibitFieldChanged = (fields: string[]) => [...exhibitIds].some((id) => fields.some((field) => JSON.stringify(stable(previousExhibits.get(id)?.[field])) !== JSON.stringify(stable(currentExhibits.get(id)?.[field]))));
  if (exhibitFieldChanged(["evidenceId", "sourceSha256"])) categories.push("Source replaced or hash changed");
  if (exhibitFieldChanged(["pageStart", "pageEnd"])) categories.push("Selected source page range changed");
  if (exhibitFieldChanged(["sheets"])) categories.push("Worksheet selection or render plan changed");
  if (exhibitFieldChanged(["description"])) categories.push("Index description changed");
  if (exhibitFieldChanged(["date"])) categories.push("Document date changed");
  if (exhibitFieldChanged(["manualAddition"])) categories.push("Manual exhibit status changed");
  if (exhibitFieldChanged(["manualWarningAcknowledgedAt"])) categories.push("Manual exhibit acknowledgement changed");
  if (exhibitFieldChanged(["repeatDecision"])) categories.push("Repeat decision changed");
  if (exhibitFieldChanged(["confirmed"]) || previousExhibits.size !== currentExhibits.size) categories.push("Included exhibit decisions changed");
  if (differs(previous.canonical, current.canonical, "templates")) categories.push("Template changed");
  const previousLayout = (previous.canonical.layout ?? {}) as Record<string, unknown>;
  const currentLayout = (current.canonical.layout ?? {}) as Record<string, unknown>;
  if (["includeDividerPages", "includeExhibitCoverPages", "countOptionalPagesInReferences"].some((key) => previousLayout[key] !== currentLayout[key])) categories.push("Optional-page rule changed");
  if (previousLayout.volumePageLimit !== currentLayout.volumePageLimit) categories.push("Volume limit changed");
  if (previousLayout.coverInsertion !== currentLayout.coverInsertion || previousLayout.exactCoverPageNumber !== currentLayout.exactCoverPageNumber || previousLayout.exactCoverVolumeLabel !== currentLayout.exactCoverVolumeLabel) categories.push("Cover treatment changed");
  if (differs(previous.canonical, current.canonical, "volumePlan")) categories.push("Volume boundary changed");
  if (differs(previous.canonical, current.canonical, "pagination")) categories.push("Pagination changed");
  if (differs(previous.canonical, current.canonical, "pageSizeChoices")) categories.push("Non-A4 page treatment changed");
  if (differs(previous.canonical, current.canonical, "resolutions")) categories.push("Build resolution or OCR exception changed");
  if (differs(previous.canonical, current.canonical, "statementSuggestions")) categories.push("Statement suggestions changed");
  return { changed: true, categories, summary: categories.join("; ") || "Substantive build inputs changed" };
}
