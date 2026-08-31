import type { ExhibitCandidate } from "./bundle-engine.ts";
import { restoreCitedCandidateDecision } from "./candidate-restore.ts";
import { bundleArrangementFromLegacyOrder, validateBundleArrangement, type BundleArrangement } from "./bundle-arrangement.ts";
import { lockPagination, type BuildResolution, type BundleLayoutSettings, type NonA4PageHandling, type PageNumberSettings, type ProjectSnapshot, type ProjectSource, type StoredTemplateReview, type TemplateDiscrepancyConfirmation, type TemplateSlot } from "./bundle-types.ts";

export type RecoverySourceDescriptor = {
  id: string;
  role: "statement" | "evidence" | "template" | "template-rendered" | "project";
  name: string;
  path: string;
  sha256: string;
  size: number;
};

export type RecoveryProjectPayload = {
  project?: { name?: string };
  candidates?: Array<Partial<ExhibitCandidate> & { id: string; sourceSha256?: string | null; statementSha256?: string | null }>;
  arrangement?: BundleArrangement;
  /** @deprecated Recovery journals created before project schema 8 only. */
  finalOrder?: string[];
  layout?: BundleLayoutSettings;
  pagination?: PageNumberSettings;
  pageSizeChoices?: Record<string, NonA4PageHandling>;
  resolutions?: BuildResolution[];
  statements?: Array<{ id: string; witnessName: string; witnessInitials: string; sourceId: string; sourceSha256?: string }>;
  templates?: Array<{ sourceId: string; slot: TemplateSlot; sourceFormat: "pdf" | "docx" | "doc"; /** @deprecated */ templateConfirmed?: boolean }>;
  templateReviews?: StoredTemplateReview[];
  templateDiscrepancyConfirmation?: TemplateDiscrepancyConfirmation;
  sources: RecoverySourceDescriptor[];
  fingerprint?: string | null;
};

type OpenedProject = { snapshot: ProjectSnapshot; sources: ProjectSource[] };
type SavedCandidate = NonNullable<RecoveryProjectPayload["candidates"]>[number];

function emailChildHash(candidate: { sourceSha256?: string | null; parentEmailProvenance?: { childSha256?: string } | null; evidenceId?: string | null }) {
  return candidate.parentEmailProvenance?.childSha256 ?? null;
}

function recoveryHashMatches(
  saved: { sourceSha256?: string | null; evidenceId?: string | null; parentEmailProvenance?: { childSha256?: string } | null },
  candidate: { evidenceId?: string | null; parentEmailProvenance?: { childSha256?: string } | null },
  sourceHashes: Map<string, string>,
) {
  const evidenceId = saved.evidenceId ?? candidate.evidenceId;
  const actualHash = evidenceId ? sourceHashes.get(evidenceId) : undefined;
  if (saved.sourceSha256 && actualHash === saved.sourceSha256) return true;
  const childHash = emailChildHash(saved) ?? emailChildHash(candidate);
  return Boolean(saved.sourceSha256 && childHash && saved.sourceSha256 === childHash);
}

function savedDecisionFor(candidate: ExhibitCandidate, deltas: Map<string, SavedCandidate>, used: Set<string>) {
  const byId = deltas.get(candidate.id);
  if (byId) return byId;
  const childHash = emailChildHash(candidate);
  if (!childHash) return undefined;
  for (const saved of deltas.values()) {
    if (used.has(saved.id)) continue;
    if (emailChildHash(saved) === childHash || saved.sourceSha256 === childHash) return saved;
  }
}

/**
 * Uses a verified saved project as the baseline and overlays only journal
 * decisions whose source-hash bindings still match the embedded project
 * sources.  This prevents both post-save edit loss and stale approval reuse.
 */
export function mergeRecoveryProjectDeltas(opened: OpenedProject, payload: RecoveryProjectPayload, recoveredSources: ProjectSource[] = []) {
  const issues: string[] = [];
  const currentStatementIds = new Set((payload.statements ?? []).map((statement) => statement.sourceId));
  const currentTemplateIds = new Set((payload.templates ?? []).map((template) => template.sourceId));
  const sourceById = new Map(opened.sources
    .filter((source) => source.role === "evidence" || (source.role === "statement" ? !payload.statements || currentStatementIds.has(source.id) : currentTemplateIds.has(source.id)))
    .map((source) => [source.id, source]));
  for (const source of recoveredSources) sourceById.set(source.id, source);
  const sources = [...sourceById.values()];
  const sourceHashes = new Map(sources.map((source) => [source.id, source.sha256]));
  const deltas = new Map((payload.candidates ?? []).map((candidate) => [candidate.id, candidate]));
  const validCandidateDeltas = new Set<string>();
  const usedSavedIds = new Set<string>();
  const baselineCandidates = opened.snapshot.candidates as ExhibitCandidate[];
  const candidates = baselineCandidates.map((candidate) => {
    const saved = savedDecisionFor(candidate, deltas, usedSavedIds);
    if (!saved) return candidate;
    usedSavedIds.add(saved.id);
    const statementDescriptor = (payload.statements ?? []).find((statement) => statement.id === (saved.statementId ?? candidate.statementId));
    const currentStatementSha256 = statementDescriptor ? sourceHashes.get(statementDescriptor.sourceId) : undefined;
    const statementHashMatches = Boolean(saved.manualAddition || (saved.statementSha256
      && statementDescriptor?.sourceSha256
      && saved.statementSha256 === statementDescriptor.sourceSha256
      && currentStatementSha256 === statementDescriptor.sourceSha256));
    if (saved.sourceSha256 && !recoveryHashMatches(saved, candidate, sourceHashes)) {
      issues.push(`Recovery decisions for ${candidate.description || candidate.id} were not restored because the source hash changed.`);
      return { ...candidate, confirmed: false };
    }
    const { id: _id, sourceSha256: _sourceSha256, statementSha256: _statementSha256, ...decision } = saved;
    if (saved.sourceSha256 && statementHashMatches) validCandidateDeltas.add(candidate.id);
    if (saved.confirmed && !statementHashMatches) issues.push(`Recovery approval for ${candidate.description || candidate.id} was not restored because the witness statement changed or its analysis hash is unavailable.`);
    return restoreCitedCandidateDecision(candidate, saved as ExhibitCandidate, decision.evidenceId ?? candidate.evidenceId ?? null, Boolean(saved.sourceSha256 && statementHashMatches));
  });
  const existingCandidateIds = new Set(candidates.map((candidate) => candidate.id));
  const existingChildHashes = new Set(candidates.map((candidate) => emailChildHash(candidate)).filter((hash): hash is string => Boolean(hash)));
  for (const saved of payload.candidates ?? []) {
    if (usedSavedIds.has(saved.id) || existingCandidateIds.has(saved.id)) continue;
    const childHash = emailChildHash(saved);
    if (childHash && existingChildHashes.has(childHash)) continue;
    if (!saved.sourceSha256 || !recoveryHashMatches(saved, saved, sourceHashes)) continue;
    const statementDescriptor = (payload.statements ?? []).find((statement) => statement.id === saved.statementId);
    const statementHashMatches = Boolean(saved.manualAddition || (saved.statementSha256
      && statementDescriptor?.sourceSha256
      && saved.statementSha256 === statementDescriptor.sourceSha256
      && sourceHashes.get(statementDescriptor.sourceId) === statementDescriptor.sourceSha256));
    const { sourceSha256: _sourceSha256, statementSha256: _statementSha256, ...decision } = saved;
    candidates.push({
      ...decision,
      confirmed: statementHashMatches ? Boolean(decision.confirmed) : false,
      repeatDecision: statementHashMatches ? decision.repeatDecision : undefined,
    } as ExhibitCandidate);
    if (saved.confirmed && !statementHashMatches) issues.push(`Recovery approval for ${saved.description || saved.id} was not restored because the witness statement changed or its analysis hash is unavailable.`);
    validCandidateDeltas.add(saved.id);
    usedSavedIds.add(saved.id);
    if (childHash) existingChildHashes.add(childHash);
  }

  const validResolutions = (payload.resolutions ?? []).filter((resolution) => {
    if (resolution.sourceId && resolution.sourceSha256 && sourceHashes.get(resolution.sourceId) !== resolution.sourceSha256) return false;
    if (resolution.candidateId && !validCandidateDeltas.has(resolution.candidateId)) return false;
    return Object.entries(resolution.templateHashes ?? {}).every(([slot, hash]) => sourceHashes.get(`template-${slot}`) === hash);
  });

  const expectedSourceHashes = new Map(payload.sources.map((source) => [source.id, source.sha256]));
  const templateApprovals = new Map<string, boolean>();
  for (const template of payload.templates ?? []) {
    if (sourceHashes.get(template.sourceId) === expectedSourceHashes.get(template.sourceId)) {
      templateApprovals.set(template.sourceId, Boolean(template.templateConfirmed));
    } else {
      issues.push(`Recovery approval for ${template.slot} template was not restored because the template source is missing or changed.`);
    }
  }

  const priorAnalysis = (opened.snapshot.analysis ?? {}) as Record<string, unknown>;
  const sourceNames = new Map(sources.map((source) => [source.id, source.name]));
  const analysis = {
    ...priorAnalysis,
    ...(payload.statements ? { statements: payload.statements.map((statement) => ({
      id: statement.id,
      witnessName: statement.witnessName,
      witnessInitials: statement.witnessInitials,
      name: sourceNames.get(statement.sourceId) ?? statement.sourceId,
    })) } : {}),
    ...(payload.templates ? { templates: payload.templates.map((template) => template.slot) } : {}),
  };

  const arrangement = payload.arrangement
    ? validateBundleArrangement(payload.arrangement)
    : payload.finalOrder
      ? bundleArrangementFromLegacyOrder(payload.finalOrder)
      : opened.snapshot.arrangement
        ? validateBundleArrangement(opened.snapshot.arrangement)
        : bundleArrangementFromLegacyOrder(opened.snapshot.finalOrder);
  const { finalOrder: _legacyFinalOrder, ...currentSnapshot } = opened.snapshot;

  return {
    opened: {
      ...opened,
      sources,
      snapshot: {
        ...currentSnapshot,
        schemaVersion: 8,
        name: payload.project?.name ?? opened.snapshot.name,
        candidates,
        analysis,
        arrangement,
        layout: payload.layout ?? opened.snapshot.layout,
        pagination: lockPagination({ ...opened.snapshot.pagination, ...(payload.pagination ?? {}) }),
        pageSizeChoices: payload.pageSizeChoices ?? opened.snapshot.pageSizeChoices ?? {},
        resolutions: validResolutions,
        templateReviews: payload.templateReviews ?? opened.snapshot.templateReviews,
        templateDiscrepancyConfirmation: payload.templateDiscrepancyConfirmation ?? opened.snapshot.templateDiscrepancyConfirmation,
      },
    } satisfies OpenedProject,
    templateApprovals,
    issues,
  };
}
