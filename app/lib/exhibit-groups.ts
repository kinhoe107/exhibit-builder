import type { AnalysisResult, EvidenceRecord, ExhibitCandidate } from "./bundle-engine.ts";
import {
  flattenBundleArrangement,
  reconcileBundleArrangement,
  validateBundleArrangement,
  type BundleArrangement,
} from "./bundle-arrangement.ts";
import { unresolvedEmailAttachments } from "./email-attachments.ts";

/** Automatic matches below this score need individual review and copy qualification. */
export const AUTOMATIC_MATCH_REVIEW_THRESHOLD = 65;
export const REVIEWER_SELECTED_RATIONALE = "Source selected by reviewer";
export const LOW_CONFIDENCE_SOURCE_QUALIFIER = "low-confidence source match confirmed by reviewer; verify the selected source before copying";

export type ReviewProvenance = {
  candidateId: string;
  confidence: number;
  rationale: string;
  confirmed: true;
  confirmationMethod?: "individual" | "bulk";
  confirmedAt?: string;
};

export type ExhibitReference = Pick<
  ExhibitCandidate,
  "paragraph" | "citation" | "statementName" | "statementId" | "witnessInitials" | "citationToken" | "citationOrdinal" | "citationCount" | "exhibitInitials" | "exhibitSequence" | "requestedExhibitPageStart" | "requestedExhibitPageEnd" | "citationResolution"
> & {
  reviewProvenance?: ReviewProvenance;
};

export function isReviewerSelectedSource(candidate: Pick<ExhibitCandidate, "rationale">) {
  return candidate.rationale === REVIEWER_SELECTED_RATIONALE;
}

export function isAutomaticLowConfidenceMatch(candidate: Pick<ExhibitCandidate, "confidence" | "rationale">) {
  return !isReviewerSelectedSource(candidate) && candidate.confidence < AUTOMATIC_MATCH_REVIEW_THRESHOLD;
}

export function reviewProvenanceFromCandidate(candidate: ExhibitCandidate): ReviewProvenance | undefined {
  if (!candidate.confirmed) return undefined;
  return {
    candidateId: candidate.id,
    confidence: candidate.confidence,
    rationale: candidate.rationale,
    confirmed: true,
    confirmationMethod: candidate.confirmationMethod,
    confirmedAt: candidate.confirmedAt,
  };
}

export function statementReferenceFromCandidate(
  candidate: ExhibitCandidate,
  extra: Record<string, unknown> = {},
): ExhibitReference & Record<string, unknown> {
  const reviewProvenance = reviewProvenanceFromCandidate(candidate);
  return {
    paragraph: candidate.paragraph,
    citation: candidate.citation,
    statementName: candidate.statementName,
    statementId: candidate.statementId,
    witnessInitials: candidate.witnessInitials,
    citationToken: candidate.citationToken,
    citationOrdinal: candidate.citationOrdinal,
    citationCount: candidate.citationCount,
    exhibitInitials: candidate.exhibitInitials,
    exhibitSequence: candidate.exhibitSequence,
    requestedExhibitPageStart: candidate.requestedExhibitPageStart,
    requestedExhibitPageEnd: candidate.requestedExhibitPageEnd,
    citationResolution: candidate.citationResolution,
    ...extra,
    ...(reviewProvenance ? { reviewProvenance } : {}),
  };
}

export function bulkConfirmableCandidates(candidates: ExhibitCandidate[], groups: ExhibitGroup[]) {
  const repeatCitationCandidateIds = new Set(
    groups.flatMap((group) => group.collisionMembers.slice(1).map((member) => member.id)),
  );
  return candidates.filter((candidate) => {
    if (candidate.parentEmailProvenance) return false;
    if (isAutomaticLowConfidenceMatch(candidate)) return false;
    const group = groups.find((item) => item.canonical.id === candidate.id);
    return !repeatCitationCandidateIds.has(candidate.id)
      && candidate.included
      && Boolean(candidate.evidenceId)
      && !candidate.confirmed
      && (!group || (!group.decisionPending && !group.selectionConflict));
  });
}

export type ExhibitGroup = {
  id: string;
  canonical: ExhibitCandidate;
  outputMark: string;
  exhibitNumber: number;
  members: ExhibitCandidate[];
  collisionMembers: ExhibitCandidate[];
  evidence: EvidenceRecord;
  sourceHash: string;
  references: ExhibitReference[];
  collision: boolean;
  decisionPending: boolean;
  selectionConflict: boolean;
};

/**
 * Reorders a finalisation list from a single drop gesture. Dropping while
 * moving down places the item after its target; moving up places it before.
 */
export function reorderGroupForDrop<T extends { canonical: { id: string } }>(
  groups: T[],
  candidateId: string,
  targetId: string,
): T[] {
  if (candidateId === targetId) return groups;
  const ordered = [...groups];
  const from = ordered.findIndex((group) => group.canonical.id === candidateId);
  const to = ordered.findIndex((group) => group.canonical.id === targetId);
  if (from < 0 || to < 0) return groups;
  const [moved] = ordered.splice(from, 1);
  ordered.splice(to, 0, moved);
  return ordered;
}

/**
 * Gives review cards a single visible-order identity. Candidate provisional
 * numbers and bundle-scoped exhibit marks are discovery data; neither may
 * reset or duplicate the Item number shown in one review list. Build this map
 * before applying a UI filter so hidden cards do not renumber retained cards.
 */
export function reviewItemNumbers<T extends { id: string }>(
  reviewCandidates: T[],
) {
  return new Map(reviewCandidates.map((candidate, index) => [candidate.id, index + 1] as const));
}

function reviewSequence(candidate: Pick<ExhibitCandidate, "sequenceOrder" | "provisionalNumber">) {
  return candidate.sequenceOrder ?? candidate.provisionalNumber ?? 0;
}

/**
 * Review cards keep their own identity until the reviewer chooses “same”.
 * A later citation that has merely selected the same file must not disappear
 * into an already-confirmed canonical card.
 */
export function reviewCandidatesForDisplay(
  candidates: ExhibitCandidate[],
  groups: ExhibitGroup[],
) {
  const order = new Map(groups.map((group, index) => [group.canonical.id, index]));
  return candidates
    .filter((candidate) => !candidate.parentEmailProvenance)
    .filter((candidate) => {
      if (!candidate.evidenceId) return true;
      if (order.has(candidate.id)) return true;
      return candidate.included && candidate.repeatDecision !== "same" && candidate.repeatDecision !== "separate";
    })
    // Review is a walk through the witness statement, not a preview of the
    // later user-defined bundle order. Keep unmatched and repeat-decision
    // cards at their original citation positions as well.
    .sort((left, right) => reviewSequence(left) - reviewSequence(right));
}

export function repeatExhibitCount(groups: ExhibitGroup[]) {
  return groups.filter((group) => !group.canonical.parentEmailProvenance && (group.collisionMembers?.length ?? 0) > 1).length;
}

export function formatRepeatExhibitNote(count: number) {
  if (count <= 0) return "";
  return ` (${count} repeat exhibit${count === 1 ? "" : "s"})`;
}

export function exhibitGroupLookup(groups: ExhibitGroup[]) {
  const byGroupId = new Map<string, ExhibitGroup>();
  const byCandidateId = new Map<string, ExhibitGroup>();
  for (const group of groups) {
    byGroupId.set(group.id, group);
    byCandidateId.set(group.canonical.id, group);
    for (const member of group.collisionMembers) byCandidateId.set(member.id, group);
  }
  return { byGroupId, byCandidateId } as const;
}

function sequence(candidate: ExhibitCandidate) {
  // An explicit token such as [AH1/3-4] names the *bundle*, rather than the
  // individual document.  It must never control the order of individual
  // exhibits within that bundle.
  return candidate.sequenceOrder ?? candidate.provisionalNumber;
}

function bundleKey(candidate: ExhibitCandidate) {
  const initials = candidate.exhibitInitials ?? candidate.witnessInitials?.trim().toUpperCase() ?? "EX";
  const bundle = candidate.exhibitSequence ?? 1;
  return `${candidate.witnessKey ?? initials}:${initials}:${bundle}`;
}

function bundleCitation(candidate: ExhibitCandidate) {
  const initials = (candidate.exhibitInitials ?? candidate.witnessInitials?.trim().toUpperCase() ?? "EX").replace(/\s+/g, "").toUpperCase();
  return `${initials}${candidate.exhibitSequence ?? 1}`;
}

/** 0.11.8 and earlier persisted this source-hash identity in finalOrder. */
export function legacyGroupId(group: ExhibitGroup) {
  const base = `source-${bundleKey(group.canonical)}:hash:${group.sourceHash}`;
  const wasSeparateRepeat = group.collision && group.collisionMembers.length === 1 && group.canonical.repeatDecision === "separate";
  return wasSeparateRepeat ? `${base}-${group.canonical.id}` : base;
}

/**
 * Maps a unique current group ID or historical source-hash alias to the live
 * group ID. An alias claimed by more than one group is omitted so it cannot
 * silently reorder the wrong exhibit.
 */
export function unambiguousExhibitIdMap(groups: readonly ExhibitGroup[]): Map<string, string> {
  const claimed = new Map<string, Set<string>>();
  for (const group of groups) {
    for (const alias of [group.id, legacyGroupId(group)]) {
      const owners = claimed.get(alias) ?? new Set<string>();
      owners.add(group.id);
      claimed.set(alias, owners);
    }
  }
  const map = new Map<string, string>();
  for (const [alias, owners] of claimed) {
    if (owners.size === 1) map.set(alias, [...owners][0]);
  }
  return map;
}

/**
 * Rewrites historical source-hash exhibit IDs to current group IDs in place,
 * preserving section membership and relative order. Current IDs are left
 * unchanged. An alias that would duplicate an exhibit already present as a
 * current ID is left for reconciliation to drop.
 */
export function translateLegacyArrangementIds(
  arrangement: BundleArrangement,
  groups: readonly ExhibitGroup[],
): BundleArrangement {
  const valid = validateBundleArrangement(arrangement);
  const map = unambiguousExhibitIdMap(groups);
  const currentIds = new Set(groups.map((group) => group.id));
  const occupied = new Set(flattenBundleArrangement(valid).filter((id) => currentIds.has(id)));
  const remap = (id: string) => {
    const mapped = map.get(id);
    if (!mapped || mapped === id) return id;
    if (occupied.has(mapped)) return id;
    occupied.add(mapped);
    return mapped;
  };
  return {
    version: 1,
    nodes: valid.nodes.map((node) => {
      if (node.type === "exhibit") return { type: "exhibit" as const, exhibitId: remap(node.exhibitId) };
      return {
        type: "section" as const,
        id: node.id,
        heading: node.heading,
        exhibits: node.exhibits.map((exhibit) => ({ type: "exhibit" as const, exhibitId: remap(exhibit.exhibitId) })),
      };
    }),
  };
}

/** Translate historical aliases, then reconcile against the live group list. */
export function reconcileExhibitArrangement(
  arrangement: BundleArrangement,
  groups: readonly ExhibitGroup[],
  insertAfter: Readonly<Record<string, string>> = {},
): BundleArrangement {
  return reconcileBundleArrangement(
    translateLegacyArrangementIds(arrangement, groups),
    groups.map((group) => group.id),
    insertAfter,
  );
}

function reference(candidate: ExhibitCandidate): ExhibitReference {
  return statementReferenceFromCandidate(candidate) as ExhibitReference;
}

function pageSignature(candidate: ExhibitCandidate, evidence: EvidenceRecord) {
  if (evidence.extension !== "pdf") return "";
  const end = candidate.pageEnd ?? evidence.pageCount ?? "last";
  return `${candidate.pageStart ?? 1}-${end}`;
}

function sheetSignature(evidence: EvidenceRecord) {
  if (evidence.extension !== "xlsx") return "";
  return (evidence.sheetSelections ?? [])
    .filter((sheet) => sheet.included)
    .map((sheet) => `${sheet.name}:${sheet.range}`)
    .sort()
    .join("|");
}

function selectionSignature(candidate: ExhibitCandidate, evidence: EvidenceRecord) {
  return `${pageSignature(candidate, evidence)}::${sheetSignature(evidence)}`;
}

/**
 * Pure, review-led grouping. A matching hash is only a collision suggestion:
 * a later citation joins the canonical source only after an explicit "same"
 * decision. "separate" deliberately retains a separate exhibit group.
 */
export function deriveExhibitGroups(
  analysis: AnalysisResult,
  candidates: ExhibitCandidate[],
): ExhibitGroup[] {
  const evidenceById = new Map(analysis.evidence.map((record) => [record.id, record]));
  const selected = candidates
    .filter((candidate) => candidate.included && candidate.evidenceId)
    .map((candidate) => ({ candidate, evidence: evidenceById.get(candidate.evidenceId!) }))
    .filter((item): item is { candidate: ExhibitCandidate; evidence: EvidenceRecord } => Boolean(item.evidence))
    .sort((left, right) => sequence(left.candidate) - sequence(right.candidate));
  const byHash = new Map<string, Array<{ candidate: ExhibitCandidate; evidence: EvidenceRecord }>>();
  for (const item of selected) {
    // The hash is the only automatic repeat signal.  [AH1/...] belongs to a
    // witness's exhibit bundle and is shared by many distinct documents.
    // Scope a hash collision to that bundle so identical documents in two
    // different witnesses' bundles remain separately reviewable.
    const clusterKey = `${bundleKey(item.candidate)}:hash:${item.evidence.sha256}`;
    byHash.set(clusterKey, [...(byHash.get(clusterKey) ?? []), item]);
  }

  const groups: ExhibitGroup[] = [];
  for (const [, cluster] of byHash) {
    const sourceHash = cluster[0].evidence.sha256;
    const collision = cluster.length > 1;
    const canonicalItem = cluster[0];
    const same = collision ? cluster.filter((item, index) => index === 0 || item.candidate.repeatDecision === "same") : cluster;
    const pending = collision && cluster.slice(1).some((item) => !item.candidate.repeatDecision || item.candidate.repeatDecision === "pending");
    const signatures = new Set(same.filter((item) => item.evidence.sha256 === sourceHash).map((item) => selectionSignature(item.candidate, item.evidence)));
    groups.push({
      // The group identity belongs to the review item, not its currently
      // selected source hash. A reviewer changing the proposed document must
      // not make the card jump to the end of a persisted finalOrder.
      id: `candidate-${canonicalItem.candidate.id}`,
      canonical: canonicalItem.candidate,
      outputMark: "",
      exhibitNumber: 0,
      members: same.map((item) => item.candidate),
      collisionMembers: cluster.map((item) => item.candidate),
      evidence: canonicalItem.evidence,
      sourceHash,
      references: same.filter((item) => !item.candidate.manualAddition).map((item) => reference(item.candidate)),
      collision,
      decisionPending: pending,
      selectionConflict: signatures.size > 1,
    });
    for (const item of cluster.slice(1)) {
      if (item.candidate.repeatDecision !== "separate") continue;
      groups.push({
        id: `candidate-${item.candidate.id}`,
        canonical: item.candidate,
        outputMark: "",
        exhibitNumber: 0,
        members: [item.candidate],
        collisionMembers: [item.candidate],
        evidence: item.evidence,
        sourceHash,
        references: item.candidate.manualAddition ? [] : [reference(item.candidate)],
        collision: true,
        decisionPending: false,
        selectionConflict: false,
      });
    }
  }
  groups.sort((left, right) => sequence(left.canonical) - sequence(right.canonical));
  const nextNumber = new Map<string, number>();
  return groups.map((group) => {
    const key = bundleKey(group.canonical);
    const number = nextNumber.get(key) ?? 1;
    nextNumber.set(key, number + 1);
    // AH1 identifies the witness's exhibit bundle.  The individual exhibit
    // number is deliberately separate and is used only in the index.
    return { ...group, outputMark: bundleCitation(group.canonical), exhibitNumber: number };
  });
}

/**
 * Applies the persisted canonical group order and then recalculates the
 * per-witness-bundle index numbers.  This is used at both the UI and engine
 * boundaries so a stale legacy sequenceOrder can never override finalOrder.
 */
export function orderExhibitGroups(
  groups: ExhibitGroup[],
  canonicalOrder: string[] = [],
): ExhibitGroup[] {
  const byId = new Map(groups.map((group) => [group.id, group]));
  const aliases = unambiguousExhibitIdMap(groups);
  const ordered: ExhibitGroup[] = [];
  const retained = new Set<string>();
  for (const id of canonicalOrder) {
    const group = byId.get(aliases.get(id) ?? "");
    if (!group || retained.has(group.id)) continue;
    ordered.push(group);
    retained.add(group.id);
  }
  const reconciled = [...ordered, ...groups.filter((group) => !retained.has(group.id))];
  const nextNumber = new Map<string, number>();
  return reconciled.map((group) => {
    const key = bundleKey(group.canonical);
    const exhibitNumber = nextNumber.get(key) ?? 1;
    nextNumber.set(key, exhibitNumber + 1);
    return { ...group, outputMark: bundleCitation(group.canonical), exhibitNumber };
  });
}

export function isCanonicalCandidate(candidate: ExhibitCandidate, groups: ExhibitGroup[]) {
  return groups.some((group) => group.canonical.id === candidate.id);
}

function parentReviewCardId(
  child: ExhibitCandidate,
  candidates: ExhibitCandidate[],
  groups: ExhibitGroup[],
  evidenceById: Map<string, EvidenceRecord>,
) {
  const hash = child.parentEmailProvenance?.parentSha256;
  if (!hash) return null;
  const parentGroup = groups.find((group) => !group.canonical.parentEmailProvenance && group.evidence.sha256 === hash);
  if (parentGroup) return parentGroup.canonical.id;
  const parentCandidate = candidates.find((candidate) => {
    if (candidate.parentEmailProvenance || !candidate.evidenceId) return false;
    return evidenceById.get(candidate.evidenceId)?.sha256 === hash;
  });
  return parentCandidate?.id ?? null;
}

/**
 * Returns the review-card identities that need human attention.  Review cards
 * represent canonical groups, so an outstanding repeat reference must surface
 * its canonical card even where the repeat decision itself has been made.
 * Email-child exhibits are not review cards: their outstanding work lands on
 * the parent email. Excluded candidates are deliberately omitted.
 */
export function pendingReviewCandidateIds(
  candidates: ExhibitCandidate[],
  groups: ExhibitGroup[],
  evidence: EvidenceRecord[] = [],
) {
  const ids = new Set<string>();
  const evidenceById = new Map(evidence.map((record) => [record.id, record]));
  const groupByCandidateId = exhibitGroupLookup(groups).byCandidateId;
  for (const candidate of candidates) {
    if (!candidate.included) continue;
    if (candidate.parentEmailProvenance) {
      if (!candidate.evidenceId || !candidate.confirmed) {
        const parentId = parentReviewCardId(candidate, candidates, groups, evidenceById);
        if (parentId) ids.add(parentId);
      }
      continue;
    }
    if (!candidate.evidenceId || !candidate.confirmed) ids.add(candidate.id);
    const record = candidate.evidenceId ? evidenceById.get(candidate.evidenceId) : undefined;
    if (record?.emailAttachments?.length && unresolvedEmailAttachments(record.emailAttachments, candidate.emailAttachmentDispositions).length) {
      const group = groupByCandidateId.get(candidate.id);
      ids.add(candidate.repeatDecision === "same" && group ? group.canonical.id : candidate.id);
    }
  }
  for (const group of groups) {
    if (group.canonical.parentEmailProvenance) continue;
    const hasOutstandingMember = group.collisionMembers.some(
      (candidate) => !candidate.parentEmailProvenance && candidate.included && (!candidate.evidenceId || !candidate.confirmed),
    );
    if (group.decisionPending || group.selectionConflict || hasOutstandingMember) {
      ids.add(group.canonical.id);
    }
  }
  return ids;
}

export function emailChildInsertAfter(groups: ExhibitGroup[]) {
  const insertAfter: Record<string, string> = {};
  for (const group of groups) {
    const provenance = group.canonical.parentEmailProvenance;
    if (!provenance) continue;
    const parent = groups.find((item) => !item.canonical.parentEmailProvenance && item.evidence.sha256 === provenance.parentSha256);
    if (parent) insertAfter[group.id] = parent.id;
  }
  return insertAfter;
}
