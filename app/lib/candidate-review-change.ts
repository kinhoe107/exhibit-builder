import type { ExhibitCandidate } from "./bundle-engine.ts";

const CONFIRMATION_SENSITIVE_KEYS = [
  "included",
  "evidenceId",
  "provisionalNumber",
  "mark",
  "pageStart",
  "pageEnd",
] as const;

function unconfirmed<T extends Pick<ExhibitCandidate, "confirmed" | "confirmationMethod" | "confirmedAt">>(candidate: T): T {
  return {
    ...candidate,
    confirmed: false,
    confirmationMethod: undefined,
    confirmedAt: undefined,
  };
}

/**
 * Applies a Review-card edit. Choosing a document is a selection, never an
 * approval. An evidence change therefore revokes confirmation on every
 * citation that shares the old or new source hash, even when the picker also
 * sends `confirmed: false`. Index description and document date are labels on
 * the confirmed match; they must not undo that confirmation or a sibling's.
 */
export function applyCandidateReviewChange(
  candidates: ExhibitCandidate[],
  candidateId: string,
  change: Partial<ExhibitCandidate>,
  evidenceById: Map<string, { sha256: string }>,
): ExhibitCandidate[] {
  const before = candidates.find((candidate) => candidate.id === candidateId);
  if (!before) return candidates;

  if ("repeatDecision" in change && !("confirmed" in change) && !("evidenceId" in change)) {
    if (before.repeatDecision === change.repeatDecision) return candidates;
    return candidates.map((candidate) => {
      if (candidate.id !== candidateId) return candidate;
      const next = { ...candidate, ...change };
      return candidate.confirmed ? unconfirmed(next) : next;
    });
  }

  const sourceFieldsChanged = CONFIRMATION_SENSITIVE_KEYS.some((key) => key in change);
  const evidenceChanged = "evidenceId" in change && change.evidenceId !== before.evidenceId;
  const nextChange = sourceFieldsChanged && !("confirmed" in change)
    ? unconfirmed({ ...before, ...change })
    : change;
  const changed = { ...before, ...nextChange };

  if (!sourceFieldsChanged || !evidenceChanged) {
    return candidates.map((candidate) => candidate.id === candidateId ? changed : candidate);
  }

  const affectedHashes = new Set([before.evidenceId, changed.evidenceId]
    .filter((id): id is string => Boolean(id))
    .map((id) => evidenceById.get(id)?.sha256)
    .filter((hash): hash is string => Boolean(hash)));

  return candidates.map((candidate) => {
    const next = candidate.id === candidateId ? changed : candidate;
    const sourceHash = next.evidenceId ? evidenceById.get(next.evidenceId)?.sha256 : undefined;
    return sourceHash && affectedHashes.has(sourceHash) ? unconfirmed(next) : next;
  });
}
