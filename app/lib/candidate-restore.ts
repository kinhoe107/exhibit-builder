import type { ExhibitCandidate } from "./bundle-engine.ts";

/**
 * Restores reviewer choices without restoring generated statement-reference
 * identity fields. Marks, initials and sequences always come from a fresh read
 * of the source statement so legacy recovery data cannot reintroduce stale
 * references.
 */
export function restoreCitedCandidateDecision(
  current: ExhibitCandidate,
  saved: ExhibitCandidate,
  evidenceId: string | null,
): ExhibitCandidate {
  return {
    ...current,
    description: saved.description,
    aliases: saved.aliases,
    reviewNote: saved.reviewNote,
    date: saved.date,
    evidenceId,
    included: saved.included,
    confirmed: Boolean(evidenceId && saved.confirmed),
    confirmationMethod: evidenceId && saved.confirmed ? saved.confirmationMethod : undefined,
    confirmedAt: evidenceId && saved.confirmed ? saved.confirmedAt : undefined,
    pageStart: saved.pageStart,
    pageEnd: saved.pageEnd,
    sequenceOrder: saved.sequenceOrder,
    repeatDecision: saved.repeatDecision,
    emailAttachmentDispositions: saved.emailAttachmentDispositions,
  };
}
