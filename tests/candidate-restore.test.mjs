import assert from "node:assert/strict";
import test from "node:test";
import { restoreCitedCandidateDecision } from "../app/lib/candidate-restore.ts";

test("project restore keeps reviewer choices but not stale generated reference identity", () => {
  const current = {
    id: "statement-1:candidate-1", mark: "LV 1", provisionalNumber: 1,
    description: "Fresh description", date: "Date not stated", paragraph: 20,
    citation: "[LV-xx]", exhibitInitials: "LV", exhibitSequence: 1,
    citationResolution: "unresolved", discoverySignals: [], evidenceId: "fresh-source",
    confidence: 90, rationale: "Fresh analysis", included: true, confirmed: false,
    witnessInitials: "LV", witnessKey: "lucia::LV", statementId: "statement-1",
  };
  const saved = {
    ...current, mark: "VA 1", exhibitInitials: "VA", witnessInitials: "VA",
    witnessKey: "lucia::VA", description: "Reviewed description", evidenceId: "old-source",
    confirmed: true, confirmationMethod: "individual", confirmedAt: "2026-08-01T12:00:00.000Z",
  };
  const restored = restoreCitedCandidateDecision(current, saved, "fresh-source");
  assert.equal(restored.mark, "LV 1");
  assert.equal(restored.exhibitInitials, "LV");
  assert.equal(restored.witnessInitials, "LV");
  assert.equal(restored.witnessKey, "lucia::LV");
  assert.equal(restored.description, "Reviewed description");
  assert.equal(restored.evidenceId, "fresh-source");
  assert.equal(restored.confirmed, true);
  assert.equal(restored.emailAttachmentDispositions, undefined);
});

test("project restore keeps hash-bound email attachment dispositions", () => {
  const current = {
    id: "statement-1:candidate-1", mark: "LV 1", provisionalNumber: 1,
    description: "Email", date: "Date not stated", paragraph: 4,
    citation: "I refer to the email [LV-xx].", exhibitInitials: "LV", exhibitSequence: 1,
    citationResolution: "none", discoverySignals: [], evidenceId: "fresh-source",
    confidence: 90, rationale: "Fresh analysis", included: true, confirmed: false,
    witnessInitials: "LV", witnessKey: "lucia::LV", statementId: "statement-1",
  };
  const saved = {
    ...current, confirmed: true, confirmationMethod: "individual",
    confirmedAt: "2026-08-01T12:00:00.000Z",
    emailAttachmentDispositions: { "abc:1:def": "print-with-email" },
  };
  const restored = restoreCitedCandidateDecision(current, saved, "fresh-source");
  assert.deepEqual(restored.emailAttachmentDispositions, { "abc:1:def": "print-with-email" });
});

test("cited restore keeps a hash-bound same/separate decision and clears it when the statement binding fails", () => {
  const current = {
    id: "statement-1:candidate-2", mark: "LV 2", provisionalNumber: 2,
    description: "Fresh description", date: "Date not stated", paragraph: 21,
    citation: "[LV-xx]", exhibitInitials: "LV", exhibitSequence: 1,
    citationResolution: "unresolved", discoverySignals: [], evidenceId: "fresh-source",
    confidence: 90, rationale: "Fresh analysis", included: true, confirmed: false,
    witnessInitials: "LV", witnessKey: "lucia::LV", statementId: "statement-1",
  };
  const saved = {
    ...current, description: "Reviewed description", evidenceId: "old-source",
    confirmed: true, confirmationMethod: "individual", confirmedAt: "2026-08-01T12:00:00.000Z",
    repeatDecision: "same",
  };
  const preserved = restoreCitedCandidateDecision(current, saved, "fresh-source", true);
  assert.equal(preserved.confirmed, true);
  assert.equal(preserved.repeatDecision, "same");
  const staleStatement = restoreCitedCandidateDecision(current, saved, "fresh-source", false);
  assert.equal(staleStatement.confirmed, false);
  assert.equal(staleStatement.repeatDecision, undefined);
});
