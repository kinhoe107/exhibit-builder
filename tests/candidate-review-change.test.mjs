import assert from "node:assert/strict";
import test from "node:test";
import { applyCandidateReviewChange } from "../app/lib/candidate-review-change.ts";
import { exhibitGroupLookup, formatRepeatExhibitNote, repeatExhibitCount, reviewCandidatesForDisplay } from "../app/lib/exhibit-groups.ts";

function candidate(id, extras = {}) {
  return {
    id,
    included: true,
    confirmed: extras.confirmed ?? false,
    confirmationMethod: extras.confirmed ? "individual" : undefined,
    confirmedAt: extras.confirmed ? "2026-08-17T00:00:00.000Z" : undefined,
    evidenceId: extras.evidenceId ?? null,
    repeatDecision: extras.repeatDecision,
    sequenceOrder: extras.sequenceOrder ?? 1,
    provisionalNumber: extras.sequenceOrder ?? 1,
    parentEmailProvenance: extras.parentEmailProvenance,
  };
}

test("selecting a source already used by another citation revokes every affected confirmation", () => {
  const evidenceById = new Map([
    ["invoice", { sha256: "invoice-hash" }],
    ["ledger", { sha256: "ledger-hash" }],
  ]);
  const first = candidate("first", { evidenceId: "invoice", confirmed: true, sequenceOrder: 1 });
  const second = candidate("second", { evidenceId: "ledger", confirmed: false, sequenceOrder: 2 });
  const oldPeer = candidate("old-peer", { evidenceId: "ledger", confirmed: true, sequenceOrder: 3 });
  const next = applyCandidateReviewChange(
    [first, second, oldPeer],
    "second",
    { evidenceId: "invoice", confidence: 100, rationale: "Source selected by reviewer", repeatDecision: "pending", confirmed: false, confirmationMethod: undefined, confirmedAt: undefined },
    evidenceById,
  );
  assert.equal(next.find((item) => item.id === "second")?.confirmed, false);
  assert.equal(next.find((item) => item.id === "first")?.confirmed, false);
  assert.equal(next.find((item) => item.id === "old-peer")?.confirmed, false);
  assert.equal(next.find((item) => item.id === "second")?.repeatDecision, "pending");
  assert.equal(next.find((item) => item.id === "first")?.confirmationMethod, undefined);
});

test("clearing a selected source revokes confirmation across its old source hash", () => {
  const evidenceById = new Map([["invoice", { sha256: "invoice-hash" }]]);
  const first = candidate("first", { evidenceId: "invoice", confirmed: true, sequenceOrder: 1 });
  const second = candidate("second", { evidenceId: "invoice", confirmed: true, sequenceOrder: 2 });
  const next = applyCandidateReviewChange(
    [first, second],
    "second",
    { evidenceId: null, repeatDecision: "pending", confirmed: false, confirmationMethod: undefined, confirmedAt: undefined },
    evidenceById,
  );
  assert.equal(next.find((item) => item.id === "first")?.confirmed, false);
  assert.equal(next.find((item) => item.id === "second")?.confirmed, false);
  assert.equal(next.find((item) => item.id === "second")?.evidenceId, null);
});

test("index description and date edits keep document confirmation on this card and its source siblings", () => {
  const evidenceById = new Map([["invoice", { sha256: "invoice-hash" }]]);
  const first = { ...candidate("first", { evidenceId: "invoice", confirmed: true, sequenceOrder: 1 }), description: "Invoice", date: "1 March 2026" };
  const second = { ...candidate("second", { evidenceId: "invoice", confirmed: true, sequenceOrder: 2, repeatDecision: "separate" }), description: "Same invoice, separate exhibit", date: "1 March 2026" };
  const renamed = applyCandidateReviewChange([first, second], "first", { description: "Paid invoice" }, evidenceById);
  assert.equal(renamed.find((item) => item.id === "first")?.description, "Paid invoice");
  assert.equal(renamed.find((item) => item.id === "first")?.confirmed, true);
  assert.equal(renamed.find((item) => item.id === "second")?.confirmed, true);
  const redated = applyCandidateReviewChange(renamed, "second", { date: "2 March 2026" }, evidenceById);
  assert.equal(redated.find((item) => item.id === "second")?.date, "2 March 2026");
  assert.equal(redated.find((item) => item.id === "first")?.confirmed, true);
  assert.equal(redated.find((item) => item.id === "second")?.confirmed, true);
});

test("changing the selected source pages revokes only that card's confirmation", () => {
  const evidenceById = new Map([["invoice", { sha256: "invoice-hash" }]]);
  const first = candidate("first", { evidenceId: "invoice", confirmed: true, sequenceOrder: 1 });
  const second = candidate("second", { evidenceId: "invoice", confirmed: true, sequenceOrder: 2, repeatDecision: "separate" });
  const next = applyCandidateReviewChange([first, second], "first", { pageStart: 2, pageEnd: 4 }, evidenceById);
  assert.equal(next.find((item) => item.id === "first")?.confirmed, false);
  assert.equal(next.find((item) => item.id === "first")?.pageStart, 2);
  assert.equal(next.find((item) => item.id === "first")?.pageEnd, 4);
  assert.equal(next.find((item) => item.id === "second")?.confirmed, true);
});

test("an explicit Confirm this document action confirms only that card", () => {
  const evidenceById = new Map([["invoice", { sha256: "invoice-hash" }]]);
  const first = candidate("first", { evidenceId: "invoice", confirmed: false, sequenceOrder: 1 });
  const second = candidate("second", { evidenceId: "invoice", confirmed: false, repeatDecision: "pending", sequenceOrder: 2 });
  const next = applyCandidateReviewChange(
    [first, second],
    "first",
    { confirmed: true, confirmationMethod: "individual", confirmedAt: "2026-08-17T12:00:00.000Z" },
    evidenceById,
  );
  assert.equal(next.find((item) => item.id === "first")?.confirmed, true);
  assert.equal(next.find((item) => item.id === "second")?.confirmed, false);
});

test("Same or Separate does not confirm a document and already-selected Same is a no-op", () => {
  const evidenceById = new Map([["invoice", { sha256: "invoice-hash" }]]);
  const first = candidate("first", { evidenceId: "invoice", confirmed: true, sequenceOrder: 1 });
  const second = candidate("second", { evidenceId: "invoice", confirmed: false, repeatDecision: "pending", sequenceOrder: 2 });
  const selected = applyCandidateReviewChange([first, second], "second", { repeatDecision: "same" }, evidenceById);
  assert.equal(selected.find((item) => item.id === "second")?.repeatDecision, "same");
  assert.equal(selected.find((item) => item.id === "second")?.confirmed, false);
  assert.equal(selected.find((item) => item.id === "first")?.confirmed, true);
  const unchanged = applyCandidateReviewChange(selected, "second", { repeatDecision: "same" }, evidenceById);
  assert.equal(unchanged, selected);
  const separate = applyCandidateReviewChange([first, second], "second", { repeatDecision: "separate" }, evidenceById);
  assert.equal(separate.find((item) => item.id === "second")?.repeatDecision, "separate");
  assert.equal(separate.find((item) => item.id === "second")?.confirmed, false);
  assert.equal(separate.find((item) => item.id === "first")?.confirmed, true);
});

test("a pending later citation keeps its own review card until Same is chosen", () => {
  const canonical = candidate("first", { evidenceId: "invoice", confirmed: true, sequenceOrder: 1 });
  const pending = candidate("second", { evidenceId: "invoice", confirmed: false, repeatDecision: "pending", sequenceOrder: 2 });
  const unconfirmedSame = candidate("third", { evidenceId: "invoice", confirmed: false, repeatDecision: "same", sequenceOrder: 3 });
  const confirmedSame = candidate("fourth", { evidenceId: "invoice", confirmed: true, repeatDecision: "same", sequenceOrder: 4 });
  const groups = [{
    id: "candidate-first",
    canonical,
    collisionMembers: [canonical, pending, unconfirmedSame, confirmedSame],
    members: [canonical, unconfirmedSame, confirmedSame],
  }];
  assert.deepEqual(
    reviewCandidatesForDisplay([canonical, pending, unconfirmedSame, confirmedSame], groups).map((item) => item.id),
    ["first", "second"],
  );
  const lookup = exhibitGroupLookup(groups);
  assert.equal(lookup.byCandidateId.get(canonical.id), groups[0]);
  assert.equal(lookup.byCandidateId.get(pending.id), groups[0]);
});

test("repeat progress names one shared source once", () => {
  const groups = [
    { canonical: candidate("a"), collisionMembers: [candidate("a"), candidate("b")] },
    { canonical: candidate("c", { parentEmailProvenance: { parentName: "mail" } }), collisionMembers: [candidate("c"), candidate("d")] },
    { canonical: candidate("e"), collisionMembers: [candidate("e")] },
  ];
  assert.equal(repeatExhibitCount(groups), 1);
  assert.equal(formatRepeatExhibitNote(1), " (1 repeat exhibit)");
  assert.equal(formatRepeatExhibitNote(2), " (2 repeat exhibits)");
  assert.equal(formatRepeatExhibitNote(0), "");
});
