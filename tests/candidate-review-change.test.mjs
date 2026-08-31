import assert from "node:assert/strict";
import test from "node:test";
import { applyCandidateReviewChange } from "../app/lib/candidate-review-change.ts";
import { exhibitGroupLookup, formatRepeatExhibitNote, repeatExhibitCount, reviewCandidatesForDisplay, bulkConfirmableCandidates, AUTOMATIC_MATCH_REVIEW_THRESHOLD } from "../app/lib/exhibit-groups.ts";

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

test("review cards remain in witness-statement sequence even when some sources are unmatched or groups are reordered", () => {
  const first = candidate("first", { evidenceId: "contract", sequenceOrder: 1 });
  const unmatched = candidate("unmatched", { evidenceId: null, sequenceOrder: 2 });
  const third = candidate("third", { evidenceId: "email", sequenceOrder: 3 });
  const repeat = candidate("repeat", { evidenceId: "contract", repeatDecision: "pending", sequenceOrder: 4 });
  const groups = [
    { id: "candidate-third", canonical: third, collisionMembers: [third], members: [third] },
    { id: "candidate-first", canonical: first, collisionMembers: [first, repeat], members: [first] },
  ];
  assert.deepEqual(
    reviewCandidatesForDisplay([third, repeat, first, unmatched], groups).map((item) => item.id),
    ["first", "unmatched", "third", "repeat"],
  );
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

test("bulk confirmation excludes an automatic score of 64 and includes 65", () => {
  function groupFor(item, extras = {}) {
    return {
      id: `candidate-${item.id}`,
      canonical: item,
      collisionMembers: extras.collisionMembers ?? [item],
      members: extras.members ?? [item],
      decisionPending: extras.decisionPending ?? false,
      selectionConflict: extras.selectionConflict ?? false,
    };
  }
  const weak = { ...candidate("weak", { evidenceId: "a" }), confidence: AUTOMATIC_MATCH_REVIEW_THRESHOLD - 1, rationale: "Filename tokens matched" };
  const boundary = { ...candidate("boundary", { evidenceId: "b" }), confidence: AUTOMATIC_MATCH_REVIEW_THRESHOLD, rationale: "Filename tokens matched" };
  const strong = { ...candidate("strong", { evidenceId: "c" }), confidence: 90, rationale: "Filename tokens matched" };
  const manual = { ...candidate("manual", { evidenceId: "d" }), confidence: 100, rationale: "Source selected by reviewer" };
  const unmatched = { ...candidate("unmatched"), confidence: 80, rationale: "Filename tokens matched" };
  const confirmed = { ...candidate("confirmed", { evidenceId: "e", confirmed: true }), confidence: 80, rationale: "Filename tokens matched" };
  const child = { ...candidate("child", { evidenceId: "f", parentEmailProvenance: { parentName: "mail" } }), confidence: 80, rationale: "Filename tokens matched" };
  const canonical = { ...candidate("canonical", { evidenceId: "g" }), confidence: 80, rationale: "Filename tokens matched" };
  const repeat = { ...candidate("repeat", { evidenceId: "g", sequenceOrder: 2, repeatDecision: "same" }), confidence: 80, rationale: "Filename tokens matched" };
  const conflicted = { ...candidate("conflicted", { evidenceId: "h" }), confidence: 80, rationale: "Filename tokens matched" };
  const ids = bulkConfirmableCandidates(
    [weak, boundary, strong, manual, unmatched, confirmed, child, canonical, repeat, conflicted],
    [
      groupFor(weak),
      groupFor(boundary),
      groupFor(strong),
      groupFor(manual),
      groupFor(unmatched),
      groupFor(confirmed),
      groupFor(canonical, { collisionMembers: [canonical, repeat], members: [canonical, repeat] }),
      groupFor(conflicted, { selectionConflict: true }),
    ],
  ).map((item) => item.id);
  assert.deepEqual(ids, ["boundary", "strong", "manual", "canonical"]);
  assert.equal(ids.includes("weak"), false);
  assert.equal(AUTOMATIC_MATCH_REVIEW_THRESHOLD, 65);
});
