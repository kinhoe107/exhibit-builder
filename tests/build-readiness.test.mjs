import assert from "node:assert/strict";
import test from "node:test";

test("explains an OCR blocker when all exhibit approvals are complete", async () => {
  const { buildBlockers } = await import("../app/lib/build-readiness.ts");
  const blockers = buildBlockers({
    includedCount: 11,
    confirmedCount: 11,
    pendingApprovalCount: 0,
    templateReviewPending: false,
    preflight: [{ id: "ocr-1", severity: "blocking", label: "OCR unavailable", detail: "This PDF needs local OCR, which is unavailable in this environment.", fileName: "Scanned.pdf" }],
  });
  assert.deepEqual(blockers.map((blocker) => blocker.label), ["OCR unavailable"]);
  assert.match(blockers[0].detail, /local OCR/i);
  assert.equal(blockers[0].fileName, "Scanned.pdf");
});

test("separates outstanding approvals from technical build checks", async () => {
  const { buildBlockers } = await import("../app/lib/build-readiness.ts");
  const blockers = buildBlockers({
    includedCount: 3,
    confirmedCount: 2,
    pendingApprovalCount: 1,
    templateReviewPending: false,
    preflight: [],
  });
  assert.equal(blockers.length, 1);
  assert.equal(blockers[0].kind, "approval");
  assert.match(blockers[0].detail, /1 included exhibit/);
});

test("empty inclusion asks for an exhibit rather than a cited match", async () => {
  const { buildBlockers } = await import("../app/lib/build-readiness.ts");
  const blockers = buildBlockers({
    includedCount: 0,
    confirmedCount: 0,
    pendingApprovalCount: 0,
    templateReviewPending: false,
    preflight: [],
  });
  assert.equal(blockers[0]?.id, "no-included-exhibits");
  assert.match(blockers[0]?.detail ?? "", /at least one exhibit/i);
  assert.doesNotMatch(blockers[0]?.detail ?? "", /cited exhibit/i);
});

test("returns no blocker when confirmations, templates and preflight are clear", async () => {
  const { buildBlockers } = await import("../app/lib/build-readiness.ts");
  assert.deepEqual(buildBlockers({ includedCount: 1, confirmedCount: 1, pendingApprovalCount: 0, templateReviewPending: false, preflight: [{ id: "ready", severity: "pass", label: "Ready to build", detail: "No blocking issue." }] }), []);
});
