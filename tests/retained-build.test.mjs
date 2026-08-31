import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

function analysisFixture(overrides = {}) {
  return {
    statementHash: "s".repeat(64),
    evidence: [{ id: "e1", sha256: "e".repeat(64) }],
    ...overrides,
  };
}

function candidateFixture(overrides = {}) {
  return {
    id: "c1",
    mark: "AH 1",
    provisionalNumber: 1,
    description: "Exhibit",
    date: "Date not stated",
    paragraph: 2,
    citation: "[AH1/xx]",
    discoverySignals: [],
    evidenceId: "e1",
    confidence: 100,
    rationale: "test",
    included: true,
    confirmed: true,
    ...overrides,
  };
}

function ocrResolution() {
  return {
    blockerId: "ocr-e1",
    action: "proceed-without-ocr",
    sourceId: "e1",
    sourceSha256: "e".repeat(64),
    note: "Approved without OCR",
    visualReviewConfirmed: true,
  };
}

async function baseInputs(overrides = {}) {
  const { retainedBuildInputsFrom } = await import("../app/lib/retained-build.ts");
  const { DEFAULT_PAGINATION, DEFAULT_BUNDLE_LAYOUT } = await import("../app/lib/bundle-types.ts");
  return retainedBuildInputsFrom({
    analysis: analysisFixture(),
    candidates: [candidateFixture()],
    templates: [],
    layout: DEFAULT_BUNDLE_LAYOUT,
    pagination: DEFAULT_PAGINATION,
    pageSizeChoices: { e1: "convert-to-a4" },
    resolutions: [ocrResolution()],
    templateDiscrepancyConfirmation: null,
    ...overrides,
  });
}

test("keep current bundle still restores after arrangement-only edits", async () => {
  const { captureRetainedBundle, dropStaleRetainedBuild, restoredBundleFromRetain } = await import("../app/lib/retained-build.ts");
  const inputs = await baseInputs();
  const retained = captureRetainedBundle({
    build: { id: "pdf-v1" },
    arrangement: { nodes: ["original-order"] },
    candidates: [candidateFixture()],
    pageSizeChoices: { e1: "convert-to-a4" },
    inputs,
  });
  const afterOrderEdit = dropStaleRetainedBuild(retained, inputs);
  const restored = restoredBundleFromRetain(afterOrderEdit);
  assert.equal(restored.build?.id, "pdf-v1");
  assert.deepEqual(restored.arrangement, { nodes: ["original-order"] });
  assert.equal(restored.pageSizeChoices.e1, "convert-to-a4");
});

test("cross-view numbering, template and layout changes drop the retained PDF", async () => {
  const { captureRetainedBundle, dropStaleRetainedBuild, restoredBundleFromRetain } = await import("../app/lib/retained-build.ts");
  const { DEFAULT_PAGINATION, DEFAULT_BUNDLE_LAYOUT } = await import("../app/lib/bundle-types.ts");
  const builtInputs = await baseInputs();
  const retained = captureRetainedBundle({
    build: { id: "pdf-v1" },
    arrangement: { nodes: ["original-order"] },
    candidates: [candidateFixture()],
    pageSizeChoices: { e1: "convert-to-a4" },
    inputs: builtInputs,
  });

  const afterPrefix = dropStaleRetainedBuild(retained, await baseInputs({ pagination: { ...DEFAULT_PAGINATION, prefix: "AH-" } }));
  const prefixRestore = restoredBundleFromRetain(afterPrefix);
  assert.equal(prefixRestore.build, null, "applied prefix must not restore the pre-edit PDF");
  assert.deepEqual(prefixRestore.arrangement, { nodes: ["original-order"] }, "arrangement-only discard remains available");

  assert.equal(restoredBundleFromRetain(dropStaleRetainedBuild(retained, await baseInputs({ layout: { ...DEFAULT_BUNDLE_LAYOUT, includeDividerPages: true } }))).build, null);
  assert.equal(restoredBundleFromRetain(dropStaleRetainedBuild(retained, await baseInputs({
    templates: [{ slot: "cover", sha256: "a".repeat(64), pdfSha256: "b".repeat(64) }],
  }))).build, null);
  assert.equal(restoredBundleFromRetain(dropStaleRetainedBuild(retained, await baseInputs({
    templateDiscrepancyConfirmation: { fingerprint: "fp", confirmedAt: "2026-08-25T09:00:00.000Z" },
  }))).build, null);
});

test("OCR undo, candidate change and page-size change drop the retained PDF", async () => {
  const { captureRetainedBundle, dropStaleRetainedBuild, restoredBundleFromRetain } = await import("../app/lib/retained-build.ts");
  const builtInputs = await baseInputs();
  const retained = captureRetainedBundle({
    build: { id: "pdf-v1" },
    arrangement: { nodes: ["original-order"] },
    candidates: [candidateFixture()],
    pageSizeChoices: { e1: "convert-to-a4" },
    inputs: builtInputs,
  });

  const afterOcrUndo = dropStaleRetainedBuild(retained, await baseInputs({ resolutions: [] }));
  const ocrRestore = restoredBundleFromRetain(afterOcrUndo);
  assert.equal(ocrRestore.build, null, "undoing OCR approval must not restore the pre-edit PDF");
  assert.deepEqual(ocrRestore.arrangement, { nodes: ["original-order"] });

  const afterCandidate = dropStaleRetainedBuild(retained, await baseInputs({
    candidates: [candidateFixture({ description: "Amended exhibit title" })],
  }));
  assert.equal(restoredBundleFromRetain(afterCandidate).build, null, "an analysis/candidate-derived change must drop the retained PDF");

  const afterPageSize = dropStaleRetainedBuild(retained, await baseInputs({
    pageSizeChoices: { e1: "keep-original" },
  }));
  assert.equal(restoredBundleFromRetain(afterPageSize).build, null, "a page-size change must drop the retained PDF");
});

test("discard after a dropped PDF keeps live candidate and page-size state", async () => {
  const { captureRetainedBundle, dropStaleRetainedBuild, restoredBundleFromRetain } = await import("../app/lib/retained-build.ts");
  const { flattenBundleArrangement } = await import("../app/lib/bundle-arrangement.ts");
  const original = candidateFixture({ id: "c1", description: "Original title" });
  const second = candidateFixture({ id: "c2", evidenceId: "e2", description: "Second" });
  const capturedArrangement = {
    version: 1,
    nodes: [
      { type: "exhibit", exhibitId: "c1" },
      { type: "exhibit", exhibitId: "c2" },
    ],
  };
  const retained = captureRetainedBundle({
    build: { id: "pdf-v1" },
    arrangement: capturedArrangement,
    candidates: [original, second],
    pageSizeChoices: { e1: "convert-to-a4" },
    inputs: await baseInputs({ candidates: [original, second], pageSizeChoices: { e1: "convert-to-a4" } }),
  });
  const liveCandidates = [candidateFixture({ id: "c1", description: "Amended exhibit title" }), second];
  const livePageSize = { e1: "keep-original" };
  const afterReview = dropStaleRetainedBuild(retained, await baseInputs({
    candidates: liveCandidates,
    pageSizeChoices: livePageSize,
  }));
  const restored = restoredBundleFromRetain(afterReview, {
    candidates: liveCandidates,
    pageSizeChoices: livePageSize,
    exhibitIds: ["c1", "c2"],
  });
  assert.equal(restored.build, null);
  assert.equal(restored.candidates[0].description, "Amended exhibit title", "discard must not roll back the live review change");
  assert.equal(restored.pageSizeChoices.e1, "keep-original", "discard must not roll back the live page-size choice");
  assert.deepEqual(flattenBundleArrangement(restored.arrangement), ["c1", "c2"], "captured arrangement still reverts");
});

test("email-child disposition change drops the retained PDF and discard keeps live review state", async () => {
  const { captureRetainedBundle, dropStaleRetainedBuild, restoredBundleFromRetain } = await import("../app/lib/retained-build.ts");
  const childIdentity = `${"p".repeat(64)}:1:${"c".repeat(64)}`;
  const analysis = analysisFixture({
    evidence: [{ id: "e1", sha256: "e".repeat(64), emailAttachments: [{ identity: childIdentity }] }],
  });
  const printed = candidateFixture({ emailAttachmentDispositions: { [childIdentity]: "print-with-email" } });
  const leftOut = candidateFixture({
    description: "Amended after leave-out",
    emailAttachmentDispositions: { [childIdentity]: "leave-out" },
  });
  const added = candidateFixture({ emailAttachmentDispositions: { [childIdentity]: "add-as-exhibit" } });
  const capturedArrangement = { version: 1, nodes: [{ type: "exhibit", exhibitId: "c1" }] };
  const retained = captureRetainedBundle({
    build: { id: "pdf-with-child" },
    arrangement: capturedArrangement,
    candidates: [printed],
    pageSizeChoices: { e1: "convert-to-a4" },
    inputs: await baseInputs({ analysis, candidates: [printed] }),
  });
  const afterLeaveOut = dropStaleRetainedBuild(retained, await baseInputs({ analysis, candidates: [leftOut], pageSizeChoices: { e1: "keep-original" } }));
  const restored = restoredBundleFromRetain(afterLeaveOut, {
    candidates: [leftOut],
    pageSizeChoices: { e1: "keep-original" },
    exhibitIds: ["c1"],
  });
  assert.equal(restored.build, null, "leave-out must drop the print-with-email PDF");
  assert.equal(restored.candidates[0].emailAttachmentDispositions[childIdentity], "leave-out");
  assert.equal(restored.candidates[0].description, "Amended after leave-out");
  assert.equal(restored.pageSizeChoices.e1, "keep-original");
  assert.equal(restoredBundleFromRetain(dropStaleRetainedBuild(retained, await baseInputs({ analysis, candidates: [added] }))).build, null);
});

test("open or cancel order preview keeps the retained PDF; genuine blockers and substantive edits still drop it", async () => {
  const {
    captureRetainedBundle,
    dropStaleRetainedBuild,
    restoredBundleFromRetain,
    retainBuildReadiness,
  } = await import("../app/lib/retained-build.ts");
  const { DEFAULT_PAGINATION } = await import("../app/lib/bundle-types.ts");
  const inputs = await baseInputs();
  const retained = captureRetainedBundle({
    build: { id: "pdf-v1" },
    arrangement: { nodes: ["original-order"] },
    candidates: [candidateFixture()],
    pageSizeChoices: { e1: "convert-to-a4" },
    inputs,
  });
  assert.equal(retainBuildReadiness([]), true);
  assert.equal(retainBuildReadiness([{ id: "order-preview-pending" }]), true);
  assert.equal(retainBuildReadiness([{ id: "order-preview-pending" }, { id: "exhibit-approvals" }]), false);
  assert.equal(retainBuildReadiness([{ id: "exhibit-approvals" }]), false);

  const duringPreview = dropStaleRetainedBuild(retained, inputs, {
    readyToBuild: retainBuildReadiness([{ id: "order-preview-pending" }]),
  });
  const afterCancel = dropStaleRetainedBuild(duringPreview, inputs, { readyToBuild: retainBuildReadiness([]) });
  assert.equal(duringPreview?.build?.id, "pdf-v1");
  assert.equal(afterCancel?.build?.id, "pdf-v1");
  assert.equal(restoredBundleFromRetain(afterCancel).build?.id, "pdf-v1");

  const afterPrefix = dropStaleRetainedBuild(retained, await baseInputs({ pagination: { ...DEFAULT_PAGINATION, prefix: "AH-" } }), {
    readyToBuild: retainBuildReadiness([{ id: "order-preview-pending" }]),
  });
  assert.equal(restoredBundleFromRetain(afterPrefix).build, null, "substantive change must still drop the PDF during preview");

  const blocked = dropStaleRetainedBuild(retained, inputs, {
    readyToBuild: retainBuildReadiness([{ id: "exhibit-approvals" }]),
  });
  assert.equal(restoredBundleFromRetain(blocked).build, null, "a non-preview readiness blocker must still drop the PDF");
});

test("order preview still disables generate while retain ignores that transient blocker", async () => {
  const source = await readFile(new URL("../app/BundleBuilder.tsx", import.meta.url), "utf8");
  assert.match(source, /id: "order-preview-pending"/);
  assert.match(source, /const readyToBuild = buildBlockerList\.length === 0;/);
  assert.match(source, /const retainReadyToBuild = retainBuildReadiness\(buildBlockerList\);/);
  assert.match(source, /dropStaleRetainedBuild\(reorderReturn, retainedBuildInputs, \{ readyToBuild: retainReadyToBuild \}\)/);
  assert.match(source, /dropStaleRetainedBuild\(current, retainedBuildInputs, \{ readyToBuild: retainReadyToBuild \}\)/);
  assert.match(source, /aria-disabled=\{!readyToBuild\}/);
});

test("keep current bundle does not restore a PDF when build readiness is blocking", async () => {
  const { captureRetainedBundle, dropStaleRetainedBuild, restoredBundleFromRetain } = await import("../app/lib/retained-build.ts");
  const inputs = await baseInputs();
  const retained = captureRetainedBundle({
    build: { id: "pdf-v1" },
    arrangement: { nodes: ["original-order"] },
    candidates: [candidateFixture()],
    pageSizeChoices: { e1: "convert-to-a4" },
    inputs,
  });
  const blocked = dropStaleRetainedBuild(retained, inputs, { readyToBuild: false });
  assert.equal(restoredBundleFromRetain(blocked).build, null);
  assert.deepEqual(restoredBundleFromRetain(blocked).arrangement, { nodes: ["original-order"] });
});
