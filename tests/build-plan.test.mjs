import assert from "node:assert/strict";
import test from "node:test";

const item = (id, pages, indexHeight = 46, optionalPages = 0) => ({
  id,
  recordIndex: Number(id.replace(/\D/g, "")) - 1,
  indexNumber: Number(id.replace(/\D/g, "")),
  witnessKey: "AH",
  initials: "AH",
  sequence: 1,
  sourceHashes: [id.repeat(64).slice(0, 64)],
  bodyStartPage: 1,
  bodyContentStartPage: 1 + optionalPages,
  bodyEndPage: pages,
  physicalPages: pages,
  optionalPages,
  contentPages: pages - optionalPages,
  indexHeight,
  references: [{ id: `${id}-ref`, relativeStart: 0, relativeEnd: pages - optionalPages - 1 }],
});

test("build plan includes cover and index overhead and never splits an exhibit", async () => {
  const { createBuildPlan } = await import("../app/lib/build-plan.ts");
  const plan = createBuildPlan([item("e1", 4), item("e2", 5), item("e3", 3)], {
    pageLimit: 10,
    coverPages: 1,
    completeIndexPages: 1,
    includeDividerPages: false,
    includeExhibitCoverPages: false,
    countOptionalPagesInReferences: false,
    matchPdfPageOrder: false,
    volumeNumbering: "restart",
    startAt: 1,
  });
  assert.deepEqual(plan.volumes.map((volume) => volume.items.map((entry) => entry.id)), [["e1"], ["e2", "e3"]]);
  assert.deepEqual(plan.volumes.map((volume) => volume.totalPages), [6, 10]);
  assert.equal(plan.bundleIdentity, "AH 1");
  assert.equal(Object.isFrozen(plan), true);
  assert.equal(Object.isFrozen(plan.volumes[0].items[0]), true);
});

test("a saved exhibit-local numbering request still uses finished-PDF page order", async () => {
  const { createBuildPlan } = await import("../app/lib/build-plan.ts");
  const base = { pageLimit: 5, coverPages: 1, completeIndexPages: 1, includeDividerPages: false, includeExhibitCoverPages: true, matchPdfPageOrder: false, volumeNumbering: "restart", startAt: 1 };
  const excluded = createBuildPlan([item("e1", 3, 46, 1), item("e2", 3, 46, 1)], { ...base, countOptionalPagesInReferences: false });
  const counted = createBuildPlan([item("e1", 3, 46, 1), item("e2", 3, 46, 1)], { ...base, countOptionalPagesInReferences: true });
  assert.deepEqual(excluded.volumes.map((volume) => volume.items[0].legalStartPage), [4, 4]);
  assert.deepEqual(counted.volumes.map((volume) => volume.items[0].legalStartPage), [4, 4]);
  assert.deepEqual(counted.volumes.map((volume) => volume.items[0].physicalContentStartPage), [4, 4]);
});

test("PDF-order references use final physical pages and follow committed item order", async () => {
  const { createBuildPlan } = await import("../app/lib/build-plan.ts");
  const options = {
    pageLimit: 0,
    coverPages: 1,
    completeIndexPages: 1,
    includeDividerPages: false,
    includeExhibitCoverPages: false,
    countOptionalPagesInReferences: false,
    matchPdfPageOrder: true,
    volumeNumbering: "continuous",
    startAt: 1,
  };
  const original = createBuildPlan([item("e1", 4), item("e2", 2)], options);
  const reordered = createBuildPlan([item("e2", 2), item("e1", 4)], options);
  assert.deepEqual(original.volumes[0].items.map((entry) => [entry.id, entry.legalStartPage, entry.legalEndPage]), [["e1", 3, 6], ["e2", 7, 8]]);
  assert.deepEqual(reordered.volumes[0].items.map((entry) => [entry.id, entry.legalStartPage, entry.legalEndPage]), [["e2", 3, 4], ["e1", 5, 8]]);
});

test("PDF-order numbering forces optional pages into its sequence but cites first source content", async () => {
  const { createBuildPlan } = await import("../app/lib/build-plan.ts");
  const plan = createBuildPlan([item("e1", 3, 46, 1)], {
    pageLimit: 0,
    coverPages: 1,
    completeIndexPages: 1,
    includeDividerPages: false,
    includeExhibitCoverPages: true,
    // A stale saved setting cannot contradict physical PDF-order behaviour.
    countOptionalPagesInReferences: false,
    matchPdfPageOrder: true,
    volumeNumbering: "continuous",
    startAt: 1,
  });
  assert.equal(plan.countOptionalPagesInReferences, true);
  assert.equal(plan.volumes[0].items[0].physicalStartPage, 3);
  assert.equal(plan.volumes[0].items[0].physicalContentStartPage, 4);
  assert.equal(plan.volumes[0].items[0].legalStartPage, 4, "the legal citation remains the first source-content page, not the divider/cover");
});

test("continuous PDF-order references continue across physical volume files", async () => {
  const { createBuildPlan } = await import("../app/lib/build-plan.ts");
  const plan = createBuildPlan([item("e1", 2), item("e2", 2)], {
    pageLimit: 4,
    coverPages: 1,
    completeIndexPages: 1,
    includeDividerPages: false,
    includeExhibitCoverPages: false,
    countOptionalPagesInReferences: false,
    matchPdfPageOrder: true,
    volumeNumbering: "continuous",
    startAt: 1,
  });
  assert.deepEqual(plan.volumes.map((volume) => volume.items[0].legalStartPage), [3, 7]);
});

test("repeats the exact complete-index overhead in every volume and validates section order", async () => {
  const { createBuildPlan } = await import("../app/lib/build-plan.ts");
  const plan = createBuildPlan([item("e1", 2), item("e2", 2), item("e3", 2)], {
    pageLimit: 8,
    coverPages: 1,
    completeIndexPages: 3,
    indexNodes: [
      { kind: "section", id: "agreements", title: "Agreements", itemIds: ["e1", "e2"] },
      { kind: "exhibit", itemId: "e3" },
    ],
    includeDividerPages: false,
    includeExhibitCoverPages: false,
    countOptionalPagesInReferences: false,
    matchPdfPageOrder: true,
    volumeNumbering: "continuous",
    startAt: 1,
  });
  assert.deepEqual(plan.volumes.map((volume) => volume.items.map((entry) => entry.id)), [["e1", "e2"], ["e3"]]);
  assert.deepEqual(plan.volumes.map((volume) => [volume.coverPages, volume.indexPages, volume.exhibitPages, volume.totalPages]), [[1, 3, 4, 8], [1, 3, 2, 6]]);
  assert.deepEqual(plan.indexNodes[0], { kind: "section", id: "agreements", title: "Agreements", itemIds: ["e1", "e2"] });
  await assert.rejects(async () => createBuildPlan([item("e1", 1), item("e2", 1)], {
    pageLimit: 0,
    coverPages: 1,
    completeIndexPages: 1,
    indexNodes: [{ kind: "section", id: "bad", title: "Bad order", itemIds: ["e2", "e1"] }],
    includeDividerPages: false,
    includeExhibitCoverPages: false,
    countOptionalPagesInReferences: false,
    matchPdfPageOrder: true,
    volumeNumbering: "continuous",
    startAt: 1,
  }), /every exhibit exactly once in canonical order/i);
});
