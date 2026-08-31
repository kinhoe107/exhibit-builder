import assert from "node:assert/strict";
import test from "node:test";

const arrangementModule = import("../app/lib/bundle-arrangement.ts");

test("legacy order migrates to a validated flat schema-1 arrangement", async () => {
  const { bundleArrangementFromLegacyOrder, flattenBundleArrangement, validateBundleArrangement } = await arrangementModule;
  const arrangement = bundleArrangementFromLegacyOrder(["one", "two"]);
  assert.deepEqual(arrangement, { version: 1, nodes: [{ type: "exhibit", exhibitId: "one" }, { type: "exhibit", exhibitId: "two" }] });
  assert.deepEqual(flattenBundleArrangement(arrangement), ["one", "two"]);
  assert.throws(() => bundleArrangementFromLegacyOrder(["one", "one"]), /occurs more than once/i);
  assert.throws(() => validateBundleArrangement({ version: 1, nodes: [{ type: "section", id: "s", heading: " ", exhibits: [] }] }), /heading/i);
});

test("sections, moves and delete-keep-items preserve every exhibit exactly once", async () => {
  const {
    addArrangementSection,
    bundleArrangementFromLegacyOrder,
    deleteArrangementSectionKeepItems,
    flattenBundleArrangement,
    moveArrangementExhibit,
    moveArrangementSection,
    moveArrangementSectionBefore,
    renameArrangementSection,
  } = await arrangementModule;
  let arrangement = bundleArrangementFromLegacyOrder(["a", "b", "c", "d"]);
  arrangement = addArrangementSection(arrangement, { id: "agreements", heading: "Agreements", index: 1, exhibitIds: ["b", "d"] });
  assert.deepEqual(arrangement.nodes, [
    { type: "exhibit", exhibitId: "a" },
    { type: "section", id: "agreements", heading: "Agreements", exhibits: [{ type: "exhibit", exhibitId: "b" }, { type: "exhibit", exhibitId: "d" }] },
    { type: "exhibit", exhibitId: "c" },
  ]);
  arrangement = moveArrangementExhibit(arrangement, "c", { sectionId: "agreements", index: 1 });
  arrangement = renameArrangementSection(arrangement, "agreements", "Core agreements");
  arrangement = moveArrangementSection(arrangement, "agreements", 0);
  assert.deepEqual(flattenBundleArrangement(arrangement), ["b", "c", "d", "a"]);
  arrangement = deleteArrangementSectionKeepItems(arrangement, "agreements");
  assert.deepEqual(arrangement.nodes.map((node) => node.type === "exhibit" ? node.exhibitId : node.id), ["b", "c", "d", "a"]);
});

test("dragging a heading before another heading keeps its exhibits together", async () => {
  const {
    addArrangementSection,
    bundleArrangementFromLegacyOrder,
    flattenBundleArrangement,
    moveArrangementSectionBefore,
  } = await arrangementModule;
  let arrangement = bundleArrangementFromLegacyOrder(["a", "b", "c", "d"]);
  arrangement = addArrangementSection(arrangement, { id: "emails", heading: "Emails", index: 0, exhibitIds: ["a", "b"] });
  arrangement = addArrangementSection(arrangement, { id: "agreements", heading: "Agreements", index: 1, exhibitIds: ["c", "d"] });
  const emailsIndex = arrangement.nodes.findIndex((node) => node.type === "section" && node.id === "emails");
  arrangement = moveArrangementSectionBefore(arrangement, "agreements", emailsIndex);
  assert.deepEqual(flattenBundleArrangement(arrangement), ["c", "d", "a", "b"]);
  assert.deepEqual(arrangement.nodes.map((node) => node.type === "section" ? node.id : node.exhibitId), ["agreements", "emails"]);
  const agreements = arrangement.nodes[0];
  assert.equal(agreements.type, "section");
  if (agreements.type === "section") {
    assert.deepEqual(agreements.exhibits.map((exhibit) => exhibit.exhibitId), ["c", "d"]);
  }
  const emailsAfterMove = arrangement.nodes.findIndex((node) => node.type === "section" && node.id === "emails");
  arrangement = moveArrangementSectionBefore(arrangement, "agreements", emailsAfterMove);
  assert.deepEqual(flattenBundleArrangement(arrangement), ["c", "d", "a", "b"]);
});

test("reconciliation retains section structure, drops stale IDs and appends new IDs", async () => {
  const { reconcileBundleArrangement } = await arrangementModule;
  const original = {
    version: 1,
    nodes: [
      { type: "section", id: "s", heading: "Evidence", exhibits: [{ type: "exhibit", exhibitId: "keep" }, { type: "exhibit", exhibitId: "stale" }] },
      { type: "exhibit", exhibitId: "also-keep" },
    ],
  };
  const reconciled = reconcileBundleArrangement(original, ["keep", "also-keep", "new"]);
  assert.deepEqual(reconciled, {
    version: 1,
    nodes: [
      { type: "section", id: "s", heading: "Evidence", exhibits: [{ type: "exhibit", exhibitId: "keep" }] },
      { type: "exhibit", exhibitId: "also-keep" },
      { type: "exhibit", exhibitId: "new" },
    ],
  });
  assert.deepEqual(original.nodes[0].exhibits.map((node) => node.exhibitId), ["keep", "stale"], "the operation is pure");
});

test("reconciliation places a new email-child exhibit immediately after its parent", async () => {
  const { reconcileBundleArrangement, flattenBundleArrangement, addArrangementSection, bundleArrangementFromLegacyOrder } = await arrangementModule;
  const original = bundleArrangementFromLegacyOrder(["candidate-email", "candidate-other"]);
  const afterParent = reconcileBundleArrangement(original, ["candidate-email", "candidate-other", "candidate-child"], { "candidate-child": "candidate-email" });
  assert.deepEqual(flattenBundleArrangement(afterParent), ["candidate-email", "candidate-child", "candidate-other"]);
  const second = reconcileBundleArrangement(afterParent, ["candidate-email", "candidate-other", "candidate-child", "candidate-child-2"], {
    "candidate-child": "candidate-email",
    "candidate-child-2": "candidate-email",
  });
  assert.deepEqual(flattenBundleArrangement(second), ["candidate-email", "candidate-child", "candidate-child-2", "candidate-other"]);
  let sectioned = addArrangementSection(bundleArrangementFromLegacyOrder(["a", "b", "c"]), { id: "emails", heading: "Emails", index: 0, exhibitIds: ["a", "b"] });
  sectioned = reconcileBundleArrangement(sectioned, ["a", "b", "c", "child"], { child: "a" });
  assert.deepEqual(flattenBundleArrangement(sectioned), ["a", "child", "b", "c"]);
});

test("reconciliation reseats an existing email-child exhibit after its parent", async () => {
  const { reconcileBundleArrangement, flattenBundleArrangement } = await arrangementModule;
  const drifted = {
    version: 1,
    nodes: [
      { type: "exhibit", exhibitId: "candidate-email" },
      { type: "exhibit", exhibitId: "candidate-other" },
      { type: "exhibit", exhibitId: "candidate-child" },
    ],
  };
  const reseated = reconcileBundleArrangement(drifted, ["candidate-email", "candidate-other", "candidate-child"], {
    "candidate-child": "candidate-email",
  });
  assert.deepEqual(flattenBundleArrangement(reseated), ["candidate-email", "candidate-child", "candidate-other"]);
});

test("unheaded earlier/later/top stay inside the contiguous run and cannot jump a heading", async () => {
  const { bundleArrangementFromLegacyOrder, addArrangementSection, exhibitContainerLocation, flattenBundleArrangement, moveArrangementExhibitInContainer } = await arrangementModule;
  let arrangement = bundleArrangementFromLegacyOrder(["a", "b", "c"]);
  arrangement = addArrangementSection(arrangement, { id: "agreements", heading: "Agreements", index: 2, exhibitIds: ["c"] });
  assert.deepEqual(flattenBundleArrangement(arrangement), ["a", "b", "c"]);
  assert.deepEqual(exhibitContainerLocation(arrangement, "a"), { sectionId: null, index: 0, length: 2, topLevelStart: 0 });
  assert.deepEqual(exhibitContainerLocation(arrangement, "b"), { sectionId: null, index: 1, length: 2, topLevelStart: 0 });
  assert.deepEqual(exhibitContainerLocation(arrangement, "c"), { sectionId: "agreements", index: 0, length: 1, topLevelStart: 0 });
  arrangement = moveArrangementExhibitInContainer(arrangement, "a", 1);
  assert.deepEqual(flattenBundleArrangement(arrangement), ["b", "a", "c"]);
  assert.throws(() => moveArrangementExhibitInContainer(arrangement, "a", 2), /outside the current container/i);
  arrangement = { version: 1, nodes: [...arrangement.nodes, { type: "exhibit", exhibitId: "d" }] };
  assert.deepEqual(exhibitContainerLocation(arrangement, "d"), { sectionId: null, index: 0, length: 1, topLevelStart: 3 });
  arrangement = moveArrangementExhibitInContainer(arrangement, "d", 0);
  assert.deepEqual(flattenBundleArrangement(arrangement), ["b", "a", "c", "d"]);
  assert.equal(arrangement.nodes.at(-1).type, "exhibit");
  assert.equal(arrangement.nodes.at(-1).exhibitId, "d");
});

test("sorting is confined to each section and existing unsectioned slots", async () => {
  const { sortBundleArrangementWithinSections } = await arrangementModule;
  const arrangement = {
    version: 1,
    nodes: [
      { type: "exhibit", exhibitId: "z" },
      { type: "exhibit", exhibitId: "y" },
      { type: "section", id: "s1", heading: "First", exhibits: [{ type: "exhibit", exhibitId: "d" }, { type: "exhibit", exhibitId: "b" }] },
      { type: "exhibit", exhibitId: "a" },
      { type: "section", id: "s2", heading: "Second", exhibits: [{ type: "exhibit", exhibitId: "c" }, { type: "exhibit", exhibitId: "a2" }] },
    ],
  };
  const sorted = sortBundleArrangementWithinSections(arrangement, (left, right) => left.localeCompare(right));
  assert.deepEqual(sorted.nodes, [
    { type: "exhibit", exhibitId: "y" },
    { type: "exhibit", exhibitId: "z" },
    { type: "section", id: "s1", heading: "First", exhibits: [{ type: "exhibit", exhibitId: "b" }, { type: "exhibit", exhibitId: "d" }] },
    { type: "exhibit", exhibitId: "a" },
    { type: "section", id: "s2", heading: "Second", exhibits: [{ type: "exhibit", exhibitId: "a2" }, { type: "exhibit", exhibitId: "c" }] },
  ]);
});
