import assert from "node:assert/strict";
import { webcrypto } from "node:crypto";
import test from "node:test";

globalThis.crypto ??= webcrypto;

test("rebuild comparison ignores generated timestamps and reports substantive categories", async () => {
  const { createSubstantiveBuildSnapshot, compareSubstantiveBuilds } = await import("../app/lib/rebuild-comparison.ts");
  const analysis = { statementHash: "a".repeat(64), evidence: [{ id: "e", sha256: "b".repeat(64) }] };
  const candidate = { id: "c", included: true, confirmed: true, evidenceId: "e", description: "Contract" };
  const pagination = { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false };
  const layout = { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 };
  const first = await createSubstantiveBuildSnapshot({ analysis, candidates: [candidate], canonicalOrder: ["c"], templates: [], layout, pagination, resolutions: [] });
  const same = await createSubstantiveBuildSnapshot({ analysis: { ...analysis, generatedAt: new Date().toISOString() }, candidates: [candidate], canonicalOrder: ["c"], templates: [], layout, pagination, resolutions: [] });
  assert.equal(compareSubstantiveBuilds(first, same).summary, "No substantive change");
  const changed = await createSubstantiveBuildSnapshot({ analysis, candidates: [{ ...candidate, description: "Amended contract" }], canonicalOrder: ["c"], templates: [], layout: { ...layout, volumePageLimit: 50 }, pagination, resolutions: [] });
  assert.deepEqual(compareSubstantiveBuilds(first, changed).categories, ["Index description changed", "Volume limit changed"]);
});

test("unused optional templates do not change the substantive fingerprint", async () => {
  const { createSubstantiveBuildSnapshot, compareSubstantiveBuilds } = await import("../app/lib/rebuild-comparison.ts");
  const analysis = { statementHash: "a".repeat(64), evidence: [{ id: "e", sha256: "b".repeat(64) }] };
  const candidate = { id: "c", included: true, confirmed: true, evidenceId: "e", description: "Contract" };
  const pagination = { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false };
  const layout = { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 };
  const baseline = await createSubstantiveBuildSnapshot({ analysis, candidates: [candidate], canonicalOrder: ["c"], templates: [], layout, pagination, resolutions: [] });
  const unusedDivider = { slot: "divider", file: new File(["template"], "Divider.pdf"), sha256: "c".repeat(64), sourceFormat: "pdf" };
  const selectedButUnused = await createSubstantiveBuildSnapshot({ analysis, candidates: [candidate], canonicalOrder: ["c"], templates: [unusedDivider], layout, pagination, resolutions: [] });
  assert.equal(compareSubstantiveBuilds(baseline, selectedButUnused).summary, "No substantive change");
  const includedDivider = await createSubstantiveBuildSnapshot({ analysis, candidates: [candidate], canonicalOrder: ["c"], templates: [unusedDivider], layout: { ...layout, includeDividerPages: true }, pagination, resolutions: [] });
  assert.ok(compareSubstantiveBuilds(baseline, includedDivider).categories.includes("Template changed"));
  const exactCover = await createSubstantiveBuildSnapshot({ analysis, candidates: [candidate], canonicalOrder: ["c"], templates: [], layout: { ...layout, coverInsertion: "exact", exactCoverPageNumber: true }, pagination, resolutions: [] });
  assert.ok(compareSubstantiveBuilds(baseline, exactCover).categories.includes("Cover treatment changed"));
});

test("manual provenance and document date are substantive rebuild inputs", async () => {
  const { createSubstantiveBuildSnapshot, compareSubstantiveBuilds } = await import("../app/lib/rebuild-comparison.ts");
  const analysis = { statementHash: "a".repeat(64), evidence: [{ id: "e", sha256: "b".repeat(64) }] };
  const base = { id: "manual", included: true, confirmed: true, evidenceId: "e", description: "Added document", date: "1 August 2026", manualAddition: true, manualWarningAcknowledgedAt: "2026-08-08T00:00:00.000Z" };
  const pagination = { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false };
  const layout = { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 };
  const first = await createSubstantiveBuildSnapshot({ analysis, candidates: [base], canonicalOrder: ["manual"], templates: [], layout, pagination, resolutions: [] });
  const changed = await createSubstantiveBuildSnapshot({ analysis, candidates: [{ ...base, date: "2 August 2026", manualAddition: false }], canonicalOrder: ["manual"], templates: [], layout, pagination, resolutions: [] });
  const difference = compareSubstantiveBuilds(first, changed);
  assert.ok(difference.categories.includes("Document date changed"));
  assert.ok(difference.categories.includes("Manual exhibit status changed"));
});

test("section headings and membership are substantive arrangement inputs", async () => {
  const { createSubstantiveBuildSnapshot, compareSubstantiveBuilds } = await import("../app/lib/rebuild-comparison.ts");
  const analysis = { statementHash: "a".repeat(64), evidence: [{ id: "e", sha256: "b".repeat(64) }] };
  const candidates = [{ id: "a", included: true, confirmed: true, evidenceId: "e", description: "Contract" }, { id: "b", included: true, confirmed: true, evidenceId: "e", description: "Letter" }];
  const pagination = { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false };
  const layout = { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 };
  const firstArrangement = { version: 1, nodes: [{ type: "section", id: "s", heading: "Agreements", exhibits: [{ type: "exhibit", exhibitId: "a" }] }, { type: "exhibit", exhibitId: "b" }] };
  const renamedArrangement = { version: 1, nodes: [{ type: "section", id: "s", heading: "Core agreements", exhibits: [{ type: "exhibit", exhibitId: "a" }] }, { type: "exhibit", exhibitId: "b" }] };
  const reorderedArrangement = { version: 1, nodes: [{ type: "exhibit", exhibitId: "b" }, { type: "section", id: "s", heading: "Agreements", exhibits: [{ type: "exhibit", exhibitId: "a" }] }] };
  const first = await createSubstantiveBuildSnapshot({ analysis, candidates, arrangement: firstArrangement, templates: [], layout, pagination, resolutions: [] });
  const renamed = await createSubstantiveBuildSnapshot({ analysis, candidates, arrangement: renamedArrangement, templates: [], layout, pagination, resolutions: [] });
  const reordered = await createSubstantiveBuildSnapshot({ analysis, candidates, arrangement: reorderedArrangement, templates: [], layout, pagination, resolutions: [] });
  assert.deepEqual(compareSubstantiveBuilds(first, renamed).categories, ["Index sections changed"]);
  assert.ok(compareSubstantiveBuilds(first, reordered).categories.includes("Order changed"));
});

test("adding a flat schema-8 arrangement does not report a false rebuild change", async () => {
  const { createSubstantiveBuildSnapshot, compareSubstantiveBuilds } = await import("../app/lib/rebuild-comparison.ts");
  const analysis = { statementHash: "a".repeat(64), evidence: [] };
  const pagination = { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false };
  const layout = { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 };
  const current = await createSubstantiveBuildSnapshot({ analysis, candidates: [], arrangement: { version: 1, nodes: [{ type: "exhibit", exhibitId: "a" }] }, templates: [], layout, pagination, resolutions: [] });
  const { arrangement: _arrangement, ...legacyCanonical } = current.canonical;
  const legacy = { fingerprint: "legacy-fingerprint", canonical: legacyCanonical };
  assert.equal(compareSubstantiveBuilds(legacy, current).changed, false);
});
