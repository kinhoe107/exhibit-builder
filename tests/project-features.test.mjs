import assert from "node:assert/strict";
import { webcrypto } from "node:crypto";
import { readFile } from "node:fs/promises";
import { File } from "node:buffer";
import { basename } from "node:path";
import test from "node:test";
import JSZip from "jszip";

test("pagination changes preserve uncommitted prefix draft state", async () => {
  const { DEFAULT_PAGINATION, updatePaginationDraft, commitPaginationChange, paginationDraftAfterWarningAccept, paginationDraftAfterWarningCancel } = await import("../app/lib/bundle-types.ts");
  const committed = { ...DEFAULT_PAGINATION, matchPdfPageOrder: false };
  const typed = updatePaginationDraft(committed, { prefix: "AX-", startAt: 10 });
  const confirmation = updatePaginationDraft(typed, { padding: 4 });
  assert.equal(confirmation.prefix, "AX-");
  assert.equal(confirmation.padding, 4);
  assert.equal(DEFAULT_PAGINATION.includePrefixInIndex, true);
  const instant = commitPaginationChange(committed, typed, { scheme: "bates" });
  assert.equal(instant.pagination.prefix, "");
  assert.equal(instant.pagination.scheme, "bates");
  assert.equal(instant.draft.prefix, "AX-");
  assert.equal(instant.draft.scheme, "bates");
  const applied = commitPaginationChange(instant.pagination, instant.draft, {});
  assert.equal(applied.pagination.prefix, "AX-");
  assert.equal(applied.draft.prefix, "AX-");
  const oneSequence = { matchPdfPageOrder: true, countTemplates: true, preliminary: "arabic", scheme: "bundle", startAt: 1 };
  const warned = commitPaginationChange(committed, typed, oneSequence);
  assert.equal(warned.pagination.startAt, 1);
  assert.equal(warned.draft.startAt, 10, "one-sequence radio must not clobber an unapplied start-at before Cancel or accept");
  assert.equal(warned.draft.matchPdfPageOrder, true);
  const cancelled = paginationDraftAfterWarningCancel(committed, warned.draft);
  assert.equal(cancelled.matchPdfPageOrder, false);
  assert.equal(cancelled.startAt, 10);
  assert.equal(cancelled.prefix, "AX-");
  const accepted = paginationDraftAfterWarningAccept(warned.pagination, warned.draft, oneSequence);
  assert.equal(accepted.matchPdfPageOrder, true);
  assert.equal(accepted.startAt, 1, "accepting one sequence commits its start-at of 1");
  assert.equal(accepted.prefix, "AX-");
  const schemeOnly = paginationDraftAfterWarningAccept(instant.pagination, instant.draft, { scheme: "bates" });
  assert.equal(schemeOnly.startAt, 10, "accepting a scheme change keeps an unapplied start-at");
  assert.equal(schemeOnly.prefix, "AX-");
  const noWarning = paginationDraftAfterWarningAccept(warned.pagination, warned.draft, oneSequence);
  assert.equal(noWarning.startAt, 1, "a one-sequence radio that commits without a warning still writes its start-at into the draft");
  assert.equal(noWarning.prefix, "AX-");
  assert.notEqual(JSON.stringify(noWarning), JSON.stringify({ ...noWarning, startAt: warned.draft.startAt }), "Apply must not remain armed solely because startAt was gated out of the instant draft");
});

test("Apply still commits a typed prefix after the finished-PDF numbering lock", async () => {
  const { DEFAULT_PAGINATION, updatePaginationDraft, commitPaginationChange, lockPagination } = await import("../app/lib/bundle-types.ts");
  const committed = { ...DEFAULT_PAGINATION };
  const typed = updatePaginationDraft(committed, { prefix: "QA-", suffix: "-END" });
  const polluted = commitPaginationChange(committed, typed, { matchPdfPageOrder: true });
  assert.equal(polluted.pagination.prefix, "", "a numbering-lock flag must not masquerade as Apply");
  const applied = commitPaginationChange(committed, typed, {});
  assert.equal(applied.pagination.prefix, "QA-");
  assert.equal(applied.pagination.suffix, "-END");
  const locked = lockPagination(applied.pagination);
  assert.equal(locked.prefix, "QA-");
  assert.equal(locked.suffix, "-END");
  assert.equal(locked.matchPdfPageOrder, true);
});

test("optional pages always count because the finished PDF is the only numbering source", async () => {
  const { countsOptionalPagesInReferences, lockPagination } = await import("../app/lib/bundle-types.ts");
  assert.equal(countsOptionalPagesInReferences({ matchPdfPageOrder: true }, { countOptionalPagesInReferences: false }), true);
  assert.equal(countsOptionalPagesInReferences({ matchPdfPageOrder: false }, { countOptionalPagesInReferences: false }), true);
  assert.equal(countsOptionalPagesInReferences({ matchPdfPageOrder: false }, { countOptionalPagesInReferences: true }), true);
  const locked = lockPagination({ matchPdfPageOrder: false, scheme: "section", countTemplates: false, preliminary: "roman", includePrefixInIndex: false });
  assert.equal(locked.matchPdfPageOrder, true);
  assert.equal(locked.scheme, "bundle");
  assert.equal(locked.countTemplates, true);
  assert.equal(locked.preliminary, "arabic");
  assert.equal(locked.includePrefixInIndex, true);
});

globalThis.File = File;
globalThis.crypto ??= webcrypto;

async function sha256(file) {
  const digest = await crypto.subtle.digest("SHA-256", await file.arrayBuffer());
  return Array.from(new Uint8Array(digest)).map((part) => part.toString(16).padStart(2, "0")).join("");
}

const fixtureRoot = new URL("./fixtures/core/", import.meta.url);
const evidenceRoot = new URL("./fixtures/core/Evidence_Inbox/", import.meta.url);

async function fixtureFile(url) {
  return new File([await readFile(url)], basename(url.pathname));
}

test("saves and reopens a serializable local bundle project", async () => {
  const { createProjectArchive, openProjectArchive } = await import("../app/lib/project-archive.ts");
  const statement = new File(["statement bytes"], "Statement.docx");
  const evidence = new File(["evidence bytes"], "Evidence.pdf");
  const bytes = await createProjectArchive({
    schemaVersion: 2,
    name: "Local bundle",
    createdAt: "2026-07-30T00:00:00.000Z",
    updatedAt: "2026-07-30T00:00:00.000Z",
    profileId: "exhibit-neutral",
    pagination: { scheme: "bates", prefix: "ABC-", suffix: "", startAt: 1, padding: 6, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8 },
    witnessSettings: { witness: { initials: "AH", nextNumber: 2 } },
    candidates: [{ id: "candidate-1", confirmed: true }],
    analysis: { bundleOnly: true },
  }, [
    { id: "statement-1", role: "statement", name: statement.name, sha256: await sha256(statement), file: statement },
    { id: "evidence-1", role: "evidence", name: evidence.name, sha256: await sha256(evidence), file: evidence },
  ]);
  const opened = await openProjectArchive(new File([bytes], "Local.bundle-project"));
  assert.equal(opened.snapshot.name, "Local bundle");
  assert.equal(opened.snapshot.pagination.prefix, "ABC-");
  assert.equal(opened.snapshot.pagination.volumeNumbering, "continuous", "older projects migrate to the safe continuous-volume default");
  assert.deepEqual(opened.sources.map((source) => source.role), ["statement", "evidence"]);
  assert.equal(await opened.sources[1].file.text(), "evidence bytes");
  assert.equal(opened.snapshot.schemaVersion, 8);
  assert.deepEqual(opened.snapshot.arrangement, { version: 1, nodes: [] }, "older projects migrate to an explicit empty arrangement");
  assert.equal("finalOrder" in opened.snapshot, false, "schema 8 exposes only the authoritative arrangement");
});

test("refuses to save a project when selected bytes no longer match their reviewed hash", async () => {
  const { createProjectArchive } = await import("../app/lib/project-archive.ts");
  const reviewed = new File(["reviewed"], "Evidence.pdf");
  const replaced = new File(["replaced"], "Evidence.pdf");
  await assert.rejects(
    createProjectArchive({
      schemaVersion: 8, name: "Changed", createdAt: "2026-08-21T00:00:00.000Z", updatedAt: "2026-08-21T00:00:00.000Z", profileId: "exhibit-neutral",
      pagination: { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
      witnessSettings: {}, candidates: [], analysis: {}, arrangement: { version: 1, nodes: [] },
    }, [{ id: "evidence", role: "evidence", name: replaced.name, sha256: await sha256(reviewed), file: replaced }]),
    /changed after it was selected/i,
  );
});

test("migrates canonical final order to schema 8 and preserves 0.10 layout settings", async () => {
  const { createProjectArchive, openProjectArchive } = await import("../app/lib/project-archive.ts");
  const statement = new File(["statement"], "Statement.docx");
  const evidence = new File(["evidence"], "Evidence.pdf");
  const bytes = await createProjectArchive({
    schemaVersion: 6, name: "Ordered", createdAt: "2026-08-03T00:00:00.000Z", updatedAt: "2026-08-03T00:00:00.000Z", profileId: "exhibit-neutral",
    pagination: { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
    layout: { includeDividerPages: true, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 75 },
    witnessSettings: {}, candidates: [], analysis: {}, finalOrder: ["group-b", "group-a"], pageSizeChoices: { evidence: "keep-original" },
  }, [
    { id: "statement", role: "statement", name: statement.name, sha256: await sha256(statement), file: statement },
    { id: "evidence", role: "evidence", name: evidence.name, sha256: await sha256(evidence), file: evidence },
  ]);
  const opened = await openProjectArchive(new File([bytes], "Ordered.bundle-project"));
  assert.deepEqual(opened.snapshot.arrangement, { version: 1, nodes: [{ type: "exhibit", exhibitId: "group-b" }, { type: "exhibit", exhibitId: "group-a" }] });
  assert.equal(opened.snapshot.layout.volumePageLimit, 75);
  assert.equal(opened.snapshot.layout.includeDividerPages, true);
  assert.equal(opened.snapshot.layout.coverInsertion, "fit-a4");
  assert.equal(opened.snapshot.layout.exactCoverPageNumber, false);
  assert.equal(opened.snapshot.layout.exactCoverVolumeLabel, false);
  assert.deepEqual(opened.snapshot.layout.builtInMatter, { matterNumbers: [], partyNames: [], forums: [], matterTitles: [] });
  assert.equal(opened.snapshot.pagination.volumeNumbering, "continuous");
  assert.deepEqual(opened.snapshot.pageSizeChoices, { evidence: "keep-original" });
});

test("schema 6 source-hash finalOrder is translated to current group IDs on exhibit reconcile", async () => {
  const { createProjectArchive, openProjectArchive } = await import("../app/lib/project-archive.ts");
  const { flattenBundleArrangement } = await import("../app/lib/bundle-arrangement.ts");
  const { deriveExhibitGroups, legacyGroupId, reconcileExhibitArrangement } = await import("../app/lib/exhibit-groups.ts");
  const statement = new File(["statement"], "Statement.docx");
  const evidenceA = new File(["evidence-a"], "A.pdf");
  const evidenceB = new File(["evidence-b"], "B.pdf");
  const groups = deriveExhibitGroups({
    evidence: [
      { id: "e-a", name: evidenceA.name, sha256: "hash-a", extension: "pdf" },
      { id: "e-b", name: evidenceB.name, sha256: "hash-b", extension: "pdf" },
    ],
  }, [
    { id: "one", included: true, evidenceId: "e-a", sequenceOrder: 1, exhibitInitials: "AH", exhibitSequence: 1, witnessInitials: "AH", witnessKey: "AH" },
    { id: "two", included: true, evidenceId: "e-b", sequenceOrder: 2, exhibitInitials: "AH", exhibitSequence: 1, witnessInitials: "AH", witnessKey: "AH" },
  ]);
  const reversedLegacy = groups.map((group) => legacyGroupId(group)).reverse();
  const bytes = await createProjectArchive({
    schemaVersion: 6, name: "Hashed order", createdAt: "2026-08-24T00:00:00.000Z", updatedAt: "2026-08-24T00:00:00.000Z", profileId: "exhibit-neutral",
    pagination: { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
    witnessSettings: {}, candidates: [], analysis: {}, finalOrder: reversedLegacy,
  }, [
    { id: "statement", role: "statement", name: statement.name, sha256: await sha256(statement), file: statement },
    { id: "evidence-a", role: "evidence", name: evidenceA.name, sha256: await sha256(evidenceA), file: evidenceA },
    { id: "evidence-b", role: "evidence", name: evidenceB.name, sha256: await sha256(evidenceB), file: evidenceB },
  ]);
  const opened = await openProjectArchive(new File([bytes], "Hashed.bundle-project"));
  assert.deepEqual(flattenBundleArrangement(opened.snapshot.arrangement), reversedLegacy, "archive migration keeps the stored historical identities");
  assert.deepEqual(
    flattenBundleArrangement(reconcileExhibitArrangement(opened.snapshot.arrangement, groups)),
    groups.map((group) => group.id).reverse(),
    "saved-project open then UI reconcile restores the committed custom order",
  );
});

test("schema 8 project round-trips sections without writing legacy finalOrder", async () => {
  const { createProjectArchive, openProjectArchive } = await import("../app/lib/project-archive.ts");
  const evidence = new File(["evidence"], "Evidence.pdf");
  const arrangement = {
    version: 1,
    nodes: [
      { type: "section", id: "agreements", heading: "Agreements", exhibits: [{ type: "exhibit", exhibitId: "group-a" }] },
      { type: "exhibit", exhibitId: "group-b" },
    ],
  };
  const bytes = await createProjectArchive({
    schemaVersion: 8, name: "Sections", createdAt: "2026-08-13T00:00:00.000Z", updatedAt: "2026-08-13T00:00:00.000Z", profileId: "exhibit-neutral",
    pagination: { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
    witnessSettings: {}, candidates: [], analysis: {}, arrangement,
  }, [{ id: "evidence", role: "evidence", name: evidence.name, sha256: await sha256(evidence), file: evidence }]);
  const zip = await JSZip.loadAsync(bytes);
  const stored = JSON.parse(await zip.file("bundle-project.json").async("text"));
  assert.equal(stored.schemaVersion, 8);
  assert.deepEqual(stored.arrangement, arrangement);
  assert.equal("finalOrder" in stored, false);
  const opened = await openProjectArchive(new File([bytes], "Sections.bundle-project"));
  assert.deepEqual(opened.snapshot.arrangement, arrangement);
});

test("portable project import resets human approvals and technical exceptions", async () => {
  const { createProjectArchive, openProjectArchive } = await import("../app/lib/project-archive.ts");
  const evidence = new File(["evidence"], "Evidence.pdf");
  const hash = await sha256(evidence);
  const bytes = await createProjectArchive({
    schemaVersion: 8, name: "Portable", createdAt: "2026-08-21T00:00:00.000Z", updatedAt: "2026-08-21T00:00:00.000Z", profileId: "exhibit-neutral",
    pagination: { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
    witnessSettings: {}, analysis: {}, arrangement: { version: 1, nodes: [] },
    candidates: [{ id: "candidate", evidenceId: "evidence", included: true, confirmed: true, confirmedAt: "2026-08-21T00:00:00.000Z", confirmationMethod: "individual" }],
    resolutions: [{ blockerId: "ocr-evidence", action: "proceed-without-ocr", sourceId: "evidence", sourceSha256: hash, approvedAt: "2026-08-21T00:00:00.000Z", visualReviewConfirmed: true }],
  }, [{ id: "evidence", role: "evidence", name: evidence.name, sha256: hash, file: evidence }]);
  const opened = await openProjectArchive(new File([bytes], "Portable.bundle-project"));
  assert.deepEqual(opened.snapshot.candidates, [{ id: "candidate", evidenceId: "evidence", included: true, confirmed: false }]);
  assert.deepEqual(opened.snapshot.resolutions, []);
});

test("schema 7 project preserves manual exhibit provenance and source bytes", async () => {
  const { createProjectArchive, openProjectArchive } = await import("../app/lib/project-archive.ts");
  const statement = new File(["statement"], "Guided.docx");
  const evidence = new File(["manual evidence"], "Uncited.pdf");
  const candidate = { id: "manual-1", evidenceId: "evidence-manual", description: "Uncited document", date: "8 August 2026", paragraph: 0, citation: "", included: true, confirmed: true, manualAddition: true, manualAddedAt: "2026-08-08T10:00:00.000Z", manualWarningAcknowledgedAt: "2026-08-08T10:00:00.000Z" };
  const bytes = await createProjectArchive({
    schemaVersion: 7, name: "Manual", createdAt: "2026-08-08T10:00:00.000Z", updatedAt: "2026-08-08T10:00:00.000Z", profileId: "exhibit-neutral",
    pagination: { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
    layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 },
    witnessSettings: {}, candidates: [candidate], analysis: {}, finalOrder: ["manual-group"],
  }, [
    { id: "statement", role: "statement", name: statement.name, sha256: await sha256(statement), file: statement },
    { id: "evidence-manual", role: "evidence", name: evidence.name, sha256: await sha256(evidence), file: evidence },
  ]);
  const opened = await openProjectArchive(new File([bytes], "Manual.bundle-project"));
  assert.equal(opened.snapshot.schemaVersion, 8);
  const { manualWarningAcknowledgedAt: _resetManualApproval, ...unapprovedCandidate } = candidate;
  assert.deepEqual(opened.snapshot.candidates, [{ ...unapprovedCandidate, confirmed: false }]);
  assert.deepEqual(opened.snapshot.arrangement, { version: 1, nodes: [{ type: "exhibit", exhibitId: "manual-group" }] });
  assert.equal(await opened.sources.find((source) => source.id === "evidence-manual").file.text(), "manual evidence");
});

test("round-trips legitimate source filenames containing repeated dots", async () => {
  const { createProjectArchive, openProjectArchive } = await import("../app/lib/project-archive.ts");
  const source = new File(["contract"], "contract..final.pdf");
  const bytes = await createProjectArchive({
    schemaVersion: 6, name: "Dots", createdAt: "2026-08-03T00:00:00.000Z", updatedAt: "2026-08-03T00:00:00.000Z", profileId: "exhibit-neutral",
    pagination: { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
    witnessSettings: {}, candidates: [], analysis: {}, finalOrder: [],
  }, [{ id: "evidence", role: "evidence", name: source.name, sha256: await sha256(source), file: source }]);
  const opened = await openProjectArchive(new File([bytes], "Dots.bundle-project"));
  assert.equal(opened.sources[0].name, "contract..final.pdf");
  assert.equal(await opened.sources[0].file.text(), "contract");
});

test("rejects a highly compressed project archive before expanding its entries", async () => {
  const { openProjectArchive } = await import("../app/lib/project-archive.ts");
  const zip = new JSZip();
  zip.file("bundle-project.json", JSON.stringify({ schemaVersion: 7, sources: [] }));
  zip.file("sources/compressed-bomb.pdf", "A".repeat(2 * 1024 * 1024));
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 9 } });
  assert.ok(bytes.byteLength < 100_000, "fixture remains small while exercising the expansion-ratio limit");
  await assert.rejects(openProjectArchive(new File([bytes], "Unsafe.bundle-project")), /unsafe paths or limits/i);
});

test("rejects duplicate project source descriptors before extraction", async () => {
  const { openProjectArchive } = await import("../app/lib/project-archive.ts");
  const zip = new JSZip();
  const hash = "a".repeat(64);
  zip.file("bundle-project.json", JSON.stringify({
    schemaVersion: 7,
    sources: [
      { id: "duplicate", role: "evidence", name: "One.pdf", sha256: hash, path: "sources/one.pdf" },
      { id: "duplicate", role: "evidence", name: "Two.pdf", sha256: hash, path: "sources/two.pdf" },
    ],
  }));
  zip.file("sources/one.pdf", "one");
  zip.file("sources/two.pdf", "two");
  const bytes = await zip.generateAsync({ type: "uint8array" });
  await assert.rejects(openProjectArchive(new File([bytes], "Duplicates.bundle-project")), /duplicate source IDs or paths/i);
});

test("saved project sources are expanded and checked sequentially", async () => {
  const source = await readFile(new URL("../app/lib/project-archive.ts", import.meta.url), "utf8");
  const extraction = source.slice(source.indexOf("const sources: ProjectSource[]"), source.indexOf("const { sources: _sources"));
  assert.match(extraction, /for \(const source of stored\.sources\)/);
  assert.doesNotMatch(extraction, /Promise\.all/);
});

test("schema 8 rejects forged or unbounded template-review metadata", async () => {
  const { openProjectArchive } = await import("../app/lib/project-archive.ts");
  const zip = new JSZip();
  zip.file("bundle-project.json", JSON.stringify({
    schemaVersion: 8,
    sources: [],
    arrangement: { version: 1, nodes: [] },
    templateReviews: [{ slot: "cover", sourceId: "template-cover", sourceFormat: "pdf", sourceSha256: "a".repeat(64), pdfSha256: "b".repeat(64), reviewState: { matterReview: { pdfSha256: "c".repeat(64) } } }],
  }));
  const bytes = await zip.generateAsync({ type: "uint8array" });
  await assert.rejects(openProjectArchive(new File([bytes], "Forged.bundle-project")), /invalid template-review metadata/i);
});

test("saved-project recovery overlays only hash-bound current journal deltas", async () => {
  const { mergeRecoveryProjectDeltas } = await import("../app/lib/recovery-restore.ts");
  const currentHash = "d".repeat(64);
  const oldTemplateHash = "a".repeat(64);
  const newTemplateHash = "b".repeat(64);
  const oldTemplate = new File(["old-template"], "Old-divider.docx");
  const newTemplate = new File(["new-template"], "New-divider.docx");
  const opened = {
    snapshot: {
      schemaVersion: 6, name: "Saved", createdAt: "2026-08-03T00:00:00.000Z", updatedAt: "2026-08-03T00:00:00.000Z", profileId: "exhibit-neutral",
      pagination: { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
      layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 },
      witnessSettings: {}, analysis: { statements: [{ id: "statement", witnessName: "Old Witness", witnessInitials: "OW", name: "Statement.docx" }], templates: ["divider"] }, finalOrder: ["old"], resolutions: [],
      candidates: [{ id: "candidate", statementId: "statement", paragraph: 12, citation: "I refer to the revised agreement [AH1/xx].", evidenceId: "evidence", description: "Saved description", confirmed: true }],
    },
    sources: [
      { id: "statement", role: "statement", name: "Statement.docx", sha256: "c".repeat(64), file: new File(["statement"], "Statement.docx") },
      { id: "evidence", role: "evidence", name: "Evidence.pdf", sha256: currentHash, file: new File(["evidence"], "Evidence.pdf") },
      { id: "template-divider", role: "template", name: oldTemplate.name, sha256: oldTemplateHash, file: oldTemplate },
    ],
  };
  const payload = {
    project: { name: "Recovered edits" },
    candidates: [{ id: "candidate", statementId: "statement", paragraph: 12, citation: "I refer to the original agreement [AH1/xx].", evidenceId: "evidence", sourceSha256: currentHash, statementSha256: "c".repeat(64), description: "Edited after save", confirmed: true, repeatDecision: "same" }],
    finalOrder: ["new"],
    layout: { includeDividerPages: true, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 50 },
    pagination: { ...opened.snapshot.pagination, prefix: "REC-" },
    resolutions: [{ blockerId: "source", action: "exclude-source", sourceId: "evidence", sourceSha256: currentHash, approvedAt: "2026-08-03T00:00:00.000Z" }],
    statements: [{ id: "statement", sourceId: "statement", sourceSha256: "c".repeat(64), witnessName: "Updated Witness", witnessInitials: "UW" }],
    templates: [{ sourceId: "template-divider", slot: "divider", sourceFormat: "docx", templateConfirmed: true }],
    sources: [
      { id: "saved-project-archive", role: "project", name: "Saved.bundle-project", path: "C:\\Saved.bundle-project", sha256: "9".repeat(64), size: 100 },
      { id: "template-divider", role: "template", name: newTemplate.name, path: "C:\\New-divider.docx", sha256: newTemplateHash, size: newTemplate.size },
    ],
  };
  const merged = mergeRecoveryProjectDeltas(opened, payload, [{ id: "template-divider", role: "template", name: newTemplate.name, sha256: newTemplateHash, file: newTemplate }]);
  assert.equal(merged.opened.snapshot.name, "Recovered edits");
  assert.equal(merged.opened.snapshot.candidates[0].description, "Edited after save");
  assert.equal(merged.opened.snapshot.candidates[0].repeatDecision, "same", "a hash-bound same/separate decision is restored when statement and source hashes still match");
  assert.deepEqual(merged.opened.snapshot.arrangement, { version: 1, nodes: [{ type: "exhibit", exhibitId: "new" }] });
  assert.equal(merged.opened.snapshot.layout.volumePageLimit, 50);
  assert.equal(merged.opened.snapshot.pagination.prefix, "REC-");
  assert.equal(merged.opened.snapshot.resolutions.length, 1);
  assert.equal(merged.opened.snapshot.analysis.statements[0].witnessName, "Updated Witness");
  assert.equal(merged.opened.snapshot.analysis.statements[0].witnessInitials, "UW");
  assert.ok(merged.opened.sources.some((source) => source.id === "statement"), "archive statement remains when the journal contains only the project path");
  assert.ok(merged.opened.sources.some((source) => source.id === "evidence"), "archive evidence remains when the journal contains only the project path");
  assert.equal(await merged.opened.sources.find((source) => source.id === "template-divider").file.text(), "new-template");
  assert.equal(merged.templateApprovals.get("template-divider"), true);

  const changed = mergeRecoveryProjectDeltas(opened, { ...payload, candidates: [{ ...payload.candidates[0], sourceSha256: "e".repeat(64) }] });
  assert.equal(changed.opened.snapshot.candidates[0].description, "Saved description");
  assert.equal(changed.opened.snapshot.candidates[0].confirmed, false);
  assert.ok(changed.issues.some((issue) => /source hash changed/i.test(issue)));
  assert.ok(changed.issues.some((issue) => /template source is missing or changed/i.test(issue)));
  assert.equal(changed.templateApprovals.has("template-divider"), false);

  const changedStatementHash = "f".repeat(64);
  const wordingChanged = mergeRecoveryProjectDeltas({
    ...opened,
    sources: opened.sources.map((source) => source.id === "statement" ? { ...source, sha256: changedStatementHash, file: new File(["statement citation wording changed"], source.name) } : source),
  }, payload);
  const wordingCandidate = wordingChanged.opened.snapshot.candidates[0];
  assert.equal(wordingCandidate.paragraph, 12, "the fresh candidate remains at the same paragraph and ordinal");
  assert.match(wordingCandidate.citation, /revised agreement/i);
  assert.equal(wordingCandidate.confirmed, false, "an old approval is not restored when only the witness citation wording changed");
  assert.equal(wordingCandidate.repeatDecision, undefined, "a stale same/separate decision is not restored after the statement hash changes");
  assert.ok(wordingChanged.issues.some((issue) => /witness statement changed/i.test(issue)));
});

test("recovery appends extra candidates without restoring a stale repeat decision after a statement change", async () => {
  const { mergeRecoveryProjectDeltas } = await import("../app/lib/recovery-restore.ts");
  const evidenceHash = "d".repeat(64);
  const statementHash = "c".repeat(64);
  const changedStatementHash = "f".repeat(64);
  const opened = {
    snapshot: {
      schemaVersion: 8, name: "Saved", createdAt: "2026-08-03T00:00:00.000Z", updatedAt: "2026-08-03T00:00:00.000Z", profileId: "exhibit-neutral",
      pagination: { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
      layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 },
      witnessSettings: {}, analysis: {}, arrangement: { version: 1, nodes: [] }, resolutions: [],
      candidates: [],
    },
    sources: [
      { id: "statement", role: "statement", name: "Statement.docx", sha256: changedStatementHash, file: new File(["changed statement"], "Statement.docx") },
      { id: "evidence", role: "evidence", name: "Evidence.pdf", sha256: evidenceHash, file: new File(["evidence"], "Evidence.pdf") },
    ],
  };
  const extra = {
    id: "extra", statementId: "statement", evidenceId: "evidence", sourceSha256: evidenceHash, statementSha256: statementHash,
    description: "Later citation", confirmed: true, included: true, repeatDecision: "same",
  };
  const manual = {
    id: "manual", evidenceId: "evidence", sourceSha256: evidenceHash, statementSha256: statementHash,
    description: "Uncited extra copy", confirmed: true, included: true, manualAddition: true, repeatDecision: "separate",
  };
  const payload = {
    project: { name: "Saved" },
    candidates: [extra, manual],
    statements: [{ id: "statement", sourceId: "statement", sourceSha256: statementHash, witnessName: "Witness", witnessInitials: "WS" }],
    sources: [{ id: "saved-project-archive", role: "project", name: "Saved.bundle-project", path: "C:\\Saved.bundle-project", sha256: "9".repeat(64), size: 100 }],
  };
  const merged = mergeRecoveryProjectDeltas(opened, payload);
  const restoredExtra = merged.opened.snapshot.candidates.find((candidate) => candidate.id === "extra");
  const restoredManual = merged.opened.snapshot.candidates.find((candidate) => candidate.id === "manual");
  assert.equal(restoredExtra?.confirmed, false);
  assert.equal(restoredExtra?.repeatDecision, undefined, "an appended cited repeat decision is dropped when the statement hash changed");
  assert.equal(restoredManual?.confirmed, true);
  assert.equal(restoredManual?.repeatDecision, "separate", "a manual-addition duplicate decision still restores");
});

test("recovery restores email-child confirmation by child hash without an archive source", async () => {
  const { mergeRecoveryProjectDeltas } = await import("../app/lib/recovery-restore.ts");
  const childHash = "f".repeat(64);
  const parentHash = "d".repeat(64);
  const statementHash = "c".repeat(64);
  const opened = {
    snapshot: {
      schemaVersion: 8, name: "Saved", createdAt: "2026-08-03T00:00:00.000Z", updatedAt: "2026-08-03T00:00:00.000Z", profileId: "exhibit-neutral",
      pagination: { scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "roman", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
      layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 },
      witnessSettings: {}, analysis: {}, finalOrder: ["candidate-email"], resolutions: [],
      candidates: [
        { id: "email-card", evidenceId: "evidence", description: "Notice email", confirmed: true, included: true },
        {
          id: "stale-child-id",
          evidenceId: "derived-missing",
          description: "Invoice",
          confirmed: false,
          included: true,
          parentEmailProvenance: { parentName: "Notice.eml", parentSha256: parentHash, childIdentity: "child-id", childSha256: childHash },
        },
      ],
    },
    sources: [
      { id: "statement", role: "statement", name: "Statement.docx", sha256: statementHash, file: new File(["statement"], "Statement.docx") },
      { id: "evidence", role: "evidence", name: "Notice.eml", sha256: parentHash, file: new File(["email"], "Notice.eml") },
    ],
  };
  const payload = {
    project: { name: "Saved" },
    candidates: [{
      id: "journal-child",
      evidenceId: "derived-missing",
      sourceSha256: childHash,
      statementId: "statement",
      statementSha256: statementHash,
      description: "Invoice from journal",
      confirmed: true,
      included: true,
      parentEmailProvenance: { parentName: "Notice.eml", parentSha256: parentHash, childIdentity: "child-id", childSha256: childHash },
    }],
    statements: [{ id: "statement", sourceId: "statement", sourceSha256: statementHash, witnessName: "Witness", witnessInitials: "WS" }],
    sources: [{ id: "saved-project-archive", role: "project", name: "Saved.bundle-project", path: "C:\\Saved.bundle-project", sha256: "9".repeat(64), size: 100 }],
  };
  const merged = mergeRecoveryProjectDeltas(opened, payload);
  const child = merged.opened.snapshot.candidates.find((candidate) => candidate.parentEmailProvenance?.childSha256 === childHash);
  assert.equal(child?.id, "stale-child-id");
  assert.equal(child?.confirmed, true);
  assert.equal(child?.description, "Invoice from journal");
  assert.equal(merged.opened.snapshot.candidates.filter((candidate) => candidate.parentEmailProvenance?.childSha256 === childHash).length, 1);
});

test("rejects multiple statements because one project is one witness exhibit bundle", async () => {
  const { analyseBundleStatements } = await import("../app/lib/bundle-engine.ts");
  const statementUrl = new URL("01_Witness_Statement_Amelia_Hart.docx", fixtureRoot);
  const evidenceNames = ["Executed_Supply_Agreement_2026-02-01.pdf", "PO_NRL-1047.pdf"];
  const [one, two] = await Promise.all([fixtureFile(statementUrl), fixtureFile(statementUrl)]);
  const evidence = await Promise.all(evidenceNames.map((name) => fixtureFile(new URL(name, evidenceRoot))));
  await assert.rejects(analyseBundleStatements([
    { id: "first", file: one, witnessName: "Amelia Hart", witnessInitials: "AH" },
    { id: "second", file: two, witnessName: "Amelia Hart", witnessInitials: "AH" },
  ], evidence), /one witness statement/i);
});

test("rejects multiple witness statements rather than grouping them", async () => {
  const { analyseBundleStatements } = await import("../app/lib/bundle-engine.ts");
  const statementUrl = new URL("01_Witness_Statement_Amelia_Hart.docx", fixtureRoot);
  const [one, two] = await Promise.all([fixtureFile(statementUrl), fixtureFile(statementUrl)]);
  const evidence = [await fixtureFile(new URL("Executed_Supply_Agreement_2026-02-01.pdf", evidenceRoot))];
  await assert.rejects(analyseBundleStatements([
    { id: "amelia", file: one, witnessName: "Amelia Hart", witnessInitials: "AH" },
    { id: "ben", file: two, witnessName: "Ben Brown", witnessInitials: "BB" },
  ], evidence), /one witness statement/i);
});
