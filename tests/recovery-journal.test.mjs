import assert from "node:assert/strict";
import { createHash, randomUUID } from "node:crypto";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { RecoveryStore, validateRecoveryPayload } = require("../electron/recovery-journal.cjs");

test("recovery journal validates, writes atomically, rejects stale revisions and hash-verifies sources", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "exhibit-recovery-test-"));
  try {
    const sourcePath = path.join(root, "source.pdf");
    await writeFile(sourcePath, "source bytes");
    const sha256 = createHash("sha256").update("source bytes").digest("hex");
    const payload = { project: { name: "Test" }, candidates: [], finalOrder: ["group-1"], layout: {}, pagination: {}, resolutions: [], statements: [], templates: [], fingerprint: null, pageSizeChoices: { "source-1": "keep-original" }, sources: [{ id: "source-1", role: "evidence", name: "source.pdf", path: sourcePath, sha256, size: 12 }] };
    const store = new RecoveryStore(path.join(root, "journal"));
    const recoveryId = randomUUID();
    await store.write(recoveryId, 1, payload);
    assert.deepEqual((await store.load(recoveryId)).payload.pageSizeChoices, { "source-1": "keep-original" });
    assert.deepEqual(await store.status(), { available: true, stored: true, recoveryId, revision: 1, projectName: "Test" });
    await assert.rejects(store.write(recoveryId, 1, payload), /stale/i);
    await Promise.all([2, 3, 4, 5, 6].map((revision) => store.write(recoveryId, revision, { ...payload, project: { name: `Test ${revision}` } })));
    assert.equal((await store.load(recoveryId)).revision, 6, "rapid recovery updates are serialised instead of failing on Windows replacement semantics");
    assert.equal(new TextDecoder().decode((await store.readSource(recoveryId, "source-1")).bytes), "source bytes");
    await writeFile(sourcePath, "changed bytes");
    await assert.rejects(store.readSource(recoveryId, "source-1"), /changed/i);
    await store.markClean(recoveryId, 6);
    assert.deepEqual(await store.status(), { available: false, stored: true });
    assert.equal((await store.load(recoveryId)).dirty, false);
    await store.clearAll();
    assert.deepEqual(await store.status(), { available: false, stored: false });
    await assert.rejects(store.load(recoveryId), /unavailable/i);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("recovery payload rejects duplicate IDs, invalid hashes and renderer journal paths", () => {
  const source = { id: "source-1", role: "evidence", name: "a.pdf", path: path.resolve("a.pdf"), sha256: "a".repeat(64), size: 1 };
  assert.throws(() => validateRecoveryPayload({ sources: [source, source], candidates: [], finalOrder: [], resolutions: [] }), /duplicate source/i);
  assert.throws(() => validateRecoveryPayload({ sources: [{ ...source, sha256: "bad" }], candidates: [], finalOrder: [], resolutions: [] }), /validation/i);
  assert.throws(() => validateRecoveryPayload({ sources: [source], candidates: [], finalOrder: "not-an-array", resolutions: [] }), /legacy final order is invalid/i);
  assert.throws(() => validateRecoveryPayload({ sources: [source], candidates: [], finalOrder: [""], resolutions: [] }), /invalid or duplicate order/i);
  assert.throws(() => validateRecoveryPayload({ sources: [source], candidates: [], finalOrder: [], resolutions: [], journalPath: path.resolve("owned-by-renderer") }), /unsupported field/i);
});

test("recovery payload accepts non-A4 page-size choices", () => {
  const source = { id: "source-1", role: "evidence", name: "a.pdf", path: path.resolve("a.pdf"), sha256: "a".repeat(64), size: 1 };
  const payload = validateRecoveryPayload({ sources: [source], candidates: [], finalOrder: [], resolutions: [], pageSizeChoices: { "source-1": "convert-to-a4" } });
  assert.deepEqual(payload.pageSizeChoices, { "source-1": "convert-to-a4" });
  assert.throws(() => validateRecoveryPayload({ sources: [source], candidates: [], finalOrder: [], resolutions: [], pageSizeChoices: { "source-1": "shrink-and-crop" } }), /page-size choice is invalid/i);
});

test("recovery payload validates bounded schema-8 arrangements", () => {
  const source = { id: "source-1", role: "evidence", name: "a.pdf", path: path.resolve("a.pdf"), sha256: "a".repeat(64), size: 1 };
  const arrangement = { version: 1, nodes: [{ type: "section", id: "s", heading: "Agreements", exhibits: [{ type: "exhibit", exhibitId: "a" }] }, { type: "exhibit", exhibitId: "b" }] };
  const payload = validateRecoveryPayload({ sources: [source], candidates: [], arrangement, resolutions: [] });
  assert.deepEqual(payload.arrangement, arrangement);
  assert.equal("finalOrder" in payload, false);
  assert.throws(() => validateRecoveryPayload({ sources: [source], candidates: [], arrangement, finalOrder: ["a", "b"], resolutions: [] }), /both arrangement and legacy final order/i);
  assert.throws(() => validateRecoveryPayload({ sources: [source], candidates: [], arrangement: { version: 1, nodes: [{ type: "exhibit", exhibitId: "a" }, { type: "exhibit", exhibitId: "a" }] }, resolutions: [] }), /duplicate exhibit/i);
  assert.throws(() => validateRecoveryPayload({ sources: [source], candidates: [], arrangement: { version: 1, nodes: [{ type: "section", id: "s", heading: "", exhibits: [] }] }, resolutions: [] }), /invalid or duplicate section/i);
});

test("recovery payload rejects template confirmations that are not bound to the reviewed PDF", () => {
  const source = { id: "source-1", role: "evidence", name: "a.pdf", path: path.resolve("a.pdf"), sha256: "a".repeat(64), size: 1 };
  const review = {
    slot: "cover", sourceId: "template-cover", sourceFormat: "pdf", sourceSha256: "b".repeat(64), pdfSha256: "c".repeat(64),
    reviewState: {
      matterReview: { sourceName: "Cover.pdf", pdfSha256: "c".repeat(64), exactByteLength: 100, pageCount: 1, extractedCharacterCount: 10, textReliability: "reliable", requiresVisualConfirmation: true, notice: "Review exact PDF", matterNumbers: [], partyNames: [], forums: [], matterTitles: [], placeholders: [] },
      matterConfirmation: { pdfSha256: "d".repeat(64), confirmedAt: "2026-08-13T00:00:00.000Z" },
    },
  };
  assert.throws(() => validateRecoveryPayload({ sources: [source], candidates: [], arrangement: { version: 1, nodes: [] }, resolutions: [], templateReviews: [review] }), /template confirmation is invalid/i);
});

test("recovery payload accepts reviewer-corrected matter details bound to the reviewed PDF", () => {
  const source = { id: "source-1", role: "evidence", name: "a.pdf", path: path.resolve("a.pdf"), sha256: "a".repeat(64), size: 1 };
  const review = {
    slot: "cover", sourceId: "template-cover", sourceFormat: "pdf", sourceSha256: "b".repeat(64), pdfSha256: "c".repeat(64),
    reviewState: {
      matterReview: { sourceName: "Cover.pdf", pdfSha256: "c".repeat(64), exactByteLength: 100, pageCount: 1, extractedCharacterCount: 10, textReliability: "reliable", requiresVisualConfirmation: true, notice: "Review exact PDF", matterNumbers: [], partyNames: [], forums: [], matterTitles: [], placeholders: [] },
      matterConfirmation: { pdfSha256: "c".repeat(64), confirmedAt: "2026-08-14T00:00:00.000Z", partyNames: ["Northbridge Renewables Limited"], patches: [{ findingId: "party-name:1:54.00:780.00:NORTHBRIDGE RENEWABLES LIMITED", value: "Northbridge Renewables Limited" }] },
    },
  };
  const payload = validateRecoveryPayload({ sources: [source], candidates: [], arrangement: { version: 1, nodes: [] }, resolutions: [], templateReviews: [review] });
  assert.deepEqual(payload.templateReviews[0].reviewState.matterConfirmation.partyNames, ["Northbridge Renewables Limited"]);
  assert.equal(payload.templateReviews[0].reviewState.matterConfirmation.patches[0].value, "Northbridge Renewables Limited");
});

test("closing or reloading cannot silently mark an unsaved recovery journal clean", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "exhibit-recovery-restart-test-"));
  try {
    const sourcePath = path.join(root, "source.pdf");
    await writeFile(sourcePath, "source bytes");
    const payload = { project: { name: "Unsaved review" }, candidates: [], finalOrder: [], layout: {}, pagination: {}, resolutions: [], statements: [], templates: [], fingerprint: null, pageSizeChoices: {}, sources: [{ id: "source-1", role: "evidence", name: "source.pdf", path: sourcePath, sha256: createHash("sha256").update("source bytes").digest("hex"), size: 12 }] };
    const recoveryId = randomUUID();
    await new RecoveryStore(path.join(root, "journal")).write(recoveryId, 1, payload);

    const afterRestart = new RecoveryStore(path.join(root, "journal"));
    assert.deepEqual(await afterRestart.status(), { available: true, stored: true, recoveryId, revision: 1, projectName: "Unsaved review" });
    assert.equal((await afterRestart.load(recoveryId)).dirty, true, "re-instantiation models app close/reload without altering the dirty flag");

    const component = await readFile(new URL("../app/BundleBuilder.tsx", import.meta.url), "utf8");
    assert.doesNotMatch(component, /beforeunload|markRecoveryCleanOnExit/, "renderer has no close/reload path that suppresses crash recovery");
    await afterRestart.markClean(recoveryId, 1);
    assert.deepEqual(await new RecoveryStore(path.join(root, "journal")).status(), { available: false, stored: true }, "only an explicit successful save marks recovery clean");
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("a corrupt recovery journal remains visible to the deletion control", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "exhibit-recovery-corrupt-test-"));
  try {
    const journalRoot = path.join(root, "journal");
    const store = new RecoveryStore(journalRoot);
    await mkdir(journalRoot, { recursive: true });
    await writeFile(store.journalPath, "{not valid json");
    const status = await store.status();
    assert.equal(status.stored, true);
    assert.equal(status.corrupt, true);
    assert.match(status.issue, /delete it from Local recovery data/i);
    await store.clearAll();
    assert.deepEqual(await store.status(), { available: false, stored: false });
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("journal collection skips reconstructed email children and treats pathless files as empty", async () => {
  const [builder, preload] = await Promise.all([
    readFile(new URL("../app/BundleBuilder.tsx", import.meta.url), "utf8"),
    readFile(new URL("../electron/preload.cjs", import.meta.url), "utf8"),
  ]);
  assert.match(builder, /if \(evidence\.derivedFromEmail\) continue/);
  assert.match(builder, /try \{\s*return desktop\.sourcePath\(file\) \|\| ""/);
  assert.match(preload, /try \{\s*return webUtils\.getPathForFile\(file\) \|\| ""/);
  assert.match(preload, /catch \{\s*return "";/);
});
