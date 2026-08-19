import assert from "node:assert/strict";
import { File } from "node:buffer";
import test from "node:test";

globalThis.File = File;

function snapshot(sourceSha256, pdfSha256) {
  const confirmation = { pdfSha256, confirmedAt: "2026-08-13T12:00:00.000Z" };
  return {
    schemaVersion: 8,
    templateReviews: [{
      slot: "cover",
      sourceId: "template-cover",
      renderedSourceId: "template-rendered-cover",
      sourceFormat: "docx",
      sourceSha256,
      pdfSha256,
      reviewState: {
        matterReview: { sourceName: "Cover.rendered.pdf", pdfSha256, exactByteLength: 8, pageCount: 1, extractedCharacterCount: 20, textReliability: "reliable", requiresVisualConfirmation: false, notice: "Review", matterNumbers: [], partyNames: [], forums: [], matterTitles: [], placeholders: [] },
        appearanceConfirmation: confirmation,
        matterConfirmation: confirmation,
      },
    }],
  };
}

test("restores template confirmations only with the exact source and rendered PDF hashes", async () => {
  const { restoreProjectTemplates } = await import("../app/lib/template-persistence.ts");
  const sourceHash = "a".repeat(64);
  const renderedHash = "b".repeat(64);
  const sources = [
    { id: "template-cover", role: "template", name: "Cover.docx", sha256: sourceHash, file: new File(["word"], "Cover.docx") },
    { id: "template-rendered-cover", role: "template-rendered", name: "Cover.rendered.pdf", sha256: renderedHash, file: new File(["%PDF-rendered"], "Cover.rendered.pdf") },
  ];
  const restored = restoreProjectTemplates(snapshot(sourceHash, renderedHash), sources);
  assert.equal(restored.length, 1);
  assert.equal(restored[0].pdfSha256, renderedHash);
  assert.equal(restored[0].reviewState?.matterConfirmation?.pdfSha256, renderedHash);
  assert.equal(restored[0].reviewState?.appearanceConfirmation?.pdfSha256, renderedHash);

  const changedSource = restoreProjectTemplates(snapshot(sourceHash, renderedHash), [{ ...sources[0], sha256: "c".repeat(64) }, sources[1]]);
  assert.equal(changedSource[0].reviewState, undefined, "replacing the Word source invalidates every approval");

  const changedRendered = restoreProjectTemplates(snapshot(sourceHash, renderedHash), [sources[0], { ...sources[1], sha256: "d".repeat(64) }]);
  assert.equal(changedRendered[0].reviewState, undefined, "reconverting to different PDF bytes invalidates every approval");

  const missingRendered = restoreProjectTemplates(snapshot(sourceHash, renderedHash), [sources[0]]);
  assert.equal(missingRendered[0].pdfFile, undefined);
  assert.equal(missingRendered[0].reviewState, undefined, "a saved Word approval cannot survive without its reviewed rendered PDF");
});

test("a supplied PDF is its own exact reviewed artifact", async () => {
  const { restoreProjectTemplates } = await import("../app/lib/template-persistence.ts");
  const hash = "e".repeat(64);
  const confirmation = { pdfSha256: hash, confirmedAt: "2026-08-13T12:00:00.000Z" };
  const project = {
    schemaVersion: 8,
    templateReviews: [{ slot: "index", sourceId: "template-index", sourceFormat: "pdf", sourceSha256: hash, pdfSha256: hash, reviewState: { matterReview: { sourceName: "Index.pdf", pdfSha256: hash, exactByteLength: 8, pageCount: 1, extractedCharacterCount: 0, textReliability: "none", requiresVisualConfirmation: true, notice: "Visual review", matterNumbers: [], partyNames: [], forums: [], matterTitles: [], placeholders: [] }, matterConfirmation: confirmation } }],
  };
  const file = new File(["%PDF-index"], "Index.pdf");
  const [restored] = restoreProjectTemplates(project, [{ id: "template-index", role: "template", name: file.name, sha256: hash, file }]);
  assert.equal(restored.pdfFile, file);
  assert.equal(restored.reviewState?.matterConfirmation?.pdfSha256, hash);
});
