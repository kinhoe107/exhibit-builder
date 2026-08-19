import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import test from "node:test";
import { PDFDocument, StandardFonts } from "pdf-lib";
import {
  compareTemplateMatterReviews,
  reviewTemplateMatterPdf,
} from "../app/lib/template-matter-review.ts";

async function textPdf(lines) {
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([595.28, 841.89]);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  lines.forEach((line, index) => page.drawText(line, { x: 54, y: 780 - index * 28, size: 12, font }));
  return new Uint8Array(await pdf.save());
}

test("extracts unverified matter details and binds them to the exact PDF bytes", async () => {
  const bytes = await textPdf([
    "IN THE LONDON COURT OF INTERNATIONAL ARBITRATION",
    "Case No. LCIA-2026-0042",
    "Northbridge Renewables Limited v Meridian Components Limited",
    "[HEARING DATE]",
  ]);
  const review = await reviewTemplateMatterPdf(bytes, "Cover.pdf");

  assert.equal(review.pdfSha256, createHash("sha256").update(bytes).digest("hex"));
  assert.equal(review.exactByteLength, bytes.byteLength);
  assert.equal(review.textReliability, "reliable");
  assert.equal(review.requiresVisualConfirmation, false);
  assert.match(review.notice, /unverified/i);
  assert.equal(review.matterNumbers[0]?.value, "LCIA-2026-0042");
  assert.deepEqual(review.partyNames.map((finding) => finding.value), [
    "Meridian Components Limited",
    "Northbridge Renewables Limited",
  ]);
  assert.match(review.forums[0]?.value ?? "", /LONDON COURT OF INTERNATIONAL ARBITRATION/);
  assert.match(review.matterTitles[0]?.value ?? "", /Northbridge Renewables Limited v Meridian Components Limited/);
  assert.equal(review.placeholders[0]?.value, "[HEARING DATE]");
  assert.ok([
    ...review.matterNumbers,
    ...review.partyNames,
    ...review.forums,
    ...review.matterTitles,
    ...review.placeholders,
  ].every((finding) => finding.unverified && finding.id && finding.geometry && finding.geometry.pageNumber === 1 && finding.geometry.fontSize > 0));
});

test("extracts party names after IN THE ARBITRATION BETWEEN without role suffixes", async () => {
  const bytes = await textPdf([
    "IN THE ARBITRATION BETWEEN",
    "Northbridge Renewables Limited",
    "AND",
    "Meridian Components Limited",
  ]);
  const review = await reviewTemplateMatterPdf(bytes, "SIAC-cover.pdf");
  assert.deepEqual(review.partyNames.map((finding) => finding.value).sort((left, right) => left.localeCompare(right, "en-GB")), [
    "Meridian Components Limited",
    "Northbridge Renewables Limited",
  ]);
  assert.ok(review.partyNames.every((finding) => finding.id && finding.geometry));
});

test("requires visual confirmation when the exact PDF has no reliable text", async () => {
  const blank = await textPdf([]);
  const review = await reviewTemplateMatterPdf(blank, "Scanned-or-outlined-cover.pdf");

  assert.equal(review.textReliability, "none");
  assert.equal(review.requiresVisualConfirmation, true);
  assert.match(review.notice, /could not be read reliably/i);
  assert.match(review.notice, /visually check the exact PDF preview/i);
  assert.deepEqual(review.matterNumbers, []);
  assert.deepEqual(review.partyNames, []);
});

test("cross-template discrepancy evidence is deterministic and hash-bound", async () => {
  const firstBytes = await textPdf(["Case No. ARB-100", "Alpha Limited v Beta Limited"]);
  const secondBytes = await textPdf(["CASE NUMBER: ARB-200", "Alpha Limited v Gamma Limited"]);
  const first = await reviewTemplateMatterPdf(firstBytes, "Cover.pdf");
  const second = await reviewTemplateMatterPdf(secondBytes, "Index.pdf");
  const references = [
    { templateId: "index", role: "index", sourceName: "Index.pdf", review: second },
    { templateId: "cover", role: "cover", sourceName: "Cover.pdf", review: first },
  ];

  const forward = compareTemplateMatterReviews(references);
  const reverse = compareTemplateMatterReviews([...references].reverse());
  assert.deepEqual(forward, reverse);
  assert.deepEqual(forward.map((item) => item.field), ["matter-number", "party-name", "matter-title"]);
  assert.ok(forward.every((item) => item.unverified && item.evidence.every((entry) => /^[a-f0-9]{64}$/.test(entry.pdfSha256))));
  assert.equal(forward[0].evidence[0].role, "cover");
});

test("normalization prevents a discrepancy from casing and spacing alone", async () => {
  const bytes = await textPdf(["Case No. ARB-100"]);
  const review = await reviewTemplateMatterPdf(bytes, "Cover.pdf");
  const same = structuredClone(review);
  same.matterNumbers[0].value = " arb - 100 ";
  same.matterNumbers[0].normalizedValue = review.matterNumbers[0].normalizedValue;
  assert.deepEqual(compareTemplateMatterReviews([
    { templateId: "cover", role: "cover", sourceName: "Cover.pdf", review },
    { templateId: "index", role: "index", sourceName: "Index.pdf", review: same },
  ]), []);
});

test("reviewer corrections are compared instead of the original misread list", async () => {
  const { matterDraftFromReview, matterValuesFromConfirmation, parseMatterDraft, parseMatterListDraft } = await import("../app/lib/template-matter-review.ts");
  const firstBytes = await textPdf(["Case No. ARB-100", "Alpha Limited v Beta Limited"]);
  const secondBytes = await textPdf(["CASE NUMBER: ARB-200", "Alpha Limited v Gamma Limited"]);
  const first = await reviewTemplateMatterPdf(firstBytes, "Cover.pdf");
  const second = await reviewTemplateMatterPdf(secondBytes, "Index.pdf");
  const corrected = parseMatterListDraft({
    matterNumbers: "ARB-200",
    partyNames: "Alpha Limited\nGamma Limited",
    forums: "",
    matterTitles: "Alpha Limited v Gamma Limited",
  });
  assert.deepEqual(compareTemplateMatterReviews([
    { templateId: "cover", role: "cover", sourceName: "Cover.pdf", review: first, confirmedValues: corrected },
    { templateId: "index", role: "index", sourceName: "Index.pdf", review: second },
  ]), []);
  assert.equal(matterValuesFromConfirmation({ pdfSha256: "a".repeat(64), confirmedAt: "2026-08-14T00:00:00.000Z" }), undefined);
  const draft = matterDraftFromReview(first);
  const party = draft.occurrences.find((occurrence) => occurrence.originalValue === "Beta Limited");
  assert.ok(party);
  party.value = "Gamma Limited";
  const parsed = parseMatterDraft(draft);
  assert.ok(parsed.patches.some((patch) => patch.findingId === party.findingId && patch.value === "Gamma Limited"));
  assert.ok(parsed.values.partyNames.includes("Gamma Limited"));
});

test("matter draft ignores patches unless confirmation hash matches the reviewed PDF", async () => {
  const { matterDraftFromReview } = await import("../app/lib/template-matter-review.ts");
  const bytes = await textPdf(["Alpha Limited v Beta Limited"]);
  const review = await reviewTemplateMatterPdf(bytes, "Cover.pdf");
  const baseline = matterDraftFromReview(review);
  const party = baseline.occurrences.find((occurrence) => occurrence.originalValue === "Beta Limited");
  assert.ok(party);
  const mismatched = matterDraftFromReview(review, {
    pdfSha256: "0".repeat(64),
    patches: [{ findingId: party.findingId, value: "Gamma Limited" }],
  });
  assert.equal(mismatched.occurrences.find((occurrence) => occurrence.findingId === party.findingId)?.value, "Beta Limited");
  const matched = matterDraftFromReview(review, {
    pdfSha256: review.pdfSha256,
    patches: [{ findingId: party.findingId, value: "Gamma Limited" }],
  });
  assert.equal(matched.occurrences.find((occurrence) => occurrence.findingId === party.findingId)?.value, "Gamma Limited");
});

test("rejects bytes that are not the rendered PDF artifact", async () => {
  await assert.rejects(
    reviewTemplateMatterPdf(new TextEncoder().encode("not a PDF"), "Cover.pdf"),
    /not a readable PDF/,
  );
});
