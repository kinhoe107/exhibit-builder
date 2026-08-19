import assert from "node:assert/strict";
import { createHash, webcrypto } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename } from "node:path";
import { File } from "node:buffer";
import test from "node:test";
import { degrees, PDFDocument, PDFHexString, PDFName, StandardFonts } from "pdf-lib";
import JSZip from "jszip";

globalThis.File = File;
globalThis.crypto ??= webcrypto;

async function reviewedTemplate(slot, file, options = {}) {
  const sourceFormat = options.sourceFormat ?? (file.name.toLowerCase().endsWith(".pdf") ? "pdf" : "docx");
  const pdfFile = options.pdfFile ?? (sourceFormat === "pdf" ? file : undefined);
  if (!pdfFile) throw new Error("A reviewed Word-template fixture needs its exact converted PDF.");
  const sourceSha256 = createHash("sha256").update(new Uint8Array(await file.arrayBuffer())).digest("hex");
  const pdfSha256 = createHash("sha256").update(new Uint8Array(await pdfFile.arrayBuffer())).digest("hex");
  const confirmedAt = "2026-08-13T10:00:00.000Z";
  const confirmation = { pdfSha256, confirmedAt };
  return {
    slot,
    file,
    sha256: sourceSha256,
    sourceFormat,
    pdfFile,
    pdfSha256,
    reviewState: {
      matterReview: {
        sourceName: pdfFile.name,
        pdfSha256,
        exactByteLength: pdfFile.size,
        pageCount: 1,
        extractedCharacterCount: 0,
        textReliability: "none",
        requiresVisualConfirmation: true,
        notice: "Fixture reviewed visually.",
        matterNumbers: [],
        partyNames: [],
        forums: [],
        matterTitles: [],
        placeholders: options.placeholders ?? [],
      },
      ...(sourceFormat === "pdf" ? {} : { appearanceConfirmation: confirmation }),
      matterConfirmation: confirmation,
      ...((options.placeholders ?? []).length ? { placeholderConfirmation: confirmation } : {}),
    },
  };
}

async function readPageLabels(bytes) {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const document = await pdfjs.getDocument({ data: new Uint8Array(bytes) }).promise;
  const labels = await document.getPageLabels();
  await document.destroy();
  return labels;
}

const workbookExportRequests = [];
async function workbookPrintTestDouble(_file, sheets) {
  workbookExportRequests.push(sheets.map((sheet) => ({ ...sheet })));
  return Promise.all(sheets.map(async (sheet) => {
    const pdf = await PDFDocument.create();
    pdf.addPage([841.89, 595.28]).drawText(`Payment Ledger - ${sheet.name} - Range ${sheet.range}`);
    return { name: sheet.name, range: sheet.range, bytes: await pdf.save() };
  }));
}

test("places finalisation drag drops after a lower target and before a higher target", async () => {
  const { reorderGroupForDrop } = await import("../app/lib/exhibit-groups.ts");
  const groups = ["A", "B", "C"].map((id) => ({ canonical: { id } }));
  assert.deepEqual(reorderGroupForDrop(groups, "A", "B").map((group) => group.canonical.id), ["B", "A", "C"]);
  assert.deepEqual(reorderGroupForDrop(groups, "B", "C").map((group) => group.canonical.id), ["A", "C", "B"]);
  assert.deepEqual(reorderGroupForDrop(groups, "C", "A").map((group) => group.canonical.id), ["C", "A", "B"]);
});

test("review item numbers remain globally sequential and stable when cards are filtered", async () => {
  const { reviewItemNumbers } = await import("../app/lib/exhibit-groups.ts");
  const reviewCandidates = [
    { id: "item-8", provisionalNumber: 8 },
    { id: "unexpected-reset", provisionalNumber: 1 },
    { id: "item-9", provisionalNumber: 9 },
  ];
  const numbers = reviewItemNumbers(reviewCandidates);
  assert.deepEqual(reviewCandidates.map((candidate) => numbers.get(candidate.id)), [1, 2, 3]);
  assert.deepEqual([reviewCandidates[0], reviewCandidates[2]].map((candidate) => numbers.get(candidate.id)), [1, 3], "filtering hides cards without renumbering the retained cards");
});

test("builds a manually added uncited exhibit without inventing statement references", async () => {
  const { analyseFiles, buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const guidedRoot = new URL("../public/guided-sample/", import.meta.url);
  const statement = await fixtureFile(new URL("01_GUIDED_SAMPLE_Witness_Statement.docx", guidedRoot));
  const source = await fixtureFile(new URL("06_SAMPLE_Unreferenced_Checklist.pdf", guidedRoot));
  const analysis = await analyseFiles(statement, [source]);
  const evidence = analysis.evidence[0];
  const timestamp = "2026-08-08T10:00:00.000Z";
  const manual = { id: "manual-checklist", mark: "EX 1", provisionalNumber: 1, description: "Sample unreferenced checklist", date: "6 August 2026", paragraph: 0, citation: "", citationResolution: "none", discoverySignals: ["Manually added by reviewer"], evidenceId: evidence.id, confidence: 100, rationale: "Reviewer intentionally added an uncited exhibit", included: true, confirmed: true, exhibitInitials: "EX", exhibitSequence: 1, witnessInitials: "EX", witnessKey: "general-exhibits::EX", sequenceOrder: 1000, manualAddition: true, manualAddedAt: timestamp, manualWarningAcknowledgedAt: timestamp };
  const build = await buildBundle(analysis, [manual]);
  assert.equal(build.records.length, 1);
  assert.equal(build.records[0].statementParagraph, null);
  assert.deepEqual(build.records[0].statementReferences, []);
  assert.equal(build.records[0].citationStatus, "not-cited-manual-addition");
  assert.equal(build.records[0].sourceHash, evidence.sha256);
  const record = build.records[0];
  const pageRange = record.exhibitPageLabelStart === record.exhibitPageLabelEnd
    ? record.exhibitPageLabelStart
    : `${record.exhibitPageLabelStart}-${record.exhibitPageLabelEnd}`;
  const pageWord = record.exhibitPageLabelStart === record.exhibitPageLabelEnd ? "page" : "pages";
  assert.deepEqual(buildStatementUpdateSuggestions(build.records).map((item) => item.line), [
    "Uncited exhibits — no statement reference",
    `${record.exhibitNumber}. ${record.description} — ${pageWord} ${pageRange}`,
  ]);
  assert.deepEqual(build.manifest.manualAdditions, [{ exhibitNumber: 1, description: manual.description, documentDate: manual.date, sourceFile: evidence.name, sourceSha256: evidence.sha256, addedAt: timestamp, warningAcknowledgedAt: timestamp, citationStatus: "not-cited-manual-addition" }]);
  const reopened = await PDFDocument.load(build.bytes);
  assert.ok(reopened.catalog.get(PDFName.of("Outlines")));
});

const fixtureRoot = new URL("./fixtures/core/", import.meta.url);
const evidenceRoot = new URL("./fixtures/core/Evidence_Inbox/", import.meta.url);
const outputRoot = new URL("../output/", import.meta.url);
const outputPdfRoot = new URL("../output/pdf/", import.meta.url);

async function fixtureFile(url) {
  return new File([await readFile(url)], basename(url.pathname));
}

test("extracts, matches, reconciles and builds the synthetic exhibit pack", async () => {
  const { analyseFiles, buildBundle } = await import(
    "../app/lib/bundle-engine.ts"
  );
  const truth = JSON.parse(
    await readFile(new URL("Ground_Truth_Manifest.json", fixtureRoot), "utf8"),
  );
  const statementUrl = new URL(
    "01_Witness_Statement_Amelia_Hart.docx",
    fixtureRoot,
  );
  const statementBytesBefore = await readFile(statementUrl);
  const statementHashBefore = createHash("sha256")
    .update(statementBytesBefore)
    .digest("hex");
  const statement = await fixtureFile(statementUrl);
  const evidenceNames = [
    "Executed_Supply_Agreement_2026-02-01.pdf",
    "PO_NRL-1047.pdf",
    "Email_Orchard_Vale_delivery_dates.eml",
    "Progress_Meeting_Minutes_2026-03-11.docx",
    "Email_processor_supply_issue_18_March.eml",
    "Northbridge_Notice_of_Delay_2026-03-25.docx",
    "Email_thermal_shutdowns_16_April.eml",
    "Verity_Inspection_Report_VES-2261.pdf",
    "Northbridge_Termination_Letter_2026-04-30.pdf",
    "Apex_Invoice_AC-7782.pdf",
    "UNREFERENCED_DRAFT_Status_Request.pdf",
  ];
  const evidence = await Promise.all(
    evidenceNames.map((name) => fixtureFile(new URL(name, evidenceRoot))),
  );

  const result = await analyseFiles(statement, evidence);
  assert.equal(result.caseTitle, truth.case);
  assert.equal(result.candidates.length, 10);
  assert.deepEqual(
    result.candidates.map((candidate) => candidate.paragraph),
    [12, 14, 15, 18, 21, 23, 28, 31, 34, 36],
  );
  assert.deepEqual(
    result.candidates.map((candidate) => candidate.mark),
    Array.from({ length: 10 }, (_, index) => `AH ${index + 1}`),
  );
  assert.match(result.candidates[0].context ?? "", /Following paragraph \d+:/, "review context retains the following witness-statement paragraph");
  assert.doesNotMatch(result.candidates[0].context ?? "", new RegExp(result.candidates[0].citation.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")), "review context does not duplicate the cited paragraph");
  assert.ok(result.candidates[0].contextParagraphs?.some((item) => item.position === "following" && item.text.length > 0), "structured review context retains a following paragraph");
  assert.ok(result.candidates[0].contextParagraphs?.every((item) => item.text !== result.candidates[0].citation), "structured context excludes the cited paragraph");

  const expectedByMark = new Map(
    truth.exhibits.map((item) => [item.exhibit_mark, item]),
  );
  for (const candidate of result.candidates) {
    const expected = expectedByMark.get(`AH${candidate.provisionalNumber}`);
    assert.ok(expected, `Unexpected exhibit ${candidate.mark}`);
    assert.equal(candidate.paragraph, expected.statement_paragraph);
    assert.equal(candidate.date, expected.document_date);
    assert.equal(candidate.confirmed, false);
    assert.ok(candidate.discoverySignals.length > 0);
    const matched = result.evidence.find(
      (item) => item.id === candidate.evidenceId,
    );
    assert.equal(matched?.name, expected.source_filename);
    assert.equal(candidate.confidence, 99);
  }

  assert.deepEqual(
    result.unreferenced.map((item) => item.name),
    ["UNREFERENCED_DRAFT_Status_Request.pdf"],
  );

  await assert.rejects(
    analyseFiles(statement, [
      new File(["This is not a PDF."], "Unreadable_Evidence.pdf"),
    ]),
    /Could not read "Unreadable_Evidence\.pdf":/i,
  );

  await assert.rejects(
    buildBundle(result, result.candidates),
    /Confirm every included provisional number/i,
  );

  const confirmedCandidates = result.candidates.map((candidate) => ({
    ...candidate,
    confirmed: true,
  }));
  const build = await buildBundle(result, confirmedCandidates);
  const reopened = await PDFDocument.load(build.bytes);
  assert.equal(reopened.getPageCount(), build.pageCount);
  assert.equal(build.records.length, 10);
  for (const page of reopened.getPages()) {
    const { width, height } = page.getSize();
    assert.ok(Math.abs(width - 595.28) < 0.02, `Non-A4 width: ${width}`);
    assert.ok(Math.abs(height - 841.89) < 0.02, `Non-A4 height: ${height}`);
  }
  assert.ok(reopened.catalog.get(PDFName.of("Outlines")));
  assert.equal(reopened.getPages()[1].node.Annots()?.size(), 10);
  assert.ok(
    build.checks.every(
      (check) =>
        check.status !== "warning" || check.label === "Unreferenced evidence",
    ),
  );
  assert.equal(build.manifest.statement.modified, false);
  assert.equal("profile" in build.manifest, false, "downloaded manifests do not expose a profile");
  assert.deepEqual(
    build.records.map((record) => record.statementParagraph),
    [12, 14, 15, 18, 21, 23, 28, 31, 34, 36],
  );

  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const rendered = await pdfjs.getDocument({
    data: new Uint8Array(build.bytes),
  }).promise;
  const indexPage = await rendered.getPage(2);
  const indexText = (await indexPage.getTextContent()).items
    .map((item) => ("str" in item ? item.str : ""))
    .join(" ");
  assert.match(indexText, /Executed Supply Agreement/);
  assert.doesNotMatch(indexText, /cited at|paragraph\s+\d+/i, "the reader-facing index excludes statement cross-references");
  for (const exhibitNumber of [3, 5, 7]) {
    const record = build.records.find((item) => item.exhibitNumber === exhibitNumber);
    assert.ok(record, `Missing email record ${exhibitNumber}`);
    const page = await rendered.getPage(record.startPage);
    const content = await page.getTextContent();
    const text = content.items
      .map((item) => ("str" in item ? item.str : ""))
      .join(" ");
    assert.match(text, /From:/);
    assert.match(text, /Sent:/);
    assert.match(text, /Subject:/);
    assert.match(text, /@(?:northbridge|meridian)\.example/);
  }
  await rendered.destroy();

  await mkdir(outputPdfRoot, { recursive: true });
  await mkdir(outputRoot, { recursive: true });
  await writeFile(
    new URL("Generated_Northbridge_v_Meridian_Exhibit_Bundle.pdf", outputPdfRoot),
    build.bytes,
  );
  await writeFile(
    new URL("Generated_Build_Manifest.json", outputRoot),
    JSON.stringify(build.manifest, null, 2),
  );
  await writeFile(
    new URL("Generated_Bundle_Validation.json", outputRoot),
    JSON.stringify(
      {
        passed: true,
        pageCount: build.pageCount,
        exhibitCount: build.records.length,
        outputSha256: build.sha256,
        checks: build.checks,
      },
      null,
      2,
    ),
  );

  const statementBytesAfter = await readFile(statementUrl);
  assert.deepEqual(statementBytesAfter, statementBytesBefore);
  assert.equal(
    createHash("sha256").update(statementBytesAfter).digest("hex"),
    statementHashBefore,
  );
});

test("canonical final order controls the generated PDF and final ZIP manifests retain audit metadata", async () => {
  const { analyseFiles, buildBundle, finalizeBuildAudit } = await import("../app/lib/bundle-engine.ts");
  const { deriveExhibitGroups } = await import("../app/lib/exhibit-groups.ts");
  const statement = await fixtureFile(new URL("01_Witness_Statement_Amelia_Hart.docx", fixtureRoot));
  const evidence = await Promise.all(["Executed_Supply_Agreement_2026-02-01.pdf", "PO_NRL-1047.pdf"].map((name) => fixtureFile(new URL(name, evidenceRoot))));
  const analysis = await analyseFiles(statement, evidence);
  const candidates = analysis.candidates.slice(0, 2).map((candidate, index) => ({ ...candidate, included: true, confirmed: true, sequenceOrder: index + 1 }));
  const naturalGroups = deriveExhibitGroups(analysis, candidates);
  const canonicalOrder = naturalGroups.map((group) => group.id).reverse();
  const built = await buildBundle(analysis, candidates, { canonicalOrder, layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 4 } });
  assert.deepEqual(built.records.map((record) => record.description), candidates.map((candidate) => candidate.description).reverse());
  assert.deepEqual(built.records.map((record) => record.exhibitNumber), [1, 2]);
  assert.ok(built.volumes?.length >= 2, "the low physical limit creates administrative volumes");

  const comparison = { changed: true, categories: ["This is the first version of this bundle."], summary: "Bundle created." };
  const finalized = await finalizeBuildAudit(built, "f".repeat(64), comparison);
  const zip = await JSZip.loadAsync(finalized.volumeZipBytes);
  for (const volume of finalized.volumes) {
    const entry = zip.file(volume.fileName.replace(/\.pdf$/i, "_Manifest.json"));
    assert.ok(entry);
    const embedded = JSON.parse(await entry.async("text"));
    assert.equal(embedded.inputFingerprint, "f".repeat(64));
    assert.deepEqual(embedded.rebuildComparison, comparison);
    assert.deepEqual(embedded, JSON.parse(JSON.stringify(volume.manifest)));
  }
  assert.equal(finalized.manifest.output.sha256, finalized.volumeZipSha256);
});

test("rejects a multi-page custom index template with a clear requirement", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const statement = await fixtureFile(new URL("01_Witness_Statement_Amelia_Hart.docx", fixtureRoot));
  const evidence = [await fixtureFile(new URL("Executed_Supply_Agreement_2026-02-01.pdf", evidenceRoot))];
  const analysis = await analyseFiles(statement, evidence);
  const candidate = { ...analysis.candidates[0], included: true, confirmed: true };
  const templateDocument = await PDFDocument.create();
  templateDocument.addPage();
  templateDocument.addPage();
  const template = new File([await templateDocument.save()], "Two-page-index.pdf", { type: "application/pdf" });
  const reviewedMultiPage = await reviewedTemplate("index", template);
  await assert.rejects(
    buildBundle(analysis, [candidate], { templates: [reviewedMultiPage] }),
    /Custom index templates must contain exactly one PDF page;.*contains 2/i,
  );
  const letterDocument = await PDFDocument.create();
  letterDocument.addPage([612, 792]);
  const letterTemplate = new File([await letterDocument.save()], "Letter-index.pdf", { type: "application/pdf" });
  const reviewedLetter = await reviewedTemplate("index", letterTemplate);
  await assert.rejects(
    buildBundle(analysis, [candidate], { templates: [reviewedLetter] }),
    /fixed-layout index background must be one portrait A4 page/i,
  );
});

test("exact covers stay A4 and receive page numbers or volume labels only when chosen", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  async function letterCover() {
    const pdf = await PDFDocument.create();
    pdf.addPage([612, 792]).drawText("EXACT-COVER-MARKER", { x: 72, y: 720, size: 18 });
    return new File([await pdf.save()], "Exact-cover.pdf", { type: "application/pdf" });
  }
  async function exhibit(name, marker) {
    const pdf = await PDFDocument.create();
    pdf.addPage([595.28, 841.89]).drawText(marker, { x: 72, y: 720, size: 14 });
    const bytes = await pdf.save();
    const file = new File([bytes], name, { type: "application/pdf" });
    return { file, bytes, sha256: createHash("sha256").update(bytes).digest("hex") };
  }
  async function pageText(bytes, pageNumber) {
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    const document = await pdfjs.getDocument({ data: new Uint8Array(bytes) }).promise;
    const page = await document.getPage(pageNumber);
    const items = (await page.getTextContent()).items.filter((item) => "str" in item);
    const text = items.map((item) => item.str).join(" ");
    await document.destroy();
    return { text, items };
  }
  const coverFile = await letterCover();
  const first = await exhibit("one.pdf", "EXHIBIT-ONE");
  const second = await exhibit("two.pdf", "EXHIBIT-TWO");
  const candidate = (id, evidenceId, order) => ({ id, mark: "AH 1", provisionalNumber: order, description: `Exhibit ${id}`, date: "2026-08-14", paragraph: order, citation: `[AH1/xx]`, exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId, confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "AH", sequenceOrder: order, statementName: "Statement.docx" });
  const evidence = [
    { id: "one", file: first.file, name: first.file.name, extension: "pdf", text: "EXHIBIT-ONE", marker: null, sha256: first.sha256, pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" },
    { id: "two", file: second.file, name: second.file.name, extension: "pdf", text: "EXHIBIT-TWO", marker: null, sha256: second.sha256, pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" },
  ];
  const candidates = [candidate("one", "one", 1), candidate("two", "two", 2)];
  const analysis = { statementName: "Statement.docx", statementHash: "s".repeat(64), caseTitle: "Test", candidates, evidence, unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const cover = await reviewedTemplate("cover", coverFile);
  const exactLayout = { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 3, coverInsertion: "exact", exactCoverPageNumber: false, exactCoverVolumeLabel: false };
  const exact = await buildBundle(analysis, candidates, { templates: [cover], layout: exactLayout, pagination: { matchPdfPageOrder: true } });
  const exactBytes = exact.volumes?.[0]?.bytes ?? exact.bytes;
  const exactCoverPage = await PDFDocument.load(exactBytes);
  const exactSize = exactCoverPage.getPage(0).getSize();
  assert.ok(Math.abs(exactSize.width - 595.28) < 0.05 && Math.abs(exactSize.height - 841.89) < 0.05, "an exact cover is still fitted to A4");
  const unmarked = await pageText(exactBytes, 1);
  assert.match(unmarked.text, /EXACT-COVER-MARKER/);
  assert.doesNotMatch(unmarked.text, /Volume 1 of/);
  assert.equal(unmarked.items.some((item) => item.str === "1" && item.transform[5] < 40), false, "exact covers do not receive a page number unless chosen");
  assert.equal(exact.manifest.exhibits[0].sourceHash, first.sha256);
  assert.equal(exact.manifest.statement.modified, false);
  const marked = await buildBundle(analysis, candidates, { templates: [cover], layout: { ...exactLayout, exactCoverPageNumber: true, exactCoverVolumeLabel: true }, pagination: { matchPdfPageOrder: true } });
  const labelled = await pageText(marked.volumes?.[0]?.bytes ?? marked.bytes, 1);
  assert.match(labelled.text, /EXACT-COVER-MARKER/);
  assert.match(labelled.text, /Volume 1 of/);
  assert.equal(labelled.items.some((item) => item.str === "1" && item.transform[5] < 40), true, "an explicit exact-cover page-number choice prints the number");
});

test("renders selected XLSX sheet records into an inspectable PDF", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const statement = await fixtureFile(new URL("01_Witness_Statement_Amelia_Hart.docx", fixtureRoot));
  const workbook = await fixtureFile(new URL("../moorland/Exhibits/Payment_Ledger_to_2026-03-31.xlsx", fixtureRoot));
  const analysis = await analyseFiles(statement, [workbook]);
  const evidence = analysis.evidence[0];
  assert.ok(evidence.workbook?.sheets.some((sheet) => sheet.renderPlan.predictedPageCount >= 1));
  const candidate = { ...analysis.candidates[0], evidenceId: evidence.id, confirmed: true, included: true, mark: "AH 1" };
  workbookExportRequests.length = 0;
  const build = await buildBundle(analysis, [candidate], { workbookExporter: workbookPrintTestDouble });
  assert.equal(workbookExportRequests[0][0].range, "", "automatic mode leaves Excel's native print range intact");
  const sheetRecord = build.records.find((record) => record.workbookSheet);
  assert.ok(sheetRecord, "A selected workbook sheet creates a BundleRecord");
  assert.equal(sheetRecord.startPage >= 2, true);
  await mkdir(outputPdfRoot, { recursive: true });
  await writeFile(new URL("Generated_Xlsx_Evidence_Bundle.pdf", outputPdfRoot), build.bytes);
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const preview = await pdfjs.getDocument({ data: new Uint8Array(build.bytes) }).promise;
  const page = await preview.getPage(sheetRecord.startPage);
  const text = (await page.getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  assert.match(text, /Payment Ledger|Range/i);
  await preview.destroy();
});

test("groups confirmed repeat citations once and surfaces outstanding repeat approvals", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { deriveExhibitGroups, exhibitGroupLookup, orderExhibitGroups, pendingReviewCandidateIds } = await import("../app/lib/exhibit-groups.ts");
  const { runPreflight } = await import("../app/lib/preflight.ts");
  const { BUNDLE_PROFILES } = await import("../app/lib/bundle-types.ts");
  const statement = await fixtureFile(new URL("01_Witness_Statement_Amelia_Hart.docx", fixtureRoot));
  const source = await fixtureFile(new URL("Executed_Supply_Agreement_2026-02-01.pdf", evidenceRoot));
  const analysis = await analyseFiles(statement, [source]);
  const evidence = analysis.evidence[0];
  const first = { ...analysis.candidates[0], paragraph: 4, mark: "AH 37", provisionalNumber: 37, evidenceId: evidence.id, included: true, confirmed: true };
  const repeat = { ...analysis.candidates[1], id: "repeat-paragraph-30", paragraph: 30, mark: "AH 108", provisionalNumber: 108, evidenceId: evidence.id, included: true, confirmed: true };

  assert.equal(deriveExhibitGroups(analysis, [first, repeat]).length, 1);
  await assert.rejects(buildBundle(analysis, [first, repeat]), /Repeat source decision required/);
  await assert.rejects(buildBundle(analysis, [first, repeat], { preflight: [] }), /Repeat source decision required/, "a stale caller preflight cannot bypass repeat review");

  const confirmedSame = { ...repeat, repeatDecision: "same" };
  const grouped = deriveExhibitGroups(analysis, [first, confirmedSame]);
  assert.equal(grouped.length, 1);
  assert.deepEqual(grouped[0].references.map((reference) => reference.paragraph), [4, 30]);
  const bundle = await buildBundle(analysis, [first, confirmedSame]);
  assert.equal(bundle.records.length, 1);
  assert.deepEqual(bundle.records[0].statementReferences.map((reference) => reference.paragraph), [4, 30]);
  assert.equal(bundle.manifest.exhibits.length, 1);

  const unconfirmedRepeat = { ...confirmedSame, confirmed: false };
  const pendingRepeatGroups = deriveExhibitGroups(analysis, [first, unconfirmedRepeat]);
  assert.equal(pendingRepeatGroups[0].decisionPending, false, "the same/separate decision is already resolved");
  const pendingIds = pendingReviewCandidateIds([first, unconfirmedRepeat], pendingRepeatGroups);
  assert.ok(pendingIds.has(first.id), "an unconfirmed noncanonical repeat surfaces its canonical review card");
  const pendingLookup = exhibitGroupLookup(pendingRepeatGroups);
  assert.equal(pendingLookup.byCandidateId.get(first.id), pendingRepeatGroups[0], "a review-card candidate ID resolves to its repeat group");
  assert.equal(pendingLookup.byGroupId.get(pendingRepeatGroups[0].id), pendingRepeatGroups[0], "a final-order group ID resolves to the same repeat group");
  assert.equal(pendingLookup.byCandidateId.get(first.id)?.collisionMembers[1].id, unconfirmedRepeat.id, "the visible canonical review card retains the hidden repeat needing approval");

  const collisionCandidate = { ...first, id: pendingRepeatGroups[0].id, evidenceId: evidence.id };
  const collisionGroup = { ...pendingRepeatGroups[0], id: "another-group", canonical: collisionCandidate };
  const collisionSafeLookup = exhibitGroupLookup([pendingRepeatGroups[0], collisionGroup]);
  assert.equal(collisionSafeLookup.byGroupId.get(pendingRepeatGroups[0].id), pendingRepeatGroups[0], "group IDs cannot be overwritten by candidate IDs");
  assert.equal(collisionSafeLookup.byCandidateId.get(collisionCandidate.id), collisionGroup, "candidate IDs remain independently addressable");

  const excluded = { ...first, id: "excluded-candidate", included: false, evidenceId: null, confirmed: false };
  assert.equal(pendingReviewCandidateIds([excluded], deriveExhibitGroups(analysis, [excluded])).size, 0, "an excluded candidate is not an outstanding approval");

  const distinctEvidence = { ...evidence, id: "different-document", name: "Different agreement.pdf", sha256: "different-content-hash" };
  const nextExhibit = { ...analysis.candidates[2], id: "next-exhibit", paragraph: 31, mark: "AH 900", provisionalNumber: 900, evidenceId: distinctEvidence.id, included: true, confirmed: true };
  const sequential = await buildBundle({ ...analysis, evidence: [evidence, distinctEvidence] }, [first, confirmedSame, nextExhibit]);
  assert.deepEqual(sequential.records.map((record) => record.mark), ["EX1", "EX1"], "the review/body mark identifies the single witness exhibit bundle");
  assert.deepEqual(sequential.records.map((record) => record.exhibitNumber), [1, 2], "individual exhibit numbers are automatically contiguous despite arbitrary saved candidate numbers");

  const copiedEvidence = { ...evidence, id: "same-content-under-another-name", name: "Copy of agreement.pdf" };
  const copyAnalysis = { ...analysis, evidence: [evidence, copiedEvidence] };
  const otherFileName = { ...confirmedSame, evidenceId: copiedEvidence.id };
  assert.equal(deriveExhibitGroups(copyAnalysis, [first, otherFileName]).length, 1, "hash equality, not filename, is the repeat signal");
  const auditChecks = runPreflight(copyAnalysis, [first], BUNDLE_PROFILES[0]);
  assert.ok(auditChecks.some((check) => check.label === "Unselected duplicate physical copies"));
  const auditedBundle = await buildBundle(copyAnalysis, [first]);
  assert.deepEqual(auditedBundle.manifest.excludedFiles, [{ fileName: "Copy of agreement.pdf", sha256: evidence.sha256, reason: "Unselected duplicate physical copy of a selected source" }]);
  const differentHash = { ...copiedEvidence, sha256: "different-content-hash" };
  assert.equal(deriveExhibitGroups({ ...analysis, evidence: [evidence, differentHash] }, [first, otherFileName]).length, 2, "different content remains separate");
  const originalGroupId = deriveExhibitGroups(copyAnalysis, [first])[0].id;
  const replacementGroupId = deriveExhibitGroups(copyAnalysis, [{ ...first, evidenceId: copiedEvidence.id }])[0].id;
  assert.equal(replacementGroupId, originalGroupId, "changing the proposed source preserves the review card and final-order identity");
  const migrationGroups = deriveExhibitGroups({ ...analysis, evidence: [evidence, distinctEvidence] }, [first, nextExhibit]);
  const legacyIds = migrationGroups.map((group) => `source-${group.canonical.witnessKey ?? "EX"}:${group.canonical.exhibitInitials ?? group.canonical.witnessInitials ?? "EX"}:${group.canonical.exhibitSequence ?? 1}:hash:${group.sourceHash}`);
  assert.deepEqual(orderExhibitGroups(migrationGroups, legacyIds.reverse()).map((group) => group.canonical.id), migrationGroups.map((group) => group.canonical.id).reverse(), "orders saved with the former source-hash IDs remain readable");

  const conflictingPages = { ...confirmedSame, pageStart: 2, pageEnd: 2 };
  const checks = runPreflight(analysis, [{ ...first, pageStart: 1, pageEnd: 1 }, conflictingPages], BUNDLE_PROFILES[0]);
  assert.ok(checks.some((check) => check.label === "Conflicting repeat selection" && check.severity === "blocking"));
});

test("creates one statement-update suggestion per citation from final bundle pages", async () => {
  const { buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const suggestions = buildStatementUpdateSuggestions([{
    mark: "RC 1",
    description: "24-page report",
    fileName: "report.pdf",
    startPage: 7,
    endPage: 30,
    statementParagraph: 5,
    sourceHash: "source-hash",
    statementReferences: [
      { paragraph: 5, citation: "First reference", exhibitInitials: "RC", exhibitSequence: 1, citationResolution: "resolved", exhibitPageStart: 7, exhibitPageEnd: 30 },
      { paragraph: 19, citation: "Repeated reference", exhibitInitials: "RC", exhibitSequence: 1, citationResolution: "resolved", exhibitPageStart: 7, exhibitPageEnd: 30 },
    ],
  }]);
  assert.deepEqual(suggestions.map((item) => item.line), [
    "Paragraph 5 - [RC1/7-30]",
    "Paragraph 19 - [RC1/7-30]",
  ]);
  assert.equal(suggestions.length, 2, "repeat references get separate statement-update lines without duplicating the exhibit");
  assert.ok(suggestions.every((item) => !item.needsReview));
  assert.doesNotMatch(suggestions.map((item) => item.line).join("\n"), /[âÃÂ]/);
});

test("parses explicit exhibit citation tokens without treating exhibit pages as source PDF pages", async () => {
  const { parseStatementCitationToken } = await import("../app/lib/bundle-engine.ts");
  assert.deepEqual(parseStatementCitationToken("See [RC1/34-38]."), { raw: "[RC1/34-38]", exhibitInitials: "RC", exhibitSequence: 1, requestedExhibitPageStart: 34, requestedExhibitPageEnd: 38, citationResolution: "resolved" });
  assert.deepEqual(parseStatementCitationToken("See [RC 1 / xx]."), { raw: "[RC 1 / xx]", exhibitInitials: "RC", exhibitSequence: 1, requestedExhibitPageStart: undefined, requestedExhibitPageEnd: undefined, citationResolution: "unresolved" });
  assert.equal(parseStatementCitationToken("See [RC-1/34-38].")?.exhibitSequence, 1);
});

test("statement suggestions use canonical formatted labels, fill xx, and flag conflicts", async () => {
  const { buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const base = { mark: "RC 1", description: "x", fileName: "x.pdf", startPage: 99, endPage: 120, statementParagraph: 5, sourceHash: "x" };
  const suggestions = buildStatementUpdateSuggestions([{ ...base, statementReferences: [
    { paragraph: 5, citation: "[RC1/7-30]", exhibitInitials: "RC", exhibitSequence: 1, citationResolution: "resolved", requestedExhibitPageStart: 7, requestedExhibitPageEnd: 30, exhibitPageStart: 7, exhibitPageEnd: 30, exhibitPageLabelStart: "LV-0007", exhibitPageLabelEnd: "LV-0030" },
    { paragraph: 6, citation: "[RC1/xx]", exhibitInitials: "RC", exhibitSequence: 1, citationResolution: "unresolved", exhibitPageStart: 7, exhibitPageEnd: 30 },
    { paragraph: 7, citation: "[RC1/34-38]", exhibitInitials: "RC", exhibitSequence: 1, citationResolution: "resolved", requestedExhibitPageStart: 34, requestedExhibitPageEnd: 38, exhibitPageStart: 7, exhibitPageEnd: 30 },
  ] }]);
  assert.deepEqual(suggestions.map((item) => item.line), ["Paragraph 5 - [RC1/LV-0007-LV-0030]", "Paragraph 6 - [RC1/7-30]", "Paragraph 7 - [RC1/7-30] (requested 34-38; review)"]);
  assert.deepEqual(suggestions.map((item) => item.needsReview), [false, false, true]);
});

test("statement suggestions use a single bundle page for one-page exhibits", async () => {
  const { buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const base = { mark: "AH1", description: "x", fileName: "x.pdf", startPage: 5, endPage: 5, statementParagraph: 18, sourceHash: "x" };
  const suggestions = buildStatementUpdateSuggestions([{ ...base, statementReferences: [
    { paragraph: 18, citation: "[AH1/5]", exhibitInitials: "AH", exhibitSequence: 1, citationResolution: "resolved", requestedExhibitPageStart: 5, requestedExhibitPageEnd: 5, exhibitPageStart: 5, exhibitPageEnd: 5 },
    { paragraph: 21, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, citationResolution: "unresolved", exhibitPageStart: 6, exhibitPageEnd: 6 },
    { paragraph: 23, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, citationResolution: "unresolved", exhibitPageStart: 7, exhibitPageEnd: 7 },
    { paragraph: 30, citation: "Repeated [AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, citationResolution: "unresolved", exhibitPageStart: 7, exhibitPageEnd: 7 },
    { paragraph: 35, citation: "[AH1/8-9]", exhibitInitials: "AH", exhibitSequence: 1, citationResolution: "resolved", requestedExhibitPageStart: 8, requestedExhibitPageEnd: 9, exhibitPageStart: 8, exhibitPageEnd: 9 },
  ] }]);
  assert.deepEqual(suggestions.map((item) => item.line), [
    "Paragraph 18 - [AH1/5]",
    "Paragraph 21 - [AH1/6]",
    "Paragraph 23 - [AH1/7]",
    "Paragraph 30 - [AH1/7]",
    "Paragraph 35 - [AH1/8-9]",
  ]);
  assert.ok(suggestions.every((item) => !item.needsReview));
});

test("appends uncited exhibits after cited suggestions using index descriptions", async () => {
  const { buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const cited = {
    mark: "AH 1",
    exhibitNumber: 2,
    description: "Share purchase agreement",
    fileName: "spa.pdf",
    startPage: 11,
    endPage: 18,
    exhibitPageStart: 11,
    exhibitPageEnd: 18,
    exhibitPageLabelStart: "11",
    exhibitPageLabelEnd: "18",
    statementParagraph: 4,
    sourceHash: "cited",
    volumeNumber: 1,
    statementReferences: [{
      paragraph: 4,
      citation: "[AH1/xx]",
      exhibitInitials: "AH",
      exhibitSequence: 1,
      citationResolution: "unresolved",
      exhibitPageStart: 11,
      exhibitPageEnd: 18,
      exhibitPageLabelStart: "11",
      exhibitPageLabelEnd: "18",
    }],
  };
  const laterManual = {
    mark: "EX 1",
    exhibitNumber: 4,
    description: "Supplemental spreadsheet",
    fileName: "costs.xlsx",
    startPage: 67,
    endPage: 67,
    exhibitPageStart: 67,
    exhibitPageEnd: 67,
    exhibitPageLabelStart: "67",
    exhibitPageLabelEnd: "67",
    statementParagraph: null,
    statementReferences: [],
    sourceHash: "sheet",
    volumeNumber: 1,
    manualAddition: true,
    citationStatus: "not-cited-manual-addition",
  };
  const earlierManual = {
    ...laterManual,
    exhibitNumber: 1,
    description: "Payment ledger [working copy]",
    fileName: "ledger.pdf",
    startPage: 61,
    endPage: 66,
    exhibitPageStart: 61,
    exhibitPageEnd: 66,
    exhibitPageLabelStart: "61",
    exhibitPageLabelEnd: "66",
    sourceHash: "ledger",
  };
  const suggestions = buildStatementUpdateSuggestions([cited, laterManual, earlierManual]);
  assert.deepEqual(suggestions.map((item) => item.line), [
    "Paragraph 4 - [AH1/11-18]",
    "Uncited exhibits — no statement reference",
    "1. Payment ledger [working copy] — pages 61-66",
    "4. Supplemental spreadsheet — page 67",
  ]);
  const uncited = suggestions.slice(2);
  assert.ok(uncited.every((item) => !item.line.startsWith("Paragraph ")));
  assert.ok(uncited.every((item) => !/\[AH\d|\[LV\d/i.test(item.line)));
  assert.match(uncited[0].line, /\[working copy\]/);
  assert.ok(suggestions.every((item) => item.line !== ""));
  assert.deepEqual(buildStatementUpdateSuggestions([cited]).map((item) => item.line), ["Paragraph 4 - [AH1/11-18]"]);
});

test("names every volume on uncited rows in a split bundle", async () => {
  const { buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const uncited = (exhibitNumber, description, volumeNumber, start, end, labels = {}) => ({
    mark: "EX 1",
    exhibitNumber,
    description,
    fileName: `${description}.pdf`,
    startPage: start,
    endPage: end,
    exhibitPageStart: start,
    exhibitPageEnd: end,
    exhibitPageLabelStart: labels.start ?? String(start),
    exhibitPageLabelEnd: labels.end ?? String(end),
    statementParagraph: null,
    statementReferences: [],
    sourceHash: `hash-${exhibitNumber}`,
    volumeNumber,
    manualAddition: true,
    citationStatus: "not-cited-manual-addition",
  });
  const split = buildStatementUpdateSuggestions([
    uncited(12, "Supplemental spreadsheet", 2, 143, 143),
    uncited(8, "Payment ledger", 1, 43, 49),
  ]);
  assert.deepEqual(split.map((item) => item.line), [
    "Uncited exhibits — no statement reference",
    "8. Payment ledger — Vol. 1, pages 43-49",
    "12. Supplemental spreadsheet — Vol. 2, page 143",
  ]);
  const labelled = buildStatementUpdateSuggestions([
    uncited(8, "Payment ledger", 1, 43, 49, { start: "Vol. 1/43", end: "Vol. 1/49" }),
    { ...uncited(9, "Cited report", 2, 50, 50), manualAddition: false, citationStatus: undefined, statementParagraph: 3, statementReferences: [{ paragraph: 3, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, citationResolution: "unresolved", exhibitPageStart: 50, exhibitPageEnd: 50, exhibitPageLabelStart: "50", exhibitPageLabelEnd: "50", volumeNumber: 2 }] },
  ]);
  assert.deepEqual(labelled.map((item) => item.line), [
    "Paragraph 3 - [AH1/Vol. 2/50]",
    "Uncited exhibits — no statement reference",
    "8. Payment ledger — pages Vol. 1/43-Vol. 1/49",
  ]);
  assert.doesNotMatch(labelled[2].line, /Vol\. 1, .*Vol\. 1/);
});

test("uncited suggestion lines omit a trailing dash when finished pages are not yet known", async () => {
  const { buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const suggestions = buildStatementUpdateSuggestions([{
    mark: "EX 1",
    exhibitNumber: 3,
    description: "Payment ledger",
    fileName: "ledger.pdf",
    startPage: 1,
    endPage: 1,
    statementParagraph: 0,
    sourceHash: "ledger",
    manualAddition: true,
    citationStatus: "not-cited-manual-addition",
    statementReferences: [],
  }]);
  assert.deepEqual(suggestions.map((item) => item.line), [
    "Uncited exhibits — no statement reference",
    "3. Payment ledger",
  ]);
  assert.doesNotMatch(suggestions[1].line, / — $/);
});

test("lists uncited exhibits from a finished mixed bundle without inventing citations", async () => {
  const { buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  async function pdfFile(name, pages, text) {
    const pdf = await PDFDocument.create();
    for (let index = 0; index < pages; index += 1) pdf.addPage([595.28, 841.89]).drawText(text ?? name);
    return new File([await pdf.save()], name, { type: "application/pdf" });
  }
  const citedFile = await pdfFile("agreement.pdf", 1, "Agreement");
  const ledgerFile = await pdfFile("ledger.pdf", 2, "Ledger");
  const sheetFile = await pdfFile("sheet.pdf", 1, "Sheet");
  const evidence = [
    { id: "ledger", file: ledgerFile, name: ledgerFile.name, extension: "pdf", text: "Ledger", marker: null, sha256: "l".repeat(64), pageCount: 2, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" },
    { id: "cited", file: citedFile, name: citedFile.name, extension: "pdf", text: "Agreement", marker: null, sha256: "c".repeat(64), pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" },
    { id: "sheet", file: sheetFile, name: sheetFile.name, extension: "pdf", text: "Sheet", marker: null, sha256: "s".repeat(64), pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" },
  ];
  const timestamp = "2026-08-08T10:00:00.000Z";
  const cited = { id: "cited", mark: "AH 1", provisionalNumber: 2, description: "Share purchase agreement", date: "Date not stated", paragraph: 4, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: "cited", confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "AH", statementId: "s", statementName: "Statement.docx", sequenceOrder: 2 };
  const ledger = { id: "ledger", mark: "EX 1", provisionalNumber: 1, description: "Payment ledger [working copy]", date: "Date not stated", paragraph: 0, citation: "", citationResolution: "none", discoverySignals: ["Manually added by reviewer"], evidenceId: "ledger", confidence: 100, rationale: "Reviewer added", included: true, confirmed: true, exhibitInitials: "EX", exhibitSequence: 1, witnessInitials: "EX", witnessKey: "general-exhibits::EX", sequenceOrder: 1, manualAddition: true, manualAddedAt: timestamp, manualWarningAcknowledgedAt: timestamp };
  const sheet = { ...ledger, id: "sheet", provisionalNumber: 3, description: "Supplemental spreadsheet", evidenceId: "sheet", sequenceOrder: 3 };
  const analysis = { statementName: "Statement.docx", statementHash: "s".repeat(64), caseTitle: "Test", candidates: [ledger, cited, sheet], evidence, unreferenced: [], statementWarnings: [], generatedAt: timestamp };
  const layout = { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false };
  const unsplit = await buildBundle(analysis, [ledger, cited, sheet], { layout });
  const unsplitSuggestions = buildStatementUpdateSuggestions(unsplit.records);
  assert.match(unsplitSuggestions[0].line, /^Paragraph 4 - \[AH1\//);
  assert.equal(unsplitSuggestions[1].line, "Uncited exhibits — no statement reference");
  const uncitedUnsplit = unsplitSuggestions.slice(2);
  assert.deepEqual(uncitedUnsplit.map((item) => item.line.replace(/ — .+$/, "")), ["1. Payment ledger [working copy]", "3. Supplemental spreadsheet"]);
  assert.ok(uncitedUnsplit.every((item) => !item.line.startsWith("Paragraph ")));
  assert.ok(uncitedUnsplit.every((item) => !/\[AH\d|\[LV\d/i.test(item.line)));
  assert.ok(uncitedUnsplit.every((item) => !/Vol\.\s*\d+/.test(item.line)));
  const ledgerRecord = unsplit.records.find((record) => record.exhibitNumber === 1);
  const sheetRecord = unsplit.records.find((record) => record.exhibitNumber === 3);
  assert.equal(ledgerRecord.description, "Payment ledger [working copy]");
  assert.ok(uncitedUnsplit[0].line.includes(ledgerRecord.exhibitPageLabelStart));
  assert.ok(uncitedUnsplit[0].line.includes(ledgerRecord.exhibitPageLabelEnd));
  assert.ok(uncitedUnsplit[1].line.includes(sheetRecord.exhibitPageLabelStart));

  const split = await buildBundle(analysis, [ledger, cited, sheet], { layout: { ...layout, volumePageLimit: 1 } });
  assert.ok((split.volumes?.length ?? 0) >= 2);
  const splitRecords = split.volumes.flatMap((volume) => volume.records);
  const splitSuggestions = buildStatementUpdateSuggestions(splitRecords);
  assert.match(splitSuggestions[0].line, /^Paragraph 4 - \[AH1\/Vol\./);
  assert.equal(splitSuggestions[1].line, "Uncited exhibits — no statement reference");
  const volLedger = splitRecords.find((record) => record.exhibitNumber === 1);
  const volSheet = splitRecords.find((record) => record.exhibitNumber === 3);
  assert.equal(volLedger.volumeNumber, 1);
  assert.ok((volSheet.volumeNumber ?? 1) > 1);
  const uncitedSplit = splitSuggestions.slice(2).map((item) => item.line);
  assert.match(uncitedSplit[0], new RegExp(`^1\\. Payment ledger \\[working copy\\] — Vol\\. ${volLedger.volumeNumber}, pages `));
  assert.match(uncitedSplit[1], new RegExp(`^3\\. Supplemental spreadsheet — Vol\\. ${volSheet.volumeNumber}, page `));
  assert.doesNotMatch(uncitedSplit.join("\n"), /Vol\. \d+, .*Vol\. \d+/);
});

test("keeps distinct RC1 documents as separate exhibits with cumulative bundle pages", async () => {
  const { buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  async function pdfFile(name, pages) {
    const pdf = await PDFDocument.create();
    for (let index = 0; index < pages; index += 1) pdf.addPage([595.28, 841.89]).drawText(`Page ${index + 1}`);
    return new File([await pdf.save()], name, { type: "application/pdf" });
  }
  const firstFile = await pdfFile("prior.pdf", 6);
  const latestFile = await pdfFile("contract.pdf", 24);
  const evidence = [
    { id: "prior", file: firstFile, name: firstFile.name, extension: "pdf", text: "prior", marker: null, sha256: "hash-prior", pageCount: 6, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" },
    { id: "latest", file: latestFile, name: latestFile.name, extension: "pdf", text: "latest", marker: null, sha256: "hash-latest", pageCount: 24, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" },
  ];
  const first = { id: "p4", mark: "RC 1", provisionalNumber: 1, description: "Prior RC1 document", date: "Date not stated", paragraph: 4, citation: "[RC1/1-6]", citationToken: "[RC1/1-6]", exhibitInitials: "RC", exhibitSequence: 1, requestedExhibitPageStart: 1, requestedExhibitPageEnd: 6, citationResolution: "resolved", discoverySignals: [], evidenceId: "prior", confidence: 100, rationale: "reviewer", included: true, confirmed: true, sequenceOrder: 1, witnessKey: "rc" };
  const latest = { ...first, id: "p30", provisionalNumber: 2, sequenceOrder: 2, paragraph: 30, citation: "[RC1/7-30]", citationToken: "[RC1/7-30]", requestedExhibitPageStart: 7, requestedExhibitPageEnd: 30, evidenceId: "latest", description: "Latest RC1 contract" };
  const repeatedPrior = { ...first, id: "p44", paragraph: 44, citation: "[RC1/1-6]", repeatDecision: "same" };
  const analysis = { statementName: "RC statement.docx", statementHash: "statement", caseTitle: "Test case", candidates: [first, latest, repeatedPrior], evidence, unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const build = await buildBundle(analysis, [first, latest, repeatedPrior], { pagination: { matchPdfPageOrder: false } });
  assert.equal(build.records.length, 2, "each distinct document remains an individual exhibit");
  assert.deepEqual(build.records.map((record) => record.exhibitNumber), [1, 2]);
  assert.deepEqual(build.records.flatMap((record) => record.statementReferences).map((reference) => [reference.exhibitPageStart, reference.exhibitPageEnd]), [[1, 6], [1, 6], [7, 30]]);
  assert.deepEqual(buildStatementUpdateSuggestions(build.records).map((item) => item.line), ["Paragraph 4 - [RC1/1-6]", "Paragraph 30 - [RC1/7-30]", "Paragraph 44 - [RC1/1-6]"]);
  const rendered = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(build.bytes) }).promise;
  const indexText = (await (await rendered.getPage(2)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  assert.match(indexText, /\b1\b/);
  assert.doesNotMatch(indexText, /RC 1/);
  await rendered.destroy();
});

test("blocks conflicting statement-reference marks instead of creating separate cursors", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const pdf = await PDFDocument.create();
  pdf.addPage([595.28, 841.89]).drawText("Shared page 1");
  pdf.addPage([595.28, 841.89]).drawText("Shared page 2");
  const file = new File([await pdf.save()], "shared.pdf", { type: "application/pdf" });
  const evidence = [{ id: "shared", file, name: file.name, extension: "pdf", text: "shared", marker: null, sha256: "shared-hash", pageCount: 2, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" }];
  const ah = { id: "ah", mark: "AH 1", provisionalNumber: 1, description: "AH document", date: "Date not stated", paragraph: 5, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, citationResolution: "unresolved", discoverySignals: [], evidenceId: "shared", confidence: 100, rationale: "reviewer", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "amelia", sequenceOrder: 1 };
  const rc = { ...ah, id: "rc", mark: "RC 1", description: "RC document", paragraph: 8, citation: "[RC1/xx]", exhibitInitials: "RC", witnessInitials: "RC", witnessKey: "robin", sequenceOrder: 2 };
  const analysis = { statementName: "two statements", statementHash: "statement", caseTitle: "Test case", candidates: [ah, rc], evidence, unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  await assert.rejects(buildBundle(analysis, [ah, rc]), /reference marks conflict \(AH1, RC1\)/i);
});

test("committed final order drives PDF-order suggestions, index records and manifest ranges", async () => {
  const { buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const { deriveExhibitGroups } = await import("../app/lib/exhibit-groups.ts");
  async function source(id) {
    const pdf = await PDFDocument.create();
    pdf.addPage([595.28, 841.89]).drawText(`Evidence ${id}`);
    const file = new File([await pdf.save()], `${id}.pdf`, { type: "application/pdf" });
    return { id, file, name: file.name, extension: "pdf", text: `Evidence ${id}`, marker: null, sha256: id.repeat(64).slice(0, 64), pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" };
  }
  const evidence = await Promise.all([source("a"), source("b"), source("c")]);
  const candidate = (id, paragraph, order) => ({ id, mark: "LV 1", provisionalNumber: order, description: `Exhibit ${id}`, date: "Date not stated", paragraph, citation: `[LV1/xx]`, citationToken: "[LV1/xx]", exhibitInitials: "LV", exhibitSequence: 1, citationResolution: "unresolved", discoverySignals: [], evidenceId: id, confidence: 100, rationale: "reviewer", included: true, confirmed: true, witnessInitials: "LV", witnessKey: "lucia::LV", sequenceOrder: order, statementId: "lucia", statementName: "Statement.docx" });
  const candidates = [candidate("a", 6, 1), candidate("b", 20, 2), candidate("c", 30, 3)];
  const analysis = { statementName: "Statement.docx", statementHash: "statement", caseTitle: "Test case", witnessInitials: "LV", candidates, evidence, unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const committedOrder = deriveExhibitGroups(analysis, candidates).map((group) => group.id).reverse();
  const build = await buildBundle(analysis, candidates, { canonicalOrder: committedOrder, pagination: { matchPdfPageOrder: true } });

  assert.deepEqual(build.records.map((record) => record.fileName), ["c.pdf", "b.pdf", "a.pdf"]);
  assert.deepEqual(build.records.map((record) => [record.startPage, record.endPage, record.exhibitPageStart, record.exhibitPageEnd]), [[3, 3, 3, 3], [4, 4, 4, 4], [5, 5, 5, 5]]);
  assert.deepEqual(buildStatementUpdateSuggestions(build.records).map((item) => item.line), ["Paragraph 6 - [LV1/5]", "Paragraph 20 - [LV1/4]", "Paragraph 30 - [LV1/3]"]);
  assert.deepEqual(build.manifest.exhibits.map((record) => [record.fileName, record.startPage, record.endPage, record.exhibitPageStart, record.exhibitPageEnd]), [["c.pdf", 3, 3, 3, 3], ["b.pdf", 4, 4, 4, 4], ["a.pdf", 5, 5, 5, 5]]);
});

test("index links target first content after templates and measured rows create extra index pages", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  async function template(slot) {
    const pdf = await PDFDocument.create();
    pdf.addPage([595.28, 841.89]).drawText(`${slot} template`);
    return reviewedTemplate(slot, new File([await pdf.save()], `${slot}.pdf`, { type: "application/pdf" }));
  }
  const statement = await fixtureFile(new URL("01_Witness_Statement_Amelia_Hart.docx", fixtureRoot));
  const source = await fixtureFile(new URL("Executed_Supply_Agreement_2026-02-01.pdf", evidenceRoot));
  const analysis = await analyseFiles(statement, [source]);
  const evidence = analysis.evidence[0];
  const one = { ...analysis.candidates[0], evidenceId: evidence.id, included: true, confirmed: true, mark: "AH 1", provisionalNumber: 1, statementName: statement.name };
  const templates = await Promise.all(["cover", "index", "divider", "exhibitCover"].map(template));
  const templated = await buildBundle(analysis, [one], { templates, layout: { includeDividerPages: true, includeExhibitCoverPages: true, countOptionalPagesInReferences: false, volumePageLimit: 0 } });
  assert.equal(templated.records[0].startPage, 5, "index link/bookmark destination starts after cover, index, divider and exhibit cover templates");
  const renderedTemplate = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(templated.bytes) }).promise;
  const indexItems = (await (await renderedTemplate.getPage(2)).getTextContent()).items.filter((item) => "str" in item && item.str.trim());
  const indexNumber = indexItems.find((item) => item.str === "1" && Math.abs(item.transform[4] - 64) < 1);
  const description = indexItems.find((item) => item.str.includes(templated.records[0].description.slice(0, 12)));
  const pageRange = indexItems.find((item) => /^\d/.test(item.str) && item.transform[4] >= 460);
  assert.ok(indexNumber, "fixed-layout template writes the item number in its documented first column");
  assert.ok(description && Math.abs(description.transform[4] - 103) < 1, "fixed-layout template writes descriptions inside the documented middle column");
  assert.ok(pageRange && pageRange.transform[4] + pageRange.width <= 543.5, "fixed-layout template right-aligns page ranges inside the documented final column");
  await renderedTemplate.destroy();

  const many = Array.from({ length: 15 }, (_, index) => ({ ...one, id: `long-${index}`, mark: `AH ${index + 1}`, provisionalNumber: index + 1, description: `Long exhibit description ${index + 1}: ${"repeated reference detail ".repeat(12)}`, repeatDecision: index ? "separate" : undefined }));
  const longIndex = await buildBundle(analysis, many);
  assert.equal(longIndex.records.length, 15);
  assert.ok(longIndex.records[0].startPage > 2, "measured long rows require more than one generated index page before content");
});

test("optional pages excluded from references are left unnumbered so all visible references agree", async () => {
  const { buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const sourcePdf = await PDFDocument.create();
  sourcePdf.addPage([595.28, 841.89]).drawText("Contract");
  const file = new File([await sourcePdf.save()], "contract.pdf", { type: "application/pdf" });
  const evidence = [{ id: "e", file, name: file.name, extension: "pdf", text: "Contract", marker: null, sha256: "e".repeat(64), pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" }];
  const candidate = { id: "c", mark: "AH 1", provisionalNumber: 1, description: "Contract", date: "2026-01-01", paragraph: 2, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: "e", confidence: 100, rationale: "test", included: true, confirmed: true, statementName: "Statement.docx", witnessInitials: "AH", witnessKey: "AH", sequenceOrder: 1 };
  const analysis = { statementName: "Statement.docx", statementHash: "s".repeat(64), caseTitle: "Test", candidates: [candidate], evidence, unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const none = await buildBundle(analysis, [candidate], { pagination: { matchPdfPageOrder: false }, layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 } });
  const physicalOnly = await buildBundle(analysis, [candidate], { pagination: { matchPdfPageOrder: false }, layout: { includeDividerPages: true, includeExhibitCoverPages: true, countOptionalPagesInReferences: false, volumePageLimit: 0 } });
  const legallyCounted = await buildBundle(analysis, [candidate], { pagination: { matchPdfPageOrder: false }, layout: { includeDividerPages: true, includeExhibitCoverPages: true, countOptionalPagesInReferences: true, volumePageLimit: 0 } });
  assert.deepEqual([none.pageCount, physicalOnly.pageCount, legallyCounted.pageCount], [3, 5, 5], "global cover/index and two built-in optional pages are physically explicit");
  assert.deepEqual([none.records[0].startPage, physicalOnly.records[0].startPage, legallyCounted.records[0].startPage], [3, 5, 5]);
  assert.deepEqual([none.records[0].exhibitPageStart, physicalOnly.records[0].exhibitPageStart, legallyCounted.records[0].exhibitPageStart], [1, 1, 3]);
  assert.deepEqual([none, physicalOnly, legallyCounted].map((build) => buildStatementUpdateSuggestions(build.records)[0].line), ["Paragraph 2 - [AH1/1]", "Paragraph 2 - [AH1/1]", "Paragraph 2 - [AH1/3]"]);
  assert.deepEqual(await readPageLabels(physicalOnly.bytes), ["1", "2", "", "", "1"], "unnumbered optional pages have explicit empty PDF labels before exhibit numbering starts");

  const wordSource = new File(["unused"], "divider.docx");
  const unusedWord = { slot: "divider", file: wordSource, sha256: createHash("sha256").update("unused").digest("hex"), sourceFormat: "docx", templateConfirmed: false };
  await assert.doesNotReject(buildBundle(analysis, [candidate], { templates: [unusedWord], pagination: { matchPdfPageOrder: false }, layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 } }));
  await assert.rejects(buildBundle(analysis, [candidate], { templates: [unusedWord], pagination: { matchPdfPageOrder: false }, layout: { includeDividerPages: true, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 } }), /Preview .* retain its exact converted PDF/i);
  const approvedWord = await reviewedTemplate("divider", wordSource, { sourceFormat: "docx", pdfFile: new File([await sourcePdf.save()], "divider.pdf", { type: "application/pdf" }) });
  const wordIncluded = await buildBundle(analysis, [candidate], { templates: [approvedWord], pagination: { matchPdfPageOrder: false }, layout: { includeDividerPages: true, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 } });
  assert.equal(wordIncluded.pageCount, 4);
  assert.equal(wordIncluded.records[0].exhibitPageStart, 1);
  assert.equal(wordIncluded.manifest.templates[0].appearanceConfirmed, true);
  assert.equal(wordIncluded.manifest.templates[0].matterConfirmed, true);

  const legacyPdfFile = new File([await sourcePdf.save()], "legacy-index.pdf", { type: "application/pdf" });
  const legacyPdfHash = createHash("sha256").update(new Uint8Array(await legacyPdfFile.arrayBuffer())).digest("hex");
  await assert.rejects(
    buildBundle(analysis, [candidate], { templates: [{ slot: "index", file: legacyPdfFile, sha256: legacyPdfHash, sourceFormat: "pdf", templateConfirmed: true }] }),
    /Review the matter details shown in the exact PDF preview/i,
    "the deprecated general approval cannot stand in for matter-details confirmation",
  );
  const placeholder = { kind: "placeholder", value: "[CASE NUMBER]", normalizedValue: "[CASE NUMBER]", pageNumbers: [1], evidence: ["[CASE NUMBER]"], unverified: true };
  const reviewedWithPlaceholder = await reviewedTemplate("index", legacyPdfFile, { placeholders: [placeholder] });
  const withoutPlaceholderConfirmation = { ...reviewedWithPlaceholder, reviewState: { ...reviewedWithPlaceholder.reviewState, placeholderConfirmation: undefined } };
  await assert.rejects(
    buildBundle(analysis, [candidate], { templates: [withoutPlaceholderConfirmation] }),
    /Confirm the possible unfinished placeholders/i,
  );
});

test("section numbering never invents fallback or negative labels for optional pages", async () => {
  const { buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const sourcePdf = await PDFDocument.create();
  sourcePdf.addPage([595.28, 841.89]).drawText("Section-numbered exhibit");
  const source = new File([await sourcePdf.save()], "section.pdf", { type: "application/pdf" });
  const evidence = { id: "section-source", file: source, name: source.name, extension: "pdf", text: "Section-numbered exhibit", marker: null, sha256: "f".repeat(64), pageCount: 1, readableText: true, encrypted: false, rotationPages: [], annotationPages: [], ocrPages: [], ocrStatus: "not-needed" };
  const candidate = { id: "section-candidate", mark: "AH 1", provisionalNumber: 1, description: "Section-numbered exhibit", date: "2026-08-11", paragraph: 2, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: evidence.id, confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "amelia::AH", statementName: "Statement.docx", sequenceOrder: 1 };
  const analysis = { statementName: "Statement.docx", statementHash: "s", caseTitle: "Test", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const options = { pagination: { matchPdfPageOrder: false, scheme: "section" }, layout: { includeDividerPages: true, includeExhibitCoverPages: true, countOptionalPagesInReferences: false, volumePageLimit: 0 } };
  const uncounted = await buildBundle(analysis, [candidate], options);
  const counted = await buildBundle(analysis, [candidate], { ...options, layout: { ...options.layout, countOptionalPagesInReferences: true } });
  assert.match(uncounted.checks.find((check) => check.label === "PDF page labels")?.detail ?? "", /Exhibit mark and page labels verified/i);
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  async function pageTexts(bytes) {
    const pdf = await pdfjs.getDocument({ data: new Uint8Array(bytes) }).promise;
    const values = [];
    for (const number of [3, 4, 5]) values.push((await (await pdf.getPage(number)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" "));
    await pdf.destroy();
    return values;
  }
  const uncountedTexts = await pageTexts(uncounted.bytes);
  assert.doesNotMatch(uncountedTexts[0], /AH1-1/);
  assert.doesNotMatch(uncountedTexts[1], /AH1-1/);
  assert.match(uncountedTexts[2], /AH1-1/);
  assert.doesNotMatch(uncountedTexts.join(" "), /\bB-|AH1--|AH1-0\b/);
  assert.deepEqual(await readPageLabels(uncounted.bytes), ["1", "2", "", "", "AH1-1"]);
  assert.equal(buildStatementUpdateSuggestions(uncounted.records)[0].line, "Paragraph 2 - [AH1/AH1-1]");
  const uncountedPdf = await pdfjs.getDocument({ data: new Uint8Array(uncounted.bytes) }).promise;
  const uncountedIndex = (await (await uncountedPdf.getPage(2)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  assert.match(uncountedIndex, /AH1-1/);
  await uncountedPdf.destroy();
  const countedTexts = await pageTexts(counted.bytes);
  assert.match(countedTexts[0], /AH1-1/);
  assert.match(countedTexts[1], /AH1-2/);
  assert.match(countedTexts[2], /AH1-3/);
  assert.doesNotMatch(countedTexts.join(" "), /\bB-|AH1--|AH1-0\b/);
  assert.deepEqual(await readPageLabels(counted.bytes), ["1", "2", "AH1-1", "AH1-2", "AH1-3"]);
  assert.equal(buildStatementUpdateSuggestions(counted.records)[0].line, "Paragraph 2 - [AH1/AH1-3]");
  const countedPdf = await pdfjs.getDocument({ data: new Uint8Array(counted.bytes) }).promise;
  const countedIndex = (await (await countedPdf.getPage(2)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  assert.match(countedIndex, /AH1-3/);
  await countedPdf.destroy();
});

test("page-number digit width and template-label setting stay consistent", async () => {
  const { buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const sourcePdf = await PDFDocument.create();
  sourcePdf.addPage([595.28, 841.89]).drawText("Source page");
  const sourceFile = new File([await sourcePdf.save()], "source.pdf", { type: "application/pdf" });
  const evidence = [{
    id: "source",
    file: sourceFile,
    name: sourceFile.name,
    extension: "pdf",
    text: "Source page",
    marker: null,
    sha256: "source-hash",
    pageCount: 1,
    readableText: true,
    encrypted: false,
    rotationPages: [],
    ocrPages: [],
    ocrStatus: "not-needed",
  }];
  const candidate = {
    id: "p1",
    mark: "AH 1",
    provisionalNumber: 1,
    description: "Source document",
    date: "Date not stated",
    paragraph: 1,
    citation: "[AH1/xx]",
    citationToken: "[AH1/xx]",
    exhibitInitials: "AH",
    exhibitSequence: 1,
    citationResolution: "unresolved",
    discoverySignals: [],
    evidenceId: "source",
    confidence: 100,
    rationale: "reviewer",
    included: true,
    confirmed: true,
    witnessInitials: "AH",
    witnessKey: "amelia",
    sequenceOrder: 1,
  };
  const analysis = {
    statementName: "statement.docx",
    statementHash: "statement",
    caseTitle: "Test case",
    candidates: [candidate],
    evidence,
    unreferenced: [],
    statementWarnings: [],
    generatedAt: new Date().toISOString(),
  };
  async function template(slot) {
    const pdf = await PDFDocument.create();
    pdf.addPage([595.28, 841.89]).drawText(slot + " template");
    const file = new File([await pdf.save()], slot + ".pdf", { type: "application/pdf" });
    return reviewedTemplate(slot, file);
  }
  const templates = await Promise.all(["cover", "index"].map(template));
  const noTemplateLabels = await buildBundle(analysis, [candidate], {
    templates,
    pagination: { matchPdfPageOrder: false, prefix: "LV-", padding: 4, countTemplates: false },
  });
  assert.equal(noTemplateLabels.records[0].startPage, 3);
  assert.deepEqual(await readPageLabels(noTemplateLabels.bytes), ["", "", "LV-0001"]);
  const noTemplatePdf = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(noTemplateLabels.bytes) }).promise;
  const noTemplateContent = (await (await noTemplatePdf.getPage(3)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  assert.match(noTemplateContent, /LV-0001/);
  assert.doesNotMatch(noTemplateContent, /LV-0000/);
  await noTemplatePdf.destroy();
  assert.equal(buildStatementUpdateSuggestions(noTemplateLabels.records)[0].line, "Paragraph 1 - [AH1/LV-0001]");

  const numberedTemplates = await buildBundle(analysis, [candidate], {
    templates,
    pagination: { matchPdfPageOrder: false, prefix: "LV-", padding: 4, countTemplates: true, preliminary: "roman", includePrefixInIndex: true },
  });
  assert.deepEqual(await readPageLabels(numberedTemplates.bytes), ["i", "ii", "LV-0001"]);
  const numberedPdf = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(numberedTemplates.bytes) }).promise;
  const indexText = (await (await numberedPdf.getPage(2)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  assert.match(indexText, /LV-0001/);
  await numberedPdf.destroy();

  const continuousPrefix = await buildBundle(analysis, [candidate], {
    templates,
    pagination: { matchPdfPageOrder: true, prefix: "AH-", suffix: "", padding: 0, includePrefixInIndex: true },
  });
  assert.deepEqual(await readPageLabels(continuousPrefix.bytes), ["AH-1", "AH-2", "AH-3"]);
  const continuousPdf = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(continuousPrefix.bytes) }).promise;
  const coverText = (await (await continuousPdf.getPage(1)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  const continuousIndexText = (await (await continuousPdf.getPage(2)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  const exhibitText = (await (await continuousPdf.getPage(3)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  assert.match(coverText, /AH-1/);
  assert.doesNotMatch(coverText, /AH 1/, "a single-volume custom cover receives no internal bundle label");
  assert.doesNotMatch(coverText, /Bundle page|Prepared locally by Exhibit Builder|Administrative divider page/);
  assert.match(continuousIndexText, /AH-2/);
  assert.match(continuousIndexText, /AH-3/, "index page range uses the same continuous prefix");
  assert.match(exhibitText, /AH-3/);
  assert.equal(buildStatementUpdateSuggestions(continuousPrefix.records)[0].line, "Paragraph 1 - [AH1/AH-3]");
  await continuousPdf.destroy();

  const defaultPrefix = await buildBundle(analysis, [candidate], {
    templates,
    pagination: { matchPdfPageOrder: true, prefix: "AH-", suffix: "", padding: 0 },
  });
  const defaultPrefixPdf = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(defaultPrefix.bytes) }).promise;
  const defaultIndexText = (await (await defaultPrefixPdf.getPage(2)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  assert.match(defaultIndexText, /AH-2/, "an unspecified includePrefixInIndex flag still prints the prefix in the index");
  await defaultPrefixPdf.destroy();

  await assert.rejects(
    buildBundle(analysis, [candidate], {
      templates,
      pagination: { matchPdfPageOrder: false, prefix: "X".repeat(200), padding: 4, includePrefixInIndex: true },
    }),
    /Index layout blocked \(UNRENDERABLE_PAGE_LABEL\).*cannot fit/i,
    "an unsafe page reference blocks the build instead of shrinking or overflowing outside the index column",
  );
  await assert.rejects(
    buildBundle(analysis, [candidate], {
      templates,
      pagination: { matchPdfPageOrder: false, prefix: "X".repeat(200), padding: 4, includePrefixInIndex: false },
    }),
    /Index layout blocked \(UNRENDERABLE_PAGE_LABEL\).*cannot fit/i,
    "there is no separate index numbering path that can hide an unprintable stamp prefix",
  );
});

test("witness detail changes preserve analysed evidence, matches and approvals", async () => {
  const { applyWitnessDetails } = await import("../app/lib/bundle-engine.ts");
  const { deriveExhibitGroups } = await import("../app/lib/exhibit-groups.ts");
  const source = new File(["unchanged source bytes"], "Evidence.txt");
  const evidence = { id: "evidence-1", file: source, name: source.name, extension: "txt", text: "unchanged source bytes", marker: null, sha256: "a".repeat(64), pageCount: 0, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" };
  const candidate = { id: "statement-1:candidate-1", mark: "AH 1", provisionalNumber: 1, description: "Approved document", aliases: ["approved alias"], reviewNote: "keep this", date: "2026-08-08", paragraph: 2, citation: "[Exhibit]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: ["placeholder"], evidenceId: evidence.id, confidence: 99, rationale: "reviewed", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "amelia hart::AH", statementId: "statement-1", statementName: "Statement.docx", sequenceOrder: 1 };
  const analysis = { statementName: "Statement.docx", statementHash: "s".repeat(64), caseTitle: "Test", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: "2026-08-08T00:00:00.000Z" };
  const statement = { id: "statement-1", file: new File(["statement"], "Statement.docx"), witnessName: "Lucia Varela", witnessInitials: "LV" };
  const updated = applyWitnessDetails(analysis, [candidate], statement);
  assert.strictEqual(updated.analysis.evidence, analysis.evidence, "evidence analysis collection is not recreated");
  assert.strictEqual(updated.analysis.evidence[0], evidence, "source extraction record is preserved by identity");
  assert.equal(updated.candidates[0].evidenceId, evidence.id);
  assert.equal(updated.candidates[0].confirmed, true);
  assert.deepEqual(updated.candidates[0].aliases, ["approved alias"]);
  assert.equal(updated.candidates[0].reviewNote, "keep this");
  assert.equal(updated.candidates[0].witnessKey, "lucia varela::LV");
  assert.equal(updated.candidates[0].exhibitInitials, "LV");
  assert.equal(deriveExhibitGroups(updated.analysis, updated.candidates)[0].outputMark, "LV1");
});

test("rejects pathological PDF dimensions and compressed DOCX expansion before rendering", async () => {
  const { analyseEvidenceFiles } = await import("../app/lib/bundle-engine.ts");
  const extremePdf = await PDFDocument.create();
  extremePdf.addPage([30_001, 100]).drawText("Extreme dimensions");
  await assert.rejects(
    analyseEvidenceFiles([new File([await extremePdf.save()], "Extreme.pdf", { type: "application/pdf" })]),
    /unsafe page dimensions/i,
  );

  const zip = new JSZip();
  zip.file("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
  zip.file("word/document.xml", `<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>${"A".repeat(2 * 1024 * 1024)}</w:t></w:r></w:p></w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 9 } });
  assert.ok(bytes.byteLength < 100_000, "DOCX safety fixture remains small on disk");
  await assert.rejects(
    analyseEvidenceFiles([new File([bytes], "Compressed.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" })]),
    /DOCX archive safety limits/i,
  );
});

test("non-A4 PDF pages are proportionally converted by default or retained by explicit choice", async () => {
  const { analyseEvidenceFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const sourcePdf = await PDFDocument.create();
  sourcePdf.addPage([612, 792]).drawText("Letter source page");
  const source = new File([await sourcePdf.save()], "Letter_Source.pdf", { type: "application/pdf" });
  const [evidence] = await analyseEvidenceFiles([source]);
  assert.equal(evidence.pageSizes?.[0].isA4, false);
  assert.equal(evidence.pageSizes?.[0].wouldAddMarginsOnA4, true);
  const candidate = { id: "letter-candidate", mark: "AH 1", provisionalNumber: 1, description: "Letter source", date: "2026-08-08", paragraph: 2, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: evidence.id, confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "amelia::AH", sequenceOrder: 1 };
  const analysis = { statementName: "Statement.docx", statementHash: "s", caseTitle: "Test", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const converted = await buildBundle(analysis, [candidate]);
  const convertedPdf = await PDFDocument.load(converted.bytes);
  assert.deepEqual(convertedPdf.getPages()[converted.records[0].startPage - 1].getSize(), { width: 595.28, height: 841.89 });
  const convertedRendered = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(converted.bytes) }).promise;
  const convertedText = (await (await convertedRendered.getPage(converted.records[0].startPage)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  assert.match(convertedText, /Letter source page/, "A4 conversion preserves the source PDF text layer instead of rasterising a screenshot");
  await convertedRendered.destroy();
  const retained = await buildBundle(analysis, [candidate], { pageSizeChoices: { [evidence.id]: "keep-original" } });
  const retainedPdf = await PDFDocument.load(retained.bytes);
  assert.deepEqual(retainedPdf.getPages()[retained.records[0].startPage - 1].getSize(), { width: 612, height: 792 });
  assert.equal(retained.manifest.output.pageSize, "mixed");
  assert.match(retained.checks.find((check) => check.label === "A4 page treatment")?.detail ?? "", /reviewer choice/i);
});

test("an existing A4 exhibit remains full-size instead of being inset like a screenshot", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const sourcePdf = await PDFDocument.create();
  sourcePdf.addPage([595.28, 841.89]).drawText("A4 edge reference", { x: 20, y: 100 });
  const source = new File([await sourcePdf.save()], "A4_Source.pdf", { type: "application/pdf" });
  const evidence = { id: "a4", file: source, name: source.name, extension: "pdf", text: "A4 edge reference", marker: null, sha256: "a".repeat(64), pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" };
  const candidate = { id: "a4-candidate", mark: "AH 1", provisionalNumber: 1, description: "A4 source", date: "2026-08-08", paragraph: 2, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: evidence.id, confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "amelia::AH", sequenceOrder: 1 };
  const analysis = { statementName: "Statement.docx", statementHash: "s", caseTitle: "Test", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const built = await buildBundle(analysis, [candidate]);
  const rendered = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(built.bytes) }).promise;
  const items = (await (await rendered.getPage(built.records[0].startPage)).getTextContent()).items;
  const sourceItem = items.find((item) => "str" in item && item.str === "A4 edge reference");
  assert.ok(sourceItem);
  assert.ok(Math.abs(sourceItem.transform[4] - 20) < 0.1, `A4 content was shifted horizontally to ${sourceItem.transform[4]}`);
  assert.ok(Math.abs(sourceItem.transform[5] - 100) < 0.1, `A4 content was shifted vertically to ${sourceItem.transform[5]}`);
  await rendered.destroy();
});

test("preserves PDF annotations when pages can be copied without geometric conversion", async () => {
  const { analyseEvidenceFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  async function annotatedSource(width, height, name) {
    const pdf = await PDFDocument.create();
    const page = pdf.addPage([width, height]);
    page.drawText("Annotated source text", { x: 40, y: 400 });
    const appearance = pdf.context.flateStream("q 1 0 0 RG 1 1 0 rg 0 0 20 20 re B Q", { Type: "XObject", Subtype: "Form", BBox: [0, 0, 20, 20], Resources: {} });
    const appearanceRef = pdf.context.register(appearance);
    const annotation = pdf.context.obj({ Type: "Annot", Subtype: "Square", Rect: [40, 380, 60, 400], Contents: PDFHexString.fromText("Material reviewer note"), AP: { N: appearanceRef } });
    page.node.set(PDFName.of("Annots"), pdf.context.obj([pdf.context.register(annotation)]));
    return new File([await pdf.save()], name, { type: "application/pdf" });
  }
  const candidateFor = (evidence) => ({ id: `candidate-${evidence.id}`, mark: "AH 1", provisionalNumber: 1, description: "Annotated source", date: "2026-08-11", paragraph: 2, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: evidence.id, confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "amelia::AH", sequenceOrder: 1 });
  const analysisFor = (candidate, evidence) => ({ statementName: "Statement.docx", statementHash: "s", caseTitle: "Test", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() });

  const [a4Evidence] = await analyseEvidenceFiles([await annotatedSource(595.28, 841.89, "Annotated_A4.pdf")]);
  assert.deepEqual(a4Evidence.annotationPages, [1]);
  const a4Candidate = candidateFor(a4Evidence);
  const a4Build = await buildBundle(analysisFor(a4Candidate, a4Evidence), [a4Candidate]);
  const a4Output = await PDFDocument.load(a4Build.bytes);
  const a4Page = a4Output.getPages()[a4Build.records[0].startPage - 1];
  assert.equal(a4Page.node.Annots()?.size(), 1, "A4 source annotation remains attached to the copied exhibit page");
  const a4Annotation = a4Output.context.lookup(a4Page.node.Annots().get(0));
  assert.ok(a4Annotation.get(PDFName.of("AP")), "the visible annotation appearance stream is retained");
  const renderedA4 = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(a4Build.bytes) }).promise;
  const displayedAnnotations = await (await renderedA4.getPage(a4Build.records[0].startPage)).getAnnotations({ intent: "display" });
  assert.equal(displayedAnnotations.length, 1);
  assert.equal(displayedAnnotations[0].subtype, "Square");
  await renderedA4.destroy();

  const [letterEvidence] = await analyseEvidenceFiles([await annotatedSource(612, 792, "Annotated_Letter.pdf")]);
  const letterCandidate = candidateFor(letterEvidence);
  const letterAnalysis = analysisFor(letterCandidate, letterEvidence);
  await assert.rejects(buildBundle(letterAnalysis, [letterCandidate]), /annotations that cannot be faithfully repositioned/i);
  const retained = await buildBundle(letterAnalysis, [letterCandidate], { pageSizeChoices: { [letterEvidence.id]: "keep-original" } });
  const retainedOutput = await PDFDocument.load(retained.bytes);
  assert.equal(retainedOutput.getPages()[retained.records[0].startPage - 1].node.Annots()?.size(), 1, "explicitly retained non-A4 page keeps its annotation dictionary");
});

test("blocks active PDF actions but preserves safe internal GoTo annotations", async () => {
  const { analyseEvidenceFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { runPreflight } = await import("../app/lib/preflight.ts");
  const { BUNDLE_PROFILES } = await import("../app/lib/bundle-types.ts");
  const candidateFor = (evidence) => ({ id: `candidate-${evidence.id}`, mark: "AH 1", provisionalNumber: 1, description: "PDF action test", date: "2026-08-13", paragraph: 2, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: evidence.id, confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "amelia::AH", sequenceOrder: 1 });
  const analysisFor = (candidate, evidence) => ({ statementName: "Statement.docx", statementHash: "s", caseTitle: "Test", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() });

  const hostilePdf = await PDFDocument.create();
  const hostilePage = hostilePdf.addPage([595.28, 841.89]);
  hostilePage.drawText("Hostile action source");
  const launch = hostilePdf.context.obj({ Type: "Action", S: "Launch", F: PDFHexString.fromText("calc.exe") });
  hostilePage.node.set(PDFName.of("AA"), hostilePdf.context.obj({ O: hostilePdf.context.register(launch) }));
  const javascript = hostilePdf.context.obj({ Type: "Action", S: "JavaScript", JS: PDFHexString.fromText("app.alert('unsafe')") });
  const hostileAnnotation = hostilePdf.context.obj({ Type: "Annot", Subtype: "Link", Rect: [20, 20, 120, 50], A: hostilePdf.context.register(javascript) });
  hostilePage.node.set(PDFName.of("Annots"), hostilePdf.context.obj([hostilePdf.context.register(hostileAnnotation)]));
  const [hostileEvidence] = await analyseEvidenceFiles([new File([await hostilePdf.save()], "Hostile_Actions.pdf", { type: "application/pdf" })]);
  assert.deepEqual(hostileEvidence.unsafePdfActions?.map(({ page, location, action }) => ({ page, location, action })), [
    { page: 1, location: "page", action: "Launch" },
    { page: 1, location: "annotation", action: "JavaScript" },
  ]);
  const hostileCandidate = candidateFor(hostileEvidence);
  const hostileAnalysis = analysisFor(hostileCandidate, hostileEvidence);
  assert.ok(runPreflight(hostileAnalysis, [hostileCandidate], BUNDLE_PROFILES[0]).some((check) => check.severity === "blocking" && check.label === "Active PDF actions are not permitted"));
  const forgedAnalysis = { ...hostileAnalysis, evidence: [{ ...hostileEvidence, unsafePdfActions: [] }] };
  await assert.rejects(buildBundle(forgedAnalysis, [hostileCandidate]), /contains active PDF actions.*Launch/i, "build-time validation must not trust cached preflight metadata");

  const safePdf = await PDFDocument.create();
  const safeFirst = safePdf.addPage([595.28, 841.89]);
  const safeSecond = safePdf.addPage([595.28, 841.89]);
  safeFirst.drawText("Safe internal link");
  safeSecond.drawText("Safe destination");
  const goTo = safePdf.context.obj({ Type: "Action", S: "GoTo", D: [safeSecond.ref, PDFName.of("Fit")] });
  const safeAnnotation = safePdf.context.obj({ Type: "Annot", Subtype: "Link", Rect: [20, 20, 120, 50], A: safePdf.context.register(goTo) });
  safeFirst.node.set(PDFName.of("Annots"), safePdf.context.obj([safePdf.context.register(safeAnnotation)]));
  const [safeEvidence] = await analyseEvidenceFiles([new File([await safePdf.save()], "Safe_Internal_Link.pdf", { type: "application/pdf" })]);
  assert.deepEqual(safeEvidence.unsafePdfActions, []);
  const safeCandidate = candidateFor(safeEvidence);
  const safeBuild = await buildBundle(analysisFor(safeCandidate, safeEvidence), [safeCandidate]);
  const safeOutput = await PDFDocument.load(safeBuild.bytes);
  const copiedAnnotation = safeOutput.context.lookup(safeOutput.getPages()[safeBuild.records[0].startPage - 1].node.Annots().get(0));
  const copiedAction = safeOutput.context.lookup(copiedAnnotation.get(PDFName.of("A")));
  assert.equal(String(copiedAction.get(PDFName.of("S"))), "/GoTo");
});

test("preserves asymmetric placement through 90, 180 and 270 degree PDF rotations and blocks rotated annotations", async () => {
  const { analyseEvidenceFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  async function rotatedFile(withAnnotation) {
    const pdf = await PDFDocument.create();
    const rotations = withAnnotation ? [90] : [90, 180, 270];
    for (const angle of rotations) {
      const page = pdf.addPage([595.28, 841.89]);
      page.drawText(`Rotated ${angle}`, { x: 40, y: 80 });
      page.setRotation(degrees(angle));
      if (withAnnotation) {
        const annotation = pdf.context.obj({ Type: "Annot", Subtype: "Text", Rect: [40, 60, 60, 80], Contents: PDFHexString.fromText("Rotated note") });
        page.node.set(PDFName.of("Annots"), pdf.context.obj([pdf.context.register(annotation)]));
      }
    }
    return new File([await pdf.save()], withAnnotation ? "Rotated_Annotated.pdf" : "Rotated.pdf", { type: "application/pdf" });
  }
  const make = (evidence) => {
    const candidate = { id: `candidate-${evidence.id}`, mark: "AH 1", provisionalNumber: 1, description: "Rotated source", date: "2026-08-11", paragraph: 2, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: evidence.id, confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "amelia::AH", sequenceOrder: 1 };
    return { candidate, analysis: { statementName: "Statement.docx", statementHash: "s", caseTitle: "Test", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() } };
  };
  const [rotatedEvidence] = await analyseEvidenceFiles([await rotatedFile(false)]);
  assert.deepEqual(rotatedEvidence.rotationPages, [1, 2, 3]);
  const rotated = make(rotatedEvidence);
  const built = await buildBundle(rotated.analysis, [rotated.candidate]);
  const rendered = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(built.bytes) }).promise;
  const expected = [
    { angle: 90, x: 841.89 - 80, y: 40, landscape: true },
    { angle: 180, x: 595.28 - 40, y: 841.89 - 80, landscape: false },
    { angle: 270, x: 80, y: 595.28 - 40, landscape: true },
  ];
  for (const [offset, item] of expected.entries()) {
    const page = await rendered.getPage(built.records[0].startPage + offset);
    const texts = (await page.getTextContent()).items;
    const label = texts.find((entry) => "str" in entry && entry.str === `Rotated ${item.angle}`);
    assert.ok(label, `rotation ${item.angle} retains its text layer`);
    assert.ok(Math.abs(label.transform[4] - item.x) < 1, `rotation ${item.angle} x-position drifted to ${label.transform[4]}`);
    assert.ok(Math.abs(label.transform[5] - item.y) < 1, `rotation ${item.angle} y-position drifted to ${label.transform[5]}`);
    const viewport = page.getViewport({ scale: 1 });
    assert.equal(viewport.width > viewport.height, item.landscape, `rotation ${item.angle} has the expected A4 orientation`);
  }
  await rendered.destroy();

  const [annotatedEvidence] = await analyseEvidenceFiles([await rotatedFile(true)]);
  const annotated = make(annotatedEvidence);
  await assert.rejects(buildBundle(annotated.analysis, [annotated.candidate]), /Rotated PDF annotations cannot be preserved safely/i);
});

test("built-in divider wraps long unbroken filenames inside the printable page", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const sourcePdf = await PDFDocument.create();
  sourcePdf.addPage([595.28, 841.89]).drawText("Source page");
  const source = new File([await sourcePdf.save()], "source.pdf", { type: "application/pdf" });
  const evidence = { id: "source", file: source, name: source.name, extension: "pdf", text: "Source page", marker: null, sha256: "d".repeat(64), pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" };
  const longName = `01_${"Witness_Statement_Lucia_Varela_Adversarial_".repeat(8)}.docx`;
  const candidate = { id: "long", mark: "LV 1", provisionalNumber: 1, description: "Source", date: "2026-08-08", paragraph: 2, citation: "[LV1/xx]", exhibitInitials: "LV", exhibitSequence: 1, discoverySignals: [], evidenceId: evidence.id, confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "LV", witnessKey: "lucia::LV", statementName: longName, sequenceOrder: 1 };
  const analysis = { statementName: longName, statementHash: "s", caseTitle: "Test", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const built = await buildBundle(analysis, [candidate], { layout: { includeDividerPages: true, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 } });
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const rendered = await pdfjs.getDocument({ data: new Uint8Array(built.bytes) }).promise;
  const divider = await rendered.getPage(3);
  const viewport = divider.getViewport({ scale: 1 });
  const items = (await divider.getTextContent()).items.filter((item) => "str" in item && item.str.includes("Witness"));
  assert.ok(items.length > 1, "long filename is wrapped over several lines");
  for (const item of items) assert.ok(item.transform[4] + item.width <= viewport.width - 40, `divider text exceeds right margin: ${item.str}`);
  await rendered.destroy();
});

test("exhibit-and-page numbering uses the promised separator in the PDF and index", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const sourcePdf = await PDFDocument.create();
  sourcePdf.addPage([595.28, 841.89]).drawText("Source page");
  const source = new File([await sourcePdf.save()], "source.pdf", { type: "application/pdf" });
  const evidence = { id: "source", file: source, name: source.name, extension: "pdf", text: "Source page", marker: null, sha256: "n".repeat(64), pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" };
  const candidate = { id: "numbered", mark: "AH 1", provisionalNumber: 1, description: "Numbered source", date: "2026-08-08", paragraph: 2, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: evidence.id, confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "amelia::AH", sequenceOrder: 1 };
  const analysis = { statementName: "Statement.docx", statementHash: "s", caseTitle: "Test", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const built = await buildBundle(analysis, [candidate], { pagination: { matchPdfPageOrder: false, scheme: "section", includePrefixInIndex: false } });
  const rendered = await (await import("pdfjs-dist/legacy/build/pdf.mjs")).getDocument({ data: new Uint8Array(built.bytes) }).promise;
  const indexText = (await (await rendered.getPage(2)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  const contentText = (await (await rendered.getPage(3)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
  assert.match(indexText, /AH1-1/);
  assert.match(indexText, /2026-08-08/, "built-in index prints the reviewer-editable document date");
  assert.match(contentText, /AH1-1/);
  assert.doesNotMatch(contentText, /AH11/);
  await rendered.destroy();
});

async function matterCoverFile(lines) {
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([595.28, 841.89]);
  lines.forEach((line, index) => page.drawText(line, { x: 54, y: 780 - index * 28, size: 12 }));
  const bytes = await pdf.save();
  return new File([bytes], "Cover.pdf", { type: "application/pdf" });
}

async function reviewedMatterTemplate(slot, file, confirmationExtra = {}) {
  const { reviewTemplateMatterPdf } = await import("../app/lib/template-matter-review.ts");
  const bytes = new Uint8Array(await file.arrayBuffer());
  const review = await reviewTemplateMatterPdf(bytes, file.name);
  const confirmation = { pdfSha256: review.pdfSha256, confirmedAt: "2026-08-14T10:00:00.000Z", ...confirmationExtra };
  return {
    slot,
    file,
    sha256: review.pdfSha256,
    sourceFormat: "pdf",
    pdfFile: file,
    pdfSha256: review.pdfSha256,
    review: review,
    reviewState: { matterReview: review, matterConfirmation: confirmation },
  };
}

async function oneExhibitBundle() {
  const pdf = await PDFDocument.create();
  pdf.addPage([595.28, 841.89]).drawText("EXHIBIT-SOURCE", { x: 72, y: 720, size: 14 });
  const bytes = await pdf.save();
  const file = new File([bytes], "source.pdf", { type: "application/pdf" });
  const sha256 = createHash("sha256").update(bytes).digest("hex");
  const evidence = { id: "source", file, name: file.name, extension: "pdf", text: "EXHIBIT-SOURCE", marker: null, sha256, pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" };
  const candidate = { id: "c", mark: "AH 1", provisionalNumber: 1, description: "Source exhibit", date: "2026-08-14", paragraph: 2, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: "source", confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "AH", sequenceOrder: 1 };
  const analysis = { statementName: "Statement.docx", statementHash: "s".repeat(64), caseTitle: "New matter", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  return { analysis, candidate, evidence, sourceBytes: bytes, sourceHash: sha256 };
}

async function pageItems(bytes, pageNumber) {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const document = await pdfjs.getDocument({ data: new Uint8Array(bytes) }).promise;
  const page = await document.getPage(pageNumber);
  const items = (await page.getTextContent()).items.filter((item) => "str" in item);
  await document.destroy();
  return items;
}

test("copied A4 exhibit pages keep the source fonts", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const sourcePdf = await PDFDocument.create();
  const page = sourcePdf.addPage([595.28, 841.89]);
  const times = await sourcePdf.embedFont(StandardFonts.TimesRoman);
  const bold = await sourcePdf.embedFont(StandardFonts.TimesRomanBold);
  page.drawText("Copied roman body", { x: 72, y: 720, size: 14, font: times });
  page.drawText("Copied bold heading", { x: 72, y: 700, size: 14, font: bold });
  const bytes = await sourcePdf.save();
  const file = new File([bytes], "typed.pdf", { type: "application/pdf" });
  const sha256 = createHash("sha256").update(bytes).digest("hex");
  const evidence = { id: "typed", file, name: file.name, extension: "pdf", text: "Copied roman body Copied bold heading", marker: null, sha256, pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" };
  const candidate = { id: "c", mark: "AH 1", provisionalNumber: 1, description: "Typed exhibit", date: "Date not stated", paragraph: 2, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: "typed", confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "AH", sequenceOrder: 1 };
  const analysis = { statementName: "Statement.docx", statementHash: "s".repeat(64), caseTitle: "New matter", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const built = await buildBundle(analysis, [candidate], {
    layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 },
  });
  const reopened = await PDFDocument.load(built.bytes);
  const exhibitPage = reopened.getPages()[built.records[0].startPage - 1];
  const fontDict = exhibitPage.node.Resources()?.lookup(PDFName.of("Font"));
  assert.ok(fontDict, "copied A4 page still has a font resource dictionary");
  const baseFonts = fontDict.keys().map((key) => String(reopened.context.lookup(fontDict.get(key)).get(PDFName.of("BaseFont"))));
  assert.ok(baseFonts.some((name) => /Times-Roman/.test(name)), `source Times-Roman must survive page copy, got ${baseFonts.join(", ")}`);
  assert.ok(baseFonts.some((name) => /Times-Bold/.test(name)), `source Times-Bold must survive page copy, got ${baseFonts.join(", ")}`);
  const items = await pageItems(built.bytes, built.records[0].startPage);
  assert.ok(items.some((item) => item.str.includes("roman body")));
  assert.ok(items.some((item) => item.str.includes("bold heading")));
});

test("finish-mode cover prints the confirmed party name and leaves exhibit bytes unchanged", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { analysis, candidate, sourceBytes, sourceHash } = await oneExhibitBundle();
  const coverFile = await matterCoverFile([
    "IN THE ARBITRATION BETWEEN",
    "Alpha Limited",
    "AND",
    "Beta Limited",
  ]);
  const reviewed = await reviewedMatterTemplate("cover", coverFile);
  const party = reviewed.review.partyNames.find((finding) => finding.value === "Alpha Limited");
  assert.ok(party?.id && party.geometry, "the cover finding must carry printable geometry");
  const cover = {
    ...reviewed,
    reviewState: {
      ...reviewed.reviewState,
      matterConfirmation: {
        pdfSha256: reviewed.pdfSha256,
        confirmedAt: "2026-08-14T10:00:00.000Z",
        partyNames: ["Amended Claimant Limited", "Beta Limited"],
        patches: [{ findingId: party.id, value: "Amended Claimant Limited" }],
      },
    },
  };
  const built = await buildBundle(analysis, [candidate], { templates: [cover], layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0, coverInsertion: "fit-a4" } });
  const items = await pageItems(built.bytes, 1);
  const text = items.map((item) => item.str).join(" ");
  assert.match(text, /Amended Claimant Limited/);
  const atParty = items.filter((item) => item.str.trim() && Math.abs(item.transform[5] - party.geometry.y) < 3);
  assert.equal(atParty.at(-1)?.str, "Amended Claimant Limited");
  assert.equal(built.manifest.templates[0].matterConfirmed, true);
  assert.deepEqual(built.manifest.templates[0].confirmedPartyNames, ["Amended Claimant Limited", "Beta Limited"]);
  assert.equal(built.manifest.exhibits[0].sourceHash, sourceHash);
  assert.equal(built.manifest.statement.modified, false);
  const sourceAfter = new Uint8Array(await analysis.evidence[0].file.arrayBuffer());
  assert.equal(createHash("sha256").update(sourceAfter).digest("hex"), sourceHash);
  assert.deepEqual(sourceAfter, sourceBytes);
});

test("finish-from-template cover write-back recentres a shorter party name", async () => {
  const { applyTemplateMatterPatches, inferCoverTextAlignment } = await import("../app/lib/template-matter-writeback.ts");
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([595.28, 841.89]);
  const font = await pdf.embedFont(StandardFonts.HelveticaBold);
  const original = "MOORLAND BATTERY STORAGE LIMITED";
  const size = 12;
  const originalWidth = font.widthOfTextAtSize(original, size);
  const originalX = (595.28 - originalWidth) / 2;
  const originalY = 700;
  page.drawText(original, { x: originalX, y: originalY, size, font });
  const review = {
    sourceName: "Cover.pdf",
    pdfSha256: "a".repeat(64),
    exactByteLength: 1,
    pageCount: 1,
    extractedCharacterCount: original.length,
    textReliability: "reliable",
    requiresVisualConfirmation: true,
    notice: "test",
    matterNumbers: [],
    partyNames: [{
      id: "party-1",
      kind: "party-name",
      value: original,
      normalizedValue: original.toLowerCase(),
      pageNumbers: [1],
      evidence: [original],
      geometry: { pageNumber: 1, x: originalX, y: originalY, width: originalWidth, height: size, fontSize: size },
      unverified: true,
    }],
    forums: [],
    matterTitles: [],
    placeholders: [],
  };
  applyTemplateMatterPatches(
    [page],
    [{ pageIndex: 0, sourcePageNumber: 1, scale: 1, offsetX: 0, offsetY: 0 }],
    review,
    [{ findingId: "party-1", value: "MOORLANDDDD" }],
    { regular: font, bold: font },
  );
  const items = await pageItems(await pdf.save(), 1);
  const replacement = items.find((item) => item.str === "MOORLANDDDD");
  assert.ok(replacement, "the shorter party name is drawn onto the finished cover");
  const replacementMid = replacement.transform[4] + replacement.width / 2;
  assert.ok(Math.abs(replacementMid - 595.28 / 2) < 8, `recentred midpoint was ${replacementMid}`);
  const atParty = items.filter((item) => item.str.trim() && Math.abs(item.transform[5] - originalY) < 3);
  assert.equal(atParty.at(-1)?.str, "MOORLANDDDD");
  assert.equal(inferCoverTextAlignment({ x: originalX, width: originalWidth }, 595.28), "center");
  assert.equal(inferCoverTextAlignment({ x: 54, width: 400 }, 595.28), "left", "a long left-aligned name is not treated as centred");
});

test("index rows and suggestions use the stamp printed on the finished PDF", async () => {
  const { buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const { analysis, candidate } = await oneExhibitBundle();
  const built = await buildBundle(analysis, [candidate], {
    pagination: { matchPdfPageOrder: true, prefix: "SS", suffix: "", padding: 0, includePrefixInIndex: false, countTemplates: true },
    layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 },
  });
  const exhibitItems = await pageItems(built.bytes, 3);
  const stamp = exhibitItems.map((item) => item.str).find((text) => /^SS\d+$/.test(text));
  assert.equal(stamp, "SS3");
  const indexText = (await pageItems(built.bytes, 2)).map((item) => item.str).join(" ");
  assert.match(indexText, new RegExp(`\\b${stamp}\\b`));
  assert.match(indexText, /2026-08-14/, "built-in index includes the document date");
  assert.equal(buildStatementUpdateSuggestions(built.records)[0].line, `Paragraph 2 - [AH1/${stamp}]`);
  assert.equal(built.records[0].exhibitPageLabelStart, stamp);
});

test("a custom index template receives a Date column only when that template already has one", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { analysis, candidate } = await oneExhibitBundle();
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([595.28, 841.89]);
  page.drawText("No.", { x: 60, y: 700, size: 10 });
  page.drawText("Date", { x: 99, y: 700, size: 10 });
  page.drawText("Description", { x: 190, y: 700, size: 10 });
  page.drawText("Page", { x: 456, y: 700, size: 10 });
  const indexFile = new File([await pdf.save()], "DatedIndex.pdf", { type: "application/pdf" });
  const reviewed = await reviewedMatterTemplate("index", indexFile);
  const built = await buildBundle(analysis, [candidate], {
    templates: [reviewed],
    layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 },
  });
  const items = await pageItems(built.bytes, 2);
  const dateItem = items.find((item) => item.str === candidate.date);
  assert.ok(dateItem, "detected Date header receives the document date");
  assert.ok(dateItem.transform[4] >= 99 && dateItem.transform[4] < 190, "the date is drawn in the detected Date column");
});

test("exact-cover mode leaves original names even if a stale draft exists", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { analysis, candidate } = await oneExhibitBundle();
  const coverFile = await matterCoverFile([
    "IN THE ARBITRATION BETWEEN",
    "Alpha Limited",
    "AND",
    "Beta Limited",
  ]);
  const reviewed = await reviewedMatterTemplate("cover", coverFile);
  const party = reviewed.review.partyNames.find((finding) => finding.value === "Alpha Limited");
  const cover = {
    ...reviewed,
    reviewState: {
      ...reviewed.reviewState,
      matterConfirmation: {
        pdfSha256: reviewed.pdfSha256,
        confirmedAt: "2026-08-14T10:00:00.000Z",
        partyNames: ["Amended Claimant Limited", "Beta Limited"],
        patches: [{ findingId: party.id, value: "Amended Claimant Limited" }],
      },
    },
  };
  const built = await buildBundle(analysis, [candidate], { templates: [cover], layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0, coverInsertion: "exact", exactCoverPageNumber: false, exactCoverVolumeLabel: false } });
  const items = await pageItems(built.bytes, 1);
  const overlay = items.filter((item) => /bold/i.test(item.fontName || "")).map((item) => item.str).join(" ");
  const text = items.map((item) => item.str).join(" ");
  assert.match(text, /Alpha Limited/);
  assert.doesNotMatch(overlay, /Amended Claimant Limited/);
  assert.doesNotMatch(text, /Amended Claimant Limited/);
});

test("built-in cover prints confirmed matter details instead of New matter", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { analysis, candidate } = await oneExhibitBundle();
  const built = await buildBundle(analysis, [candidate], {
    layout: {
      includeDividerPages: false,
      includeExhibitCoverPages: false,
      countOptionalPagesInReferences: false,
      volumePageLimit: 0,
      builtInMatter: {
        matterNumbers: ["ARB-2026-0099"],
        partyNames: ["Northbridge Renewables Limited", "Meridian Components Limited"],
        forums: ["Singapore International Arbitration Centre"],
        matterTitles: ["Confirmed Matter Title"],
      },
    },
  });
  const text = (await pageItems(built.bytes, 1)).map((item) => item.str).join(" ");
  assert.match(text, /Confirmed Matter Title/);
  assert.match(text, /ARB-2026-0099/);
  assert.match(text, /Northbridge Renewables Limited/);
  assert.doesNotMatch(text, /New matter/);
  assert.equal(built.manifest.caseTitle, "Confirmed Matter Title");
});

test("index matter-text patches do not shift the fixed exhibit-row columns", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { analysis, candidate } = await oneExhibitBundle();
  const pdf = await PDFDocument.create();
  pdf.addPage([595.28, 841.89]).drawText("Alpha Limited v Beta Limited", { x: 54, y: 780, size: 12 });
  const indexFile = new File([await pdf.save()], "Index.pdf", { type: "application/pdf" });
  const reviewed = await reviewedMatterTemplate("index", indexFile);
  const title = reviewed.review.matterTitles[0] ?? reviewed.review.partyNames[0];
  assert.ok(title?.id && title.geometry);
  const index = {
    ...reviewed,
    reviewState: {
      ...reviewed.reviewState,
      matterConfirmation: {
        pdfSha256: reviewed.pdfSha256,
        confirmedAt: "2026-08-14T10:00:00.000Z",
        matterTitles: ["Amended Index Title"],
        patches: [{ findingId: title.id, value: "Amended Index Title" }],
      },
    },
  };
  const built = await buildBundle(analysis, [candidate], { templates: [index], layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 } });
  const items = await pageItems(built.bytes, 2);
  const text = items.map((item) => item.str).join(" ");
  assert.match(text, /Amended Index Title/);
  const atTitle = items.filter((item) => item.str.trim() && Math.abs(item.transform[5] - title.geometry.y) < 3);
  assert.equal(atTitle.at(-1)?.str, "Amended Index Title");
  const indexNumber = items.find((item) => item.str === "1" && Math.abs(item.transform[4] - 64) < 1);
  const description = items.find((item) => item.str.includes("Source exhibit"));
  const pageRange = items.find((item) => /^\d/.test(item.str) && item.transform[4] >= 460);
  assert.ok(indexNumber, "item number stays in the documented first column");
  assert.ok(description && Math.abs(description.transform[4] - 103) < 1, "descriptions stay in the documented middle column");
  assert.ok(pageRange && pageRange.transform[4] + pageRange.width <= 543.5, "page ranges stay in the documented final column");
  assert.doesNotMatch(text, /2026-08-14/, "a three-column custom index does not invent a Date column");
});

test("unmatched exhibit blockers name the statement paragraph, not the bundle mark", async () => {
  const { runPreflight } = await import("../app/lib/preflight.ts");
  const { BUNDLE_PROFILES } = await import("../app/lib/bundle-types.ts");
  const analysis = {
    statementName: "Witness statement",
    statementHash: "a".repeat(64),
    caseTitle: "Matter",
    candidates: [],
    evidence: [],
    unreferenced: [],
    statementWarnings: [],
    generatedAt: "2026-08-14T00:00:00.000Z",
  };
  const cited = {
    id: "cited-unmatched",
    mark: "LV 1",
    provisionalNumber: 1,
    description: "Letter of 3 March",
    date: "3 March 2026",
    paragraph: 12,
    citation: "[LV1]",
    discoverySignals: [],
    evidenceId: null,
    confidence: 40,
    rationale: "cited",
    included: true,
    confirmed: true,
    witnessInitials: "LV",
    exhibitInitials: "LV",
    exhibitSequence: 1,
  };
  const citedCheck = runPreflight(analysis, [cited], BUNDLE_PROFILES[0]).find((check) => check.label === "Unmatched exhibit");
  assert.equal(citedCheck?.detail, "Paragraph 12 has no selected source file.");
  assert.equal(citedCheck?.candidateId, "cited-unmatched");
  assert.doesNotMatch(citedCheck?.detail ?? "", /LV/);

  const multi = runPreflight(analysis, [{ ...cited, id: "multi-unmatched", citationCount: 3, citationOrdinal: 2 }], BUNDLE_PROFILES[0]).find((check) => check.label === "Unmatched exhibit");
  assert.equal(multi?.detail, "Paragraph 12, reference 2 of 3, has no selected source file.");
  assert.doesNotMatch(multi?.detail ?? "", /LV/);

  const manual = runPreflight(analysis, [{
    ...cited,
    id: "manual-unmatched",
    paragraph: 0,
    citation: "",
    manualAddition: true,
    description: "Uncited checklist",
  }], BUNDLE_PROFILES[0]).find((check) => check.label === "Unmatched exhibit");
  assert.equal(manual?.detail, "Uncited checklist has no selected source file.");
  assert.doesNotMatch(manual?.detail ?? "", /LV/);

  const unnamedManual = runPreflight(analysis, [{
    ...cited,
    id: "unnamed-manual-unmatched",
    paragraph: 0,
    citation: "",
    manualAddition: true,
    description: "  ",
  }], BUNDLE_PROFILES[0]).find((check) => check.label === "Unmatched exhibit");
  assert.equal(unnamedManual?.detail, "This added exhibit has no selected source file.");
});

test("empty index headings stay saved but do not print", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { analysis, candidate } = await oneExhibitBundle();
  const arrangement = {
    version: 1,
    nodes: [
      { type: "exhibit", exhibitId: candidate.id },
      { type: "section", id: "empty", heading: "Unused agreements heading", exhibits: [] },
    ],
  };
  const built = await buildBundle(analysis, [candidate], {
    arrangement,
    layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 0 },
  });
  const items = await pageItems(built.bytes, 2);
  const indexText = items.map((item) => item.str).join(" ");
  assert.match(indexText, /Source exhibit/);
  assert.doesNotMatch(indexText, /Unused agreements heading/);
});
