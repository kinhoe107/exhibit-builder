import assert from "node:assert/strict";
import { createHash, webcrypto } from "node:crypto";
import { basename, join } from "node:path";
import { tmpdir } from "node:os";
import { createRequire } from "node:module";
import { mkdir, readFile, readdir, writeFile } from "node:fs/promises";
import { File } from "node:buffer";
import test from "node:test";
import { PDFDocument, PDFName, StandardFonts, rgb } from "pdf-lib";

globalThis.File = File;
globalThis.crypto ??= webcrypto;
const require = createRequire(import.meta.url);
const { exportWorkbookSheets } = require("../electron/workbook-export.cjs");

const repoRoot = new URL("../../", import.meta.url);
const packRoot = new URL("../../Sample_Pack_ICC_Adversarial/", import.meta.url);
const outputRoot = new URL("../../Sample_Pack_ICC_Adversarial/Test_Output/", import.meta.url);

async function fixture(url) {
  return new File([await readFile(url)], basename(url.pathname));
}

function sha(bytes) { return createHash("sha256").update(bytes).digest("hex"); }

function reviewedPdfTemplate(slot, file, bytes) {
  const pdfSha256 = sha(bytes);
  const confirmation = { pdfSha256, confirmedAt: "2026-08-13T10:00:00.000Z" };
  return { slot, file, sha256: pdfSha256, sourceFormat: "pdf", pdfFile: file, pdfSha256, reviewState: { matterReview: { sourceName: file.name, pdfSha256, exactByteLength: bytes.byteLength, pageCount: 1, extractedCharacterCount: 0, textReliability: "none", requiresVisualConfirmation: true, notice: "Adversarial fixture visually reviewed.", matterNumbers: [], partyNames: [], forums: [], matterTitles: [], placeholders: [] }, matterConfirmation: confirmation } };
}

async function testWorkbookExporter(_file, sheets) {
  return Promise.all(sheets.map(async (sheet) => {
    const document = await PDFDocument.create();
    const font = await document.embedFont(StandardFonts.Helvetica);
    const page = document.addPage([841.89, 595.28]);
    page.drawRectangle({ x: 48, y: 500, width: 745, height: 34, color: rgb(0.09, 0.21, 0.36) });
    page.drawText(`${sheet.name} - native print test double`, { x: 60, y: 512, size: 12, font, color: rgb(1, 1, 1) });
    for (let pageIndex = 1; pageIndex < (sheet.expectedPageCount ?? 1); pageIndex += 1) document.addPage([841.89, 595.28]).drawText(`${sheet.name} - native print test double - page ${pageIndex + 1}`, { x: 60, y: 512, size: 12, font });
    return { name: sheet.name, range: sheet.range, bytes: await document.save() };
  }));
}

async function nativeWorkbookExporter(file, sheets) {
  return exportWorkbookSheets(tmpdir(), { fileName: file.name, bytes: new Uint8Array(await file.arrayBuffer()), sheets });
}

test("runs the ICC-style adversarial pack through analysis and bundle assembly", async () => {
  const { analyseBundleStatements, buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const statementUrl = new URL("Witness_Statement/01_Witness_Statement_Lucia_Varela_Adversarial.docx", packRoot);
  const statementBytesBefore = await readFile(statementUrl);
  const statementHashBefore = sha(statementBytesBefore);
  const statement = new File([statementBytesBefore], basename(statementUrl.pathname));
  const exhibitNames = (await readdir(new URL("Exhibits/", packRoot)))
    .filter((name) => /\.(pdf|docx|eml|xlsx)$/i.test(name))
    .filter((name) => !/^Exhibit_Bundle(?:\.|_)/i.test(name))
    .filter((name) => !/_Build_Manifest/i.test(name))
    .filter((name) => !/^Statement_Update_Suggestions/i.test(name))
    .sort();
  const exhibitsDir = decodeURIComponent(new URL("Exhibits/", packRoot).pathname);
  assert.ok(!/Test_Output/i.test(exhibitsDir), "private-pack analysis must not read Test_Output");
  assert.ok(!exhibitNames.some((name) => /^Exhibit_Bundle(?:\.|_)/i.test(name)), "a previously generated bundle must never be analysed as source evidence");
  assert.ok(!exhibitNames.some((name) => /(?:Build_Manifest|Statement_Update_Suggestions)/i.test(name)), "generated manifests and suggestion files must never be analysed as source evidence");
  const evidence = await Promise.all(exhibitNames.map((name) => fixture(new URL(`Exhibits/${encodeURIComponent(name)}`, packRoot))));
  const analysis = await analyseBundleStatements([{ id: "lucia", file: statement, witnessName: "Witness Statement Lucia Varela Adversarial", witnessInitials: "VA" }], evidence);
  assert.equal(analysis.caseTitle, "New matter", "the neutral fallback matter title contains no proceeding type");
  assert.equal(analysis.witnessInitials, "LV", "the complete LV1 placeholders establish the one bundle mark");
  assert.ok(analysis.evidence.length >= 31, `expected 31+ source files, got ${analysis.evidence.length}`);
  assert.ok(analysis.candidates.length >= 15, `expected at least 15 citation candidates, got ${analysis.candidates.length}`);
  assert.ok(analysis.candidates.some((c) => c.discoverySignals.some((s) => /I refer to|exhibited|copy/i.test(s))), "citation-language discovery should be present");
  assert.ok(analysis.evidence.some((e) => e.name.startsWith("Scanned_Site_Instruction") && e.pageCount === 1), "scanned OCR fixture should be admitted");
  assert.ok(analysis.evidence.some((e) => e.name.includes("EPC_Contract_Exact_Copy") && e.sha256 === analysis.evidence.find((x) => x.name === "EPC_Contract_Executed_2025-11-18.pdf")?.sha256), "exact duplicate should hash-match");
  assert.ok(analysis.evidence.some((e) => e.name.includes("EPC_Contract_Near_Duplicate") && e.sha256 !== analysis.evidence.find((x) => x.name === "EPC_Contract_Executed_2025-11-18.pdf")?.sha256), "near duplicate should hash-differ");
  const protectionEmail = analysis.evidence.find((e) => e.name === "Email_chain_Protection_Settings_2026-04-12.eml");
  const informalEmail = analysis.evidence.find((e) => e.name === "Email_chain_Informal_Site_Status_2026-04-21.eml");
  assert.equal(protectionEmail?.emailAttachments?.length, 2, "cited-with-attachments email should expose two children");
  assert.equal(informalEmail?.emailAttachments?.length, 1, "cited-without-attachments email should expose one child");
  assert.deepEqual(protectionEmail.emailAttachments.map((child) => child.name).sort(), ["Protection_Test_Certificate.pdf", "Relay_Schedule_RevB.pdf"]);
  assert.equal(informalEmail.emailAttachments[0].name, "Draft_Site_Photo_Log.pdf");
  const tokenByParagraph = new Map(analysis.candidates.map((candidate) => [candidate.paragraph, candidate.citationToken]));
  assert.equal(tokenByParagraph.get(45), "(LV-xx)");
  assert.match(tokenByParagraph.get(46) ?? "", /^\(exhibit\)$/i);
  assert.equal(tokenByParagraph.get(47), "[LV-25]");
  const staleFilled = analysis.candidates.find((candidate) => candidate.paragraph === 47);
  assert.ok(staleFilled);
  assert.notEqual(staleFilled.provisionalNumber, 25, "stale [LV-25] must not become exhibit 25");
  assert.equal(staleFilled.witnessInitials, "LV");
  assert.equal(staleFilled.exhibitSequence, 1, "stale hyphenated mark must inherit the LV1 bundle, not create LV-25");

  const paragraph18 = analysis.candidates.filter((candidate) => candidate.paragraph === 18);
  assert.equal(paragraph18.length, 2, "paragraph 18 is a payment ledger and a milestone invoice");
  assert.equal(paragraph18.find((candidate) => candidate.citationToken)?.description, "payment ledger");
  assert.equal(paragraph18.find((candidate) => !candidate.citationToken)?.description, "milestone invoice");
  assert.equal(paragraph18.find((candidate) => !candidate.citationToken)?.date, "28 February 2026");

  const byName = new Map(analysis.evidence.map((e) => [e.name, e]));
  const mapping = new Map([
    [6, "EPC_Contract_Executed_2025-11-18.pdf"],
    [10, "Calder_Grid_Connection_Offer_2025-10-14.pdf"],
    [13, "Email_chain_Battery_Delivery_2026-01-08.eml"],
    [14, "Email_chain_Battery_Delivery_2026-01-08.eml"],
    [16, "Baseline_Programme_RevC_2025-12-05.xlsx"],
    [20, "Weather_Station_Readings_Q1_2026.xlsx"],
    [26, "Independent_Engineer_Report_2026-04-02.pdf"],
    [28, "Email_chain_Variation_Dispute_2026-03-03.eml"],
    [33, "EPC_Contract_Executed_2025-11-18.pdf"],
    [36, "Damages_Calculation_to_2026-04-30.xlsx"],
    [43, "Email_chain_Protection_Settings_2026-04-12.eml"],
    [44, "Email_chain_Informal_Site_Status_2026-04-21.eml"],
    [45, "Notice_to_Proceed_2025-12-01.pdf"],
    [46, "Variation_Order_01_Grid_Protection_2026-01-22.pdf"],
    [47, "Termination_Notice_2026-04-30.pdf"],
  ]);
  function mappedSource(candidate) {
    if (candidate.paragraph === 18) {
      return candidate.citationToken ? "Payment_Ledger_to_2026-03-31.xlsx" : "Milestone_Invoice_0461_2026-02-28.pdf";
    }
    return mapping.get(candidate.paragraph);
  }
  const unsafeAutomaticSelections = analysis.candidates.flatMap((candidate) => {
    if (!candidate.evidenceId) return [];
    const selectedName = analysis.evidence.find((record) => record.id === candidate.evidenceId)?.name;
    const expectedName = mappedSource(candidate);
    return expectedName && selectedName !== expectedName
      ? [{ paragraph: candidate.paragraph, description: candidate.description, selectedName, expectedName, confidence: candidate.confidence }]
      : [];
  });
  assert.deepEqual(unsafeAutomaticSelections, [], `automatic matching must leave a citation unselected rather than propose the wrong private-pack source: ${JSON.stringify(unsafeAutomaticSelections)}`);
  const candidates = analysis.candidates.map((candidate) => {
    const target = mappedSource(candidate);
    if (!target) return { ...candidate, included: false, confirmed: false };
    const evidenceRecord = byName.get(target);
    assert.ok(evidenceRecord, `missing mapped source ${target}`);
    const emailAttachmentDispositions = candidate.paragraph === 43
      ? Object.fromEntries(protectionEmail.emailAttachments.map((child) => [child.identity, "print-with-email"]))
      : candidate.paragraph === 44
        ? Object.fromEntries(informalEmail.emailAttachments.map((child) => [child.identity, "leave-out"]))
        : undefined;
    return {
      ...candidate,
      evidenceId: evidenceRecord.id,
      included: true,
      confirmed: true,
      repeatDecision: [14, 33].includes(candidate.paragraph) ? "same" : candidate.repeatDecision,
      ...(emailAttachmentDispositions ? { emailAttachmentDispositions } : {}),
    };
  });
  const { deriveExhibitGroups, reviewCandidatesForDisplay, reviewItemNumbers } = await import("../app/lib/exhibit-groups.ts");
  const adversarialReviewCandidates = reviewCandidatesForDisplay(candidates, deriveExhibitGroups(analysis, candidates));
  const adversarialReviewNumbers = reviewItemNumbers(adversarialReviewCandidates);
  assert.deepEqual(adversarialReviewCandidates.map((candidate) => adversarialReviewNumbers.get(candidate.id)), adversarialReviewCandidates.map((_candidate, index) => index + 1), "ICC review cards remain globally sequential even when parsed exhibit bundle marks vary");
  const selected = candidates.filter((candidate) => candidate.included);
  assert.equal(new Set(selected.map((candidate) => candidate.id)).size, selected.length);
  const templates = [];
  for (const slot of ["cover", "index"]) {
    const name = slot === "cover" ? "ICC_Style_Front_Cover_Template.pdf" : "ICC_Style_Index_Template.pdf";
    const bytes = await readFile(new URL(`Templates/${name}`, packRoot));
    const file = new File([bytes], name, { type: "application/pdf" });
    templates.push(reviewedPdfTemplate(slot, file, bytes));
  }
  // The mechanical Node harness cannot create the desktop renderer's OCR
  // worker. Keep this end-to-end assembly run non-blocking so it can validate
  // grouping, pagination, index links and output integrity; the application
  // itself retains the neutral profile's OCR blocking rule for release.
  const nativeExcel = process.env.EXHIBIT_BUILDER_NATIVE_EXCEL === "1";
  const workbookExporter = nativeExcel ? nativeWorkbookExporter : testWorkbookExporter;
  const buildOptions = { templates, pagination: { matchPdfPageOrder: true, prefix: "", includePrefixInIndex: false }, profileId: "review-draft", workbookExporter };
  await assert.rejects(
    buildBundle(analysis, selected, buildOptions),
    /This Excel sheet is too wide to print on A4\. Programme is too wide for landscape A4[\s\S]*77%[\s\S]*85%/i,
    "the intentionally over-wide Programme merge is an expected adversarial fidelity blocker",
  );
  const buildCandidates = selected.filter((candidate) => mappedSource(candidate) !== "Baseline_Programme_RevC_2025-12-05.xlsx");
  const build = await buildBundle(analysis, buildCandidates, buildOptions);
  assert.ok(build.records.length >= 8, `expected 8+ individual bundle entries, got ${build.records.length}`);
  assert.deepEqual(build.records.map((record) => record.exhibitNumber), build.records.map((_record, index) => index + 1), "index numbers must follow final canonical order without repeats");
  const repeatGroups = build.records.flatMap((record) => record.statementReferences?.length > 1 ? [record] : []);
  assert.ok(repeatGroups.length >= 2, "repeat source citations should be represented by one record with multiple references");
  const contractRecords = build.records.filter((record) => /EPC_Contract_Executed/.test(record.fileName));
  assert.equal(contractRecords.length, 1, "exact duplicate/repeat source should be included once");
  const protectionRecords = build.records.filter((record) => /Email_chain_Protection_Settings/.test(record.fileName));
  const informalRecords = build.records.filter((record) => /Email_chain_Informal_Site_Status/.test(record.fileName));
  assert.equal(protectionRecords.length, 1, "cited email with attachments should remain one bundle entry");
  assert.equal(informalRecords.length, 1, "cited email without attachments should remain one bundle entry");
  const protectionRecord = protectionRecords[0];
  const informalRecord = informalRecords[0];
  assert.equal(protectionRecord.emailAttachments?.length, 2);
  assert.ok(protectionRecord.emailAttachments.every((child) => child.disposition === "print-with-email"));
  assert.equal(informalRecord.emailAttachments?.length, 1);
  assert.equal(informalRecord.emailAttachments[0].disposition, "leave-out");
  assert.ok(
    (protectionRecord.endPage - protectionRecord.startPage + 1) > (informalRecord.endPage - informalRecord.startPage + 1),
    "printed attachments should extend the protection-settings exhibit beyond an email-only page range",
  );
  assert.ok(build.records.some((record) => record.workbookSheet), "at least one selected XLSX sheet should be rendered into the bundle");
  assert.ok(build.manifest.excludedFiles.some((item) => /UNREFERENCED_Internal_Draft/.test(item.fileName)), "unreferenced draft should stay outside the bundle");
  assert.ok(build.manifest.excludedFiles.some((item) => /EPC_Contract_Exact_Copy/.test(item.fileName)), "exact duplicate should stay outside the bundle");
  assert.equal(build.manifest.statement.modified, false);
  assert.equal(build.manifest.statement.sha256, statementHashBefore);
  const suggestions = buildStatementUpdateSuggestions(build.records);
  const citedSuggestions = suggestions.filter((item) => /^Paragraph /.test(item.line));
  assert.ok(citedSuggestions.some((item) => /^Paragraph 6 - \[LV1\//.test(item.line)), "suggestions should use LV1 bundle notation");
  assert.ok(citedSuggestions.every((suggestion) => build.records.some((record) => {
    if (!record.statementReferences.some((reference) => reference.paragraph === suggestion.paragraph)) return false;
    const range = record.startPage === record.endPage ? `${record.startPage}` : `${record.startPage}-${record.endPage}`;
    return suggestion.line.includes(`[LV1/${range}]`);
  })), "every ICC suggestion must use the final visible PDF/index page range");
  assert.ok(suggestions.every((item) => !/cited at paragraph/i.test(item.line)), "suggestions should not contain index-only citation text");
  assert.ok(citedSuggestions.some((item) => /^Paragraph 47 - \[LV1\//.test(item.line)), "stale [LV-25] must be replaced by the final LV1 page range");
  assert.ok(suggestions.every((item) => !/\[LV-25\]/.test(item.line)), "copy-ready suggestions must not retain the stale filled-in mark");
  const reopened = await PDFDocument.load(build.bytes);
  assert.ok(reopened.getPageCount() > 20, `bundle should be substantial, got ${reopened.getPageCount()} pages`);
  let landscapeA4Pages = 0;
  for (const page of reopened.getPages()) {
    const { width, height } = page.getSize();
    const portrait = Math.abs(width - 595.28) < 0.02 && Math.abs(height - 841.89) < 0.02;
    const landscape = Math.abs(width - 841.89) < 0.02 && Math.abs(height - 595.28) < 0.02;
    if (landscape) landscapeA4Pages += 1;
    assert.ok(portrait || landscape, "every output page should be A4 portrait or A4 landscape");
  }
  assert.equal(build.manifest.output.pageSize, "A4", "manifest should report the physical output as A4");
  assert.equal(build.manifest.output.orientation.landscape, landscapeA4Pages, "manifest landscape count should match the reopened PDF");
  assert.equal(build.manifest.output.orientation.nonA4, 0, "manifest should report no non-A4 pages for the adversarial pack");
  assert.ok(reopened.catalog.get(PDFName.of("Outlines")), "bundle should contain navigation outlines");
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const rendered = await pdfjs.getDocument({ data: new Uint8Array(build.bytes) }).promise;
  const firstExhibitPage = Math.min(...build.records.map((record) => record.startPage));
  const indexItems = [];
  for (let pageNumber = 2; pageNumber < firstExhibitPage; pageNumber += 1) {
    const page = await rendered.getPage(pageNumber);
    indexItems.push(...(await page.getTextContent()).items.filter((item) => "str" in item && item.str.trim()));
  }
  const indexText = indexItems.map((item) => item.str).join(" ");
  assert.doesNotMatch(indexText, /cited at paragraph/i, "reader-facing index should not expose statement paragraph text");
  assert.doesNotMatch(indexText, /LV-\d/, "page-number prefix should not appear in index ranges unless explicitly enabled");
  const generatedNumbers = indexItems.filter((item) => Math.abs(item.transform[4] - 64) < 0.1).map((item) => Number(item.str));
  assert.deepEqual(generatedNumbers, build.records.map((_record, index) => index + 1), "custom-template index numbers stay sequential inside the No. column");
  const generatedDescriptions = indexItems.filter((item) => Math.abs(item.transform[4] - 103) < 0.1);
  assert.ok(generatedDescriptions.length >= build.records.length, "every index description is rendered in the template description column");
  assert.ok(generatedDescriptions.every((item) => item.transform[4] + item.width <= 443.5), "index descriptions do not cross into the page-range column");
  const generatedRanges = indexItems.filter((item) => item.transform[4] >= 460 && item.transform[5] <= 665 && item.transform[5] >= 100);
  assert.ok(generatedRanges.length >= build.records.length, "every page range is rendered in the template page-range column");
  assert.ok(generatedRanges.every((item) => item.transform[4] + item.width <= 543.1), "page ranges do not exceed the template table boundary");
  async function recordText(record) {
    let text = "";
    for (let pageNumber = record.startPage; pageNumber <= record.endPage; pageNumber += 1) {
      const page = await rendered.getPage(pageNumber);
      text += ` ${(await page.getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ")}`;
    }
    return text;
  }
  const protectionText = await recordText(protectionRecord);
  const informalText = await recordText(informalRecord);
  assert.match(protectionText, /RELAY-SCHEDULE-REVB-BODY/);
  assert.match(protectionText, /PROTECTION-TEST-CERTIFICATE-BODY/);
  assert.doesNotMatch(informalText, /DRAFT-SITE-PHOTO-LOG-BODY/, "left-out attachment must not appear in the informal site-status exhibit");
  if (process.env.EXHIBIT_BUILDER_NATIVE_EXCEL === "1") {
    const nativeWorkbookExpectations = new Map([
      ["Payment_Ledger_to_2026-03-31.xlsx", [/Commissioning milestone/i, /1,820,000/, /Invoice 0461/i]],
      ["Weather_Station_Readings_Q1_2026.xlsx", [/2026-01-31/, /Total rainfall/i, /Maximum wind gust/i, /Wet or damp days/i]],
      ["Damages_Calculation_to_2026-04-30.xlsx", [/Approved access credit/i, /Post-cure period/i, /518,000/]],
    ]);
    for (const record of build.records.filter((item) => item.workbookSheet)) {
      let renderedText = "";
      for (let pageNumber = record.startPage; pageNumber <= record.endPage; pageNumber += 1) {
        const page = await rendered.getPage(pageNumber);
        renderedText += ` ${(await page.getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ")}`;
      }
      for (const pattern of nativeWorkbookExpectations.get(record.fileName) ?? []) {
        assert.match(renderedText, pattern, `${record.fileName} must retain the expected printed cell text`);
      }
    }
  }
  await rendered.destroy();
  const multiVolume = await buildBundle(analysis, buildCandidates, {
    templates,
    pagination: { matchPdfPageOrder: false, prefix: "LV-", includePrefixInIndex: false },
    profileId: "review-draft",
    workbookExporter,
    layout: { includeDividerPages: false, includeExhibitCoverPages: true, countOptionalPagesInReferences: false, volumePageLimit: 45 },
  });
  assert.ok(multiVolume.volumes?.length >= 2, "adversarial fixture should create genuinely separate physical volumes");
  assert.ok(multiVolume.volumes.every((volume) => volume.pageCount <= 45 || volume.checks.some((check) => check.label === "Oversize exhibit volume")), "every adversarial volume should meet the physical limit or retain an explicit oversize advisory");
  assert.ok(buildStatementUpdateSuggestions(multiVolume.records).filter((item) => /^Paragraph /.test(item.line)).every((item) => /\[LV1\/Vol\. \d+\//.test(item.line)), "every multi-volume suggestion, including Volume 1, should identify its volume");
  assert.doesNotMatch(JSON.stringify(multiVolume.manifest), /LV\s*2(?:\D|$)/, "administrative splitting must not create a second witness exhibit bundle");
  for (const volume of multiVolume.volumes) {
    const verified = await PDFDocument.load(volume.bytes);
    assert.equal(verified.getPageCount(), volume.pageCount);
    assert.ok(verified.catalog.get(PDFName.of("Outlines")), `${volume.label} should contain its own bookmarks`);
    const planVolume = multiVolume.buildPlan.volumes.find((item) => item.number === volume.number);
    const internalLinkCount = verified.getPages()
      .slice(planVolume.coverPages, planVolume.coverPages + planVolume.indexPages)
      .reduce((total, page) => total + (page.node.Annots()?.size() ?? 0), 0);
    assert.equal(internalLinkCount, volume.records.length, `${volume.label} links only rows whose exhibits are physically local`);
    const renderedVolume = await pdfjs.getDocument({ data: new Uint8Array(volume.bytes) }).promise;
    let completeIndexText = "";
    for (let pageNumber = planVolume.coverPages + 1; pageNumber <= planVolume.coverPages + planVolume.indexPages; pageNumber += 1) {
      completeIndexText += ` ${(await (await renderedVolume.getPage(pageNumber)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ")}`;
    }
    for (const owningVolume of multiVolume.volumes) assert.match(completeIndexText, new RegExp(`Vol\\. ${owningVolume.number}`), `${volume.label} repeats ownership entries for Volume ${owningVolume.number}`);
    await renderedVolume.destroy();
  }
  const statementBytesAfter = await readFile(statementUrl);
  assert.equal(sha(statementBytesAfter), statementHashBefore, "witness statement must remain byte-identical");
  await mkdir(outputRoot, { recursive: true });
  await writeFile(new URL("Generated_Adversarial_Exhibit_Bundle.pdf", outputRoot), build.bytes);
  await writeFile(new URL("Generated_Adversarial_Build_Manifest.json", outputRoot), JSON.stringify(build.manifest, null, 2));
  await writeFile(new URL("Generated_Adversarial_Multi_Volume.zip", outputRoot), multiVolume.volumeZipBytes);
  await writeFile(new URL("Generated_Adversarial_Multi_Volume_Manifest.json", outputRoot), JSON.stringify(multiVolume.manifest, null, 2));
  await writeFile(new URL("Mechanical_Test_Report.json", outputRoot), JSON.stringify({
    passed: true,
    sourceExhibitCount: analysis.evidence.length,
    candidateCount: analysis.candidates.length,
    selectedCandidateCount: selected.length,
    bundleRecordCount: build.records.length,
    bundlePageCount: build.pageCount,
    outputSha256: build.sha256,
    repeatedRecordCount: repeatGroups.length,
    landscapeA4Pages,
    scannedFixture: analysis.evidence.find((e) => e.name.startsWith("Scanned_Site_Instruction"))?.ocrStatus ?? "unknown",
    suggestions: suggestions.map((item) => item.line),
  }, null, 2));
  console.log(JSON.stringify({ pack: packRoot.pathname, sourceExhibitCount: analysis.evidence.length, candidateCount: analysis.candidates.length, bundleRecords: build.records.length, pageCount: build.pageCount, output: new URL("Generated_Adversarial_Exhibit_Bundle.pdf", outputRoot).pathname }, null, 2));
});

test("amending a name on the ICC cover prints the new wording without touching exhibit bytes", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { reviewTemplateMatterPdf } = await import("../app/lib/template-matter-review.ts");
  const coverBytes = await readFile(new URL("Templates/ICC_Style_Front_Cover_Template.pdf", packRoot));
  const coverFile = new File([coverBytes], "ICC_Style_Front_Cover_Template.pdf", { type: "application/pdf" });
  const review = await reviewTemplateMatterPdf(coverBytes, coverFile.name);
  const finding = [...review.partyNames, ...review.matterTitles, ...review.matterNumbers][0];
  assert.ok(finding?.id && finding.geometry, "the ICC cover fixture must expose at least one printable matter finding");
  const amended = "Amended ICC Party Limited";
  const cover = {
    slot: "cover",
    file: coverFile,
    sha256: review.pdfSha256,
    sourceFormat: "pdf",
    pdfFile: coverFile,
    pdfSha256: review.pdfSha256,
    reviewState: {
      matterReview: review,
      matterConfirmation: {
        pdfSha256: review.pdfSha256,
        confirmedAt: "2026-08-14T10:00:00.000Z",
        partyNames: finding.kind === "party-name" ? [amended] : review.partyNames.map((item) => item.value),
        matterTitles: finding.kind === "matter-title" ? [amended] : review.matterTitles.map((item) => item.value),
        matterNumbers: finding.kind === "matter-number" ? [amended] : review.matterNumbers.map((item) => item.value),
        patches: [{ findingId: finding.id, value: amended }],
      },
    },
  };
  const exhibitPdf = await PDFDocument.create();
  exhibitPdf.addPage([595.28, 841.89]).drawText("ADVERSARIAL-EXHIBIT");
  const exhibitBytes = await exhibitPdf.save();
  const exhibitFile = new File([exhibitBytes], "adversarial-exhibit.pdf", { type: "application/pdf" });
  const sourceHash = sha(exhibitBytes);
  const evidence = { id: "e", file: exhibitFile, name: exhibitFile.name, extension: "pdf", text: "ADVERSARIAL-EXHIBIT", marker: null, sha256: sourceHash, pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" };
  const candidate = { id: "c", mark: "LV 1", provisionalNumber: 1, description: "Adversarial exhibit", date: "2026-08-14", paragraph: 6, citation: "[LV1/xx]", exhibitInitials: "LV", exhibitSequence: 1, discoverySignals: [], evidenceId: "e", confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "LV", witnessKey: "LV", sequenceOrder: 1 };
  const statementBytes = new TextEncoder().encode("Read-only adversarial witness statement fixture");
  const statementFile = new File([statementBytes], "Statement.docx");
  const statementHash = sha(statementBytes);
  const statementInput = { id: "adversarial-statement", file: statementFile, witnessName: "Adversarial witness", witnessInitials: "LV" };
  const analysis = { statementName: statementFile.name, statementHash, statementId: statementInput.id, statementSources: [{ statementId: statementInput.id, fileName: statementFile.name, sha256: statementHash }], statementSnapshots: [statementInput], statementHandles: [statementInput], caseTitle: "New matter", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const built = await buildBundle(analysis, [candidate], { templates: [cover], pagination: { matchPdfPageOrder: true }, profileId: "review-draft" });
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const rendered = await pdfjs.getDocument({ data: new Uint8Array(built.bytes) }).promise;
  const items = (await (await rendered.getPage(1)).getTextContent()).items.filter((item) => "str" in item);
  await rendered.destroy();
  assert.ok(items.some((item) => item.str.includes("Amended ICC Party Limited")), "the finished ICC cover must show the amended wording");
  assert.equal(built.manifest.exhibits[0].sourceHash, sourceHash);
  assert.equal(sha(new Uint8Array(await exhibitFile.arrayBuffer())), sourceHash);
  assert.ok(built.manifest.templates[0].confirmedPartyNames || built.manifest.templates[0].confirmedMatterTitles || built.manifest.templates[0].confirmedMatterNumbers);
});
