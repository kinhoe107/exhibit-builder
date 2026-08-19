import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { basename } from "node:path";
import { File } from "node:buffer";
import test from "node:test";
import JSZip from "jszip";

globalThis.File = File;

test("recognises bracketed witness placeholders in inline-numbered paragraphs", async () => {
  const { analyseBundleStatements } = await import("../app/lib/bundle-engine.ts");
  const path = new URL("./fixtures/moorland/Witness_Statement/01_Witness_Statement_Priya_Nair.docx", import.meta.url);
  const statement = new File([await readFile(path)], basename(path.pathname));
  const analysis = await analyseBundleStatements([
    { id: "priya", file: statement, witnessName: "Priya Nair", witnessInitials: "PN" },
  ], []);

  assert.equal(analysis.candidates.length, 24);
  assert.deepEqual(analysis.candidates.slice(0, 3).map((candidate) => candidate.paragraph), [6, 7, 8]);
  assert.deepEqual(analysis.candidates.slice(0, 3).map((candidate) => candidate.mark), ["PN 1", "PN 2", "PN 3"]);
  assert.ok(analysis.candidates.every((candidate) => candidate.discoverySignals.includes("statement exhibit placeholder or mark")));
  assert.match(analysis.statementWarnings.join(" "), /exhibit placeholders/i);
});

test("accepts different bounded placeholder prefixes and numbering styles", async () => {
  const { analyseFiles } = await import("../app/lib/bundle-engine.ts");
  const zip = new JSZip();
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>1. I refer to the contract [AB-xx].</w:t></w:r></w:p><w:p><w:r><w:t>2. I refer to the email [Witness-XX].</w:t></w:r></w:p><w:p><w:r><w:t>3. I refer to the report [EX 01].</w:t></w:r></w:p><w:p><w:r><w:t>4. I refer to the schedule [AB1].</w:t></w:r></w:p><w:p><w:r><w:t>5. I refer to the appendix [EXH-A].</w:t></w:r></w:p></w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: "uint8array" });
  const analysis = await analyseFiles(new File([bytes], "Generic_Placeholders.docx"), []);

  assert.deepEqual(analysis.candidates.map((candidate) => candidate.paragraph), [1, 2, 3, 4, 5]);
  assert.equal(analysis.candidates.length, 5);
  assert.ok(analysis.candidates.every((candidate) => candidate.discoverySignals.includes("statement exhibit placeholder or mark")));
});

test("one complete reference establishes the bundle mark for incomplete placeholders", async () => {
  const { analyseBundleStatements } = await import("../app/lib/bundle-engine.ts");
  const zip = new JSZip();
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>1. I refer to the contract [LV1/xx].</w:t></w:r></w:p><w:p><w:r><w:t>2. I refer to the spreadsheet [LV-xx].</w:t></w:r></w:p></w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: "uint8array" });
  const statement = new File([bytes], "01_Witness_Statement_Lucia_Varela_Adversarial.docx");
  const analysis = await analyseBundleStatements([{ id: "lucia", file: statement, witnessName: "Witness Statement Lucia Varela Adversarial", witnessInitials: "VA" }], []);

  assert.equal(analysis.witnessInitials, "LV");
  assert.deepEqual(analysis.candidates.map((candidate) => candidate.exhibitInitials), ["LV", "LV"]);
  assert.deepEqual(analysis.candidates.map((candidate) => candidate.exhibitSequence), [1, 1]);
  assert.equal(analysis.candidates[1].citationToken, "[LV-xx]");
  assert.equal(analysis.candidates[1].citationResolution, "none");
});

test("creates separate candidates for several exhibits in the same paragraph", async () => {
  const { analyseFiles, parseStatementCitationTokens } = await import("../app/lib/bundle-engine.ts");
  const paragraph = "2. I also refer to the emails from the claimant [AH-xx; AH-xx; AH-xx]";
  const tokens = parseStatementCitationTokens(paragraph);
  assert.equal(tokens.length, 3);
  assert.deepEqual(tokens.map((token) => token.raw), ["[AH-xx]", "[AH-xx]", "[AH-xx]"]);

  const zip = new JSZip();
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${paragraph}</w:t></w:r></w:p></w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: "uint8array" });
  const evidence = [
    new File(["Claimant confirmation email"], "Claimant_Confirmation_Email.txt"),
    new File(["Claimant delivery email"], "Claimant_Delivery_Email.txt"),
    new File(["Claimant payment email"], "Claimant_Payment_Email.txt"),
  ];
  const analysis = await analyseFiles(new File([bytes], "Several_Exhibits.docx"), evidence);

  assert.equal(analysis.candidates.length, 3);
  assert.deepEqual(analysis.candidates.map((candidate) => candidate.paragraph), [2, 2, 2]);
  assert.equal(new Set(analysis.candidates.map((candidate) => candidate.id)).size, 3);
  assert.ok(analysis.candidates.every((candidate) => candidate.citation === "I also refer to the emails from the claimant [AH-xx; AH-xx; AH-xx]"));
  assert.ok(analysis.candidates.every((candidate) => candidate.evidenceId === null), "a grouped ambiguous paragraph must require review instead of greedily claiming files");
  assert.deepEqual(analysis.candidates.map((candidate) => candidate.citationOrdinal), [1, 2, 3]);
  assert.ok(analysis.candidates.every((candidate) => candidate.citationCount === 3));
});

test("recognises grouped, repeated and plain guided placeholders without ordinary bracketed prose", async () => {
  const { parseStatementCitationTokens } = await import("../app/lib/bundle-engine.ts");
  assert.equal(parseStatementCitationTokens("[AH xx; AH xx]").length, 2);
  assert.equal(parseStatementCitationTokens("[Exhibit]; [Exhibit]").length, 2);
  assert.equal(parseStatementCitationTokens("[Exhib xx]").length, 1);
  assert.equal(parseStatementCitationTokens("[for background only]").length, 0);
});

test("does not split a legal exhibit page range", async () => {
  const { parseStatementCitationTokens } = await import("../app/lib/bundle-engine.ts");
  const tokens = parseStatementCitationTokens("I refer to the report [AH1/12-18].");
  assert.equal(tokens.length, 1);
  assert.equal(tokens[0].exhibitInitials, "AH");
  assert.equal(tokens[0].exhibitSequence, 1);
  assert.equal(tokens[0].requestedExhibitPageStart, 12);
  assert.equal(tokens[0].requestedExhibitPageEnd, 18);
});

test("manual exhibits have no fabricated statement reference and appear as uncited suggestions", async () => {
  const { buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const { deriveExhibitGroups } = await import("../app/lib/exhibit-groups.ts");
  const evidence = { id: "manual-source", file: new File(["sample"], "Manual.txt"), name: "Manual.txt", extension: "txt", text: "sample", marker: null, sha256: "a".repeat(64), pageCount: 0, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" };
  const candidate = { id: "manual", mark: "EX 1", provisionalNumber: 1, description: "Manual document", date: "8 August 2026", paragraph: 0, citation: "", citationResolution: "none", discoverySignals: ["Manually added by reviewer"], evidenceId: evidence.id, confidence: 100, rationale: "Reviewer added", included: true, confirmed: true, witnessInitials: "EX", witnessKey: "general-exhibits::EX", exhibitInitials: "EX", exhibitSequence: 1, manualAddition: true, manualAddedAt: "2026-08-08T00:00:00.000Z", manualWarningAcknowledgedAt: "2026-08-08T00:00:00.000Z" };
  const analysis = { statementName: "Statement.docx", statementHash: "b".repeat(64), caseTitle: "Example", candidates: [candidate], evidence: [evidence], unreferenced: [], statementWarnings: [], generatedAt: "2026-08-08T00:00:00.000Z" };
  const [group] = deriveExhibitGroups(analysis, [candidate]);
  assert.deepEqual(group.references, []);
  assert.deepEqual(buildStatementUpdateSuggestions([{ mark: "EX1", exhibitNumber: 1, description: candidate.description, fileName: evidence.name, startPage: 1, endPage: 1, exhibitPageStart: 1, exhibitPageEnd: 1, statementParagraph: null, statementReferences: [], sourceHash: evidence.sha256, manualAddition: true, citationStatus: "not-cited-manual-addition" }]).map((item) => item.line), [
    "Uncited exhibits — no statement reference",
    "1. Manual document — page 1",
  ]);
});

test("guided sample is instructional and detects six references including a three-reference paragraph", async () => {
  const { analyseBundleStatements, SAMPLE_EVIDENCE, SAMPLE_REQUIRED_FILES, SAMPLE_STATEMENT, SAMPLE_TEMPLATES } = await import("../app/lib/bundle-engine.ts");
  const root = new URL("../public/guided-sample/", import.meta.url);
  const statementUrl = new URL(SAMPLE_STATEMENT, root);
  const statement = new File([await readFile(statementUrl)], basename(statementUrl.pathname));
  const evidence = await Promise.all(SAMPLE_EVIDENCE.map(async (name) => new File([await readFile(new URL(name, root))], name)));
  const analysis = await analyseBundleStatements([{ id: "guided", file: statement, witnessName: "Guided Sample", witnessInitials: "AH" }], evidence);
  assert.equal(analysis.candidates.length, 6);
  assert.deepEqual(analysis.candidates.map((candidate) => candidate.paragraph), [1, 2, 3, 3, 3, 6]);
  assert.equal(analysis.candidates.filter((candidate) => candidate.paragraph === 3).length, 3);
  assert.deepEqual(analysis.candidates.slice(2, 5).map((candidate) => candidate.description), ["Sample project report", "Sample claimant email", "Sample cost workbook"]);
  assert.match(analysis.candidates[0].citation, /\[Exhibit\]/);
  assert.match(analysis.candidates[1].citation, /\[Exhib xx\]/);
  assert.ok(analysis.candidates.slice(2, 5).every((candidate) => candidate.evidenceId === null));
  assert.deepEqual(new Set(analysis.evidence.map((record) => record.extension)), new Set(["pdf", "docx", "eml", "xlsx"]));
  const workbook = analysis.evidence.find((record) => record.name === "05_SAMPLE_Cost_Workbook.xlsx");
  assert.deepEqual(workbook?.workbook?.sheets.map((sheet) => sheet.name), ["Summary - Include", "Detail - Optional", "Working Notes - Exclude"]);
  const email = analysis.evidence.find((record) => record.name === "04_SAMPLE_Claimant_Email.eml");
  assert.equal(email?.emailAttachments?.length, 1);
  assert.equal(email?.emailAttachments?.[0].name, "Guided_Attachment_Note.txt");
  assert.equal(email?.emailAttachments?.[0].supported, true);
  assert.equal(await email?.emailAttachments?.[0].file.text(), "SAMPLE ATTACHMENT");
  assert.ok(analysis.evidence.some((record) => record.name === "06_SAMPLE_Unreferenced_Checklist.pdf"));
  assert.ok(analysis.evidence.every((record) => /SAMPLE DOCUMENT - FOR EXHIBIT BUILDER DEMONSTRATION ONLY/i.test(record.text)));
  assert.deepEqual(SAMPLE_TEMPLATES, [
    { slot: "cover", name: "00_GUIDED_SAMPLE_Cover_Template.pdf" },
    { slot: "index", name: "00_GUIDED_SAMPLE_Index_Template.pdf" },
  ]);
  assert.equal(SAMPLE_REQUIRED_FILES.some((name) => /Cover_Template|Index_Template/.test(name)), false);
  for (const template of SAMPLE_TEMPLATES) assert.ok(await readFile(new URL(template.name, root)));
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const indexTemplate = await pdfjs.getDocument({ data: new Uint8Array(await readFile(new URL(SAMPLE_TEMPLATES[1].name, root))) }).promise;
  const indexText = (await (await indexTemplate.getPage(1)).getTextContent()).items.map((item) => item.str).join(" ");
  assert.match(indexText, /ITEM NO\./);
  assert.match(indexText, /EXHIBIT DESCRIPTION - INSERTED AUTOMATICALLY/);
  assert.match(indexText, /BUNDLE PAGES/);
  const documentZip = await JSZip.loadAsync(await statement.arrayBuffer());
  const documentXml = await documentZip.file("word/document.xml").async("text");
  const numberingXml = await documentZip.file("word/numbering.xml").async("text");
  const coreXml = await documentZip.file("docProps/core.xml").async("text");
  assert.match(documentXml, /This is an instructional guide, not a realistic witness statement/);
  assert.match(documentXml, /Evidence exhibits can be PDF, DOCX, EML or XLSX files/);
  assert.match(documentXml, /Several exhibits can be detected separately in the same paragraph/);
  assert.match(documentXml, /which worksheet tabs and detected cell ranges Microsoft Excel will print into the bundle/);
  assert.match(documentXml, /complete three-column table with placeholder headings/);
  assert.match(documentXml, /<w:numPr>/);
  assert.match(numberingXml, /<w:lvlText w:val="%1\."/);
  assert.doesNotMatch(coreXml, /<dc:creator>[^<]+<\/dc:creator>/);
});

async function analyseParagraphs(paragraphs, evidence = []) {
  const { analyseFiles } = await import("../app/lib/bundle-engine.ts");
  const zip = new JSZip();
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs.map((text, index) => `<w:p><w:r><w:t>${index + 1}. ${text}</w:t></w:r></w:p>`).join("")}</w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: "uint8array" });
  return analyseFiles(new File([bytes], "Parser_Cases.docx"), evidence);
}

test("recognises round and square exhibit tokens, including AHxx without a hyphen", async () => {
  const { parseStatementCitationTokens } = await import("../app/lib/bundle-engine.ts");
  const round = parseStatementCitationTokens("I refer to the email (CC-23).");
  assert.equal(round.length, 1);
  assert.equal(round[0].raw, "(CC-23)");
  const exhibit = parseStatementCitationTokens("The letter is exhibited (Exhibit).");
  assert.equal(exhibit.length, 1);
  assert.equal(exhibit[0].raw, "(Exhibit)");
  const compact = parseStatementCitationTokens("I refer to the invoice (AHxx).");
  assert.equal(compact.length, 1);
  assert.equal(compact[0].raw, "(AHxx)");
  assert.equal(parseStatementCitationTokens("[AHxx]").length, 1);
  assert.equal(parseStatementCitationTokens("[AHxx]")[0].raw, "[AHxx]");
});

test("splits comma and semicolon token lists only when every part is a token", async () => {
  const { parseStatementCitationTokens } = await import("../app/lib/bundle-engine.ts");
  const comma = parseStatementCitationTokens("I refer to the emails (BB-xx, BBxx, BB-x).");
  assert.equal(comma.length, 3);
  assert.deepEqual(comma.map((token) => token.raw), ["(BB-xx)", "(BBxx)", "(BB-x)"]);
  assert.equal(parseStatementCitationTokens("[AH-xx, AH-xx]").length, 2);
  assert.equal(parseStatementCitationTokens("[the contract, AH-xx]").length, 0);
});

test("does not tokenise ordinary asides or split a page range", async () => {
  const { parseStatementCitationTokens } = await import("../app/lib/bundle-engine.ts");
  assert.equal(parseStatementCitationTokens("The email (emphasis added) was hostile.").length, 0);
  assert.equal(parseStatementCitationTokens("See the report [for background only].").length, 0);
  assert.equal(parseStatementCitationTokens("The email chain (page 4) was hostile in tone.").length, 0);
  assert.equal(parseStatementCitationTokens("The meeting (March 3) was short.").length, 0);
  assert.equal(parseStatementCitationTokens("See (para 12) and (COVID-19).").length, 0);
  const range = parseStatementCitationTokens("I refer to the report [AH1/12-18].");
  assert.equal(range.length, 1);
  assert.equal(range[0].raw, "[AH1/12-18]");
});

test("keeps token and narrative wording as one card and adds a later attach sentence by span", async () => {
  const combined = await analyseParagraphs(["I refer to the email (CC-23)."]);
  assert.equal(combined.candidates.length, 1);
  assert.equal(combined.candidates[0].citationToken, "(CC-23)");
  assert.ok(combined.candidates[0].discoverySignals.includes('"I refer to" language'));

  const mixed = await analyseParagraphs(["I refer to the email (CC-23). I attach the March invoices."]);
  assert.equal(mixed.candidates.length, 2);
  assert.equal(mixed.candidates[0].citationToken, "(CC-23)");
  assert.equal(mixed.candidates[1].citationToken, undefined);
  assert.ok(mixed.candidates[1].discoverySignals.includes("attach or enclose language"));
});

test("creates unmatched attach and enclose cards only when the verb governs a document noun", async () => {
  const positives = await analyseParagraphs([
    "I attach the March invoices.",
    "The contract is attached.",
    "I enclose the purchase orders.",
    "The appendices are appended.",
    "I attach Dr. Smith's report.",
  ]);
  assert.equal(positives.candidates.length, 5);
  assert.ok(positives.candidates.every((candidate) => !candidate.citationToken));
  assert.ok(positives.candidates.every((candidate) => candidate.discoverySignals.includes("attach or enclose language")));

  const negatives = await analyseParagraphs([
    "I attached great importance to the timing of the meeting.",
    "The 3 March email chain was hostile in tone.",
    "The invoice was attached to that email.",
  ]);
  assert.equal(negatives.candidates.length, 0);
});

test("does not split an English noun list in one unbracketed sentence", async () => {
  const analysis = await analyseParagraphs(["I attach the contract and the invoice."]);
  assert.equal(analysis.candidates.length, 1);
  assert.equal(analysis.candidates[0].citationToken, undefined);
});

test("splits a token-bearing sentence only at and-to-the plus a document noun", async () => {
  const paragraph = "I also refer to the payment ledger [LV1/xx] and to the milestone invoice dated 28 February 2026.";
  const analysis = await analyseParagraphs([paragraph], [
    new File(["payment ledger rows"], "Payment_Ledger_to_2026-03-31.txt"),
    new File(["milestone invoice dated 28 February 2026"], "Milestone_Invoice_0461_2026-02-28.txt"),
  ]);
  assert.equal(analysis.candidates.length, 2);
  assert.equal(analysis.candidates[0].citationToken, "[LV1/xx]");
  assert.equal(analysis.candidates[0].description, "payment ledger");
  assert.equal(analysis.candidates[0].date, "Date not stated");
  assert.equal(analysis.candidates[1].citationToken, undefined);
  assert.equal(analysis.candidates[1].description, "milestone invoice");
  assert.equal(analysis.candidates[1].date, "28 February 2026");
  assert.ok(analysis.candidates.every((candidate) => candidate.citation === paragraph));
  assert.ok(analysis.candidates.every((candidate) => candidate.citationCount === 2));
  assert.ok(analysis.candidates.every((candidate) => candidate.evidenceId === null), "two candidates from this gated split must not auto-claim files");
});

test("does not split email and invoice without a token or and-to-the clause", async () => {
  const untokenised = await analyseParagraphs(["I refer to the email and invoice."]);
  assert.equal(untokenised.candidates.length, 1);
  assert.equal(untokenised.candidates[0].citationToken, undefined);

  const withToken = await analyseParagraphs(["I refer to the email [LV1/xx] and invoice."]);
  assert.equal(withToken.candidates.length, 1);
  assert.equal(withToken.candidates[0].citationToken, "[LV1/xx]");
});

test("keeps semicolon-bounded descriptions when a sentence is not an and-to-the split", async () => {
  const analysis = await analyseParagraphs(["I refer to the invoice [LV1/xx]; I also refer to the email [LV2/xx]."]);
  assert.equal(analysis.candidates.length, 2);
  assert.equal(analysis.candidates[0].description, "Apex Controls Invoice AC-7782");
  assert.match(analysis.candidates[1].description, /email/i);
  assert.notEqual(analysis.candidates[1].description, analysis.candidates[0].description);
});

test("and-to-the after several tokens keeps each head description on its own token", async () => {
  const analysis = await analyseParagraphs(["I refer to the invoice [LV1/xx]; I also refer to the email [LV2/xx] and to the report dated 1 March 2026."]);
  assert.equal(analysis.candidates.length, 3);
  assert.equal(analysis.candidates[0].citationToken, "[LV1/xx]");
  assert.equal(analysis.candidates[0].description, "Apex Controls Invoice AC-7782");
  assert.equal(analysis.candidates[1].citationToken, "[LV2/xx]");
  assert.match(analysis.candidates[1].description, /email/i);
  assert.equal(analysis.candidates[2].citationToken, undefined);
  assert.equal(analysis.candidates[2].description, "report");
  assert.equal(analysis.candidates[2].date, "1 March 2026");
});
