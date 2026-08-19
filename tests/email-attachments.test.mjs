import assert from "node:assert/strict";
import { createHash, webcrypto } from "node:crypto";
import { File } from "node:buffer";
import test from "node:test";
import JSZip from "jszip";
import { PDFDocument } from "pdf-lib";
import { BUNDLE_PROFILES } from "../app/lib/bundle-types.ts";

globalThis.File = File;
globalThis.crypto ??= webcrypto;

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function base64Lines(bytes) {
  return Buffer.from(bytes).toString("base64").match(/.{1,76}/g)?.join("\r\n") ?? "";
}

function attachmentPart({ name, mimeType, bytes, nested = false }) {
  return [
    `Content-Type: ${mimeType}${name ? `; name="${name}"` : ""}`,
    `Content-Disposition: attachment; filename="${name}"`,
    `Content-Transfer-Encoding: ${nested ? "7bit" : "base64"}`,
    "",
    nested ? Buffer.from(bytes).toString("utf8") : base64Lines(bytes),
  ].join("\r\n");
}

function makeEml({ subject = "Exhibit email", parts, body = "Please see the attached file." }) {
  return [
    "From: Witness <witness@example.test>",
    "To: Counsel <counsel@example.test>",
    "Date: Sat, 15 Aug 2026 10:00:00 +0000",
    `Subject: ${subject}`,
    "MIME-Version: 1.0",
    'Content-Type: multipart/mixed; boundary="BOUND"',
    "",
    "--BOUND",
    "Content-Type: text/plain; charset=utf-8",
    "",
    body,
    ...parts.flatMap((part) => ["--BOUND", part]),
    "--BOUND--",
    "",
  ].join("\r\n");
}

async function pdfBytes(text) {
  const pdf = await PDFDocument.create();
  pdf.addPage([595.28, 841.89]).drawText(text, { x: 72, y: 760, size: 18 });
  return new Uint8Array(await pdf.save());
}

async function statementFile() {
  const zip = new JSZip();
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>1. I refer to the email [AH-xx].</w:t></w:r></w:p></w:body></w:document>`);
  return new File([await zip.generateAsync({ type: "uint8array" })], "Statement.docx");
}

async function syntheticWorkbook() {
  const zip = new JSZip();
  zip.file("xl/workbook.xml", `<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Ledger" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file("xl/worksheets/sheet1.xml", `<worksheet><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Amount</t></is></c></row></sheetData></worksheet>`);
  return new Uint8Array(await zip.generateAsync({ type: "uint8array" }));
}

async function workbookPrintTestDouble(_file, sheets) {
  return Promise.all(sheets.map(async (sheet) => {
    const pdf = await PDFDocument.create();
    pdf.addPage([841.89, 595.28]).drawText(`Excel print ${sheet.name}`, { x: 40, y: 300, size: 18 });
    return { name: sheet.name, range: sheet.range, bytes: await pdf.save() };
  }));
}

test("extracts hash-bound children, keeps duplicate names distinct, and ignores asides of nested parse failure", async () => {
  const { extractEmailChildren, emailChildIdentity, rederiveEmailChildren, unresolvedEmailAttachments } = await import("../app/lib/email-attachments.ts");
  const { parseBundleEmail } = await import("../app/lib/email.ts");
  const firstPdf = await pdfBytes("FIRST-PDF");
  const secondPdf = await pdfBytes("SECOND-PDF");
  const nested = makeEml({
    subject: "Nested",
    parts: [attachmentPart({ name: "inner.pdf", mimeType: "application/pdf", bytes: firstPdf })],
  });
  const raw = makeEml({
    parts: [
      attachmentPart({ name: "same.pdf", mimeType: "application/pdf", bytes: firstPdf }),
      attachmentPart({ name: "same.pdf", mimeType: "application/pdf", bytes: secondPdf }),
      attachmentPart({ name: "forwarded.eml", mimeType: "message/rfc822", bytes: Buffer.from(nested, "utf8") }),
      attachmentPart({ name: "photo.png", mimeType: "image/png", bytes: new Uint8Array([137, 80, 78, 71]) }),
    ],
  });
  const parsed = await parseBundleEmail(raw);
  const parentSha = sha256(Buffer.from(raw, "utf8"));
  const children = await extractEmailChildren(parsed, parentSha);
  assert.equal(children.filter((child) => child.name === "same.pdf").length, 2);
  assert.equal(new Set(children.map((child) => child.identity)).size, children.length);
  assert.ok(children.some((child) => child.ordinal === "1" && child.sha256 === sha256(firstPdf)));
  assert.ok(children.some((child) => child.ordinal === "2" && child.sha256 === sha256(secondPdf)));
  assert.ok(children.some((child) => child.name === "forwarded.eml" && child.extension === "eml"));
  assert.ok(children.some((child) => child.name === "inner.pdf" && child.nested && child.ordinal.startsWith("3.")));
  const photo = children.find((child) => child.name === "photo.png");
  assert.equal(photo.supported, false);
  assert.equal(children.every((child) => child.identity === emailChildIdentity(parentSha, child.ordinal, child.sha256)), true);
  const rederived = await rederiveEmailChildren(new File([raw], "parent.eml"), parentSha);
  assert.deepEqual(rederived.map((child) => child.identity), children.map((child) => child.identity));
  assert.equal(unresolvedEmailAttachments(children, undefined).length, children.length);
  assert.equal(unresolvedEmailAttachments(children, { [photo.identity]: "leave-out" }).some((child) => child.identity === photo.identity), false);
});

test("old projects treat selected emails with attachments as unresolved rather than silently included", async () => {
  const { analyseFiles, attachDerivedEmailEvidence } = await import("../app/lib/bundle-engine.ts");
  const { runPreflight } = await import("../app/lib/preflight.ts");
  const { buildBlockers } = await import("../app/lib/build-readiness.ts");
  const pdf = await pdfBytes("ATTACHED");
  const raw = makeEml({ parts: [attachmentPart({ name: "invoice.pdf", mimeType: "application/pdf", bytes: pdf })] });
  const analysis = await analyseFiles(await statementFile(), [new File([raw], "Notice.eml")]);
  const candidate = { ...analysis.candidates[0], evidenceId: analysis.evidence[0].id, included: true, confirmed: true };
  const checks = runPreflight(analysis, [candidate], BUNDLE_PROFILES[0]);
  const unresolved = checks.find((check) => check.code === "email.attachment_unresolved");
  assert.ok(unresolved);
  assert.match(unresolved.detail, /Notice\.eml/);
  assert.match(unresolved.detail, /invoice\.pdf/);
  assert.equal(unresolved.candidateId, candidate.id);
  const blockers = buildBlockers({ includedCount: 1, confirmedCount: 1, pendingApprovalCount: 0, templateReviewPending: false, preflight: checks });
  assert.ok(blockers.some((blocker) => blocker.candidateId === candidate.id && blocker.actionLabel === "Open the exhibit card"));
  const hydrated = await attachDerivedEmailEvidence(analysis, [{
    ...candidate,
    manualAddition: true,
    parentEmailProvenance: {
      parentName: "Notice.eml",
      parentSha256: analysis.evidence[0].sha256,
      childIdentity: analysis.evidence[0].emailAttachments[0].identity,
      childSha256: analysis.evidence[0].emailAttachments[0].sha256,
    },
  }]);
  assert.ok(hydrated.analysis.evidence.some((record) => record.derivedFromEmail?.childIdentity === analysis.evidence[0].emailAttachments[0].identity));
  assert.equal(hydrated.candidates[0].evidenceId, hydrated.analysis.evidence.find((record) => record.derivedFromEmail)?.id);
});

test("print-with-email extends the same exhibit page range and leave-out omits the child from the PDF", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const pdf = await pdfBytes("ATTACHMENT-BODY");
  const raw = makeEml({ parts: [attachmentPart({ name: "invoice.pdf", mimeType: "application/pdf", bytes: pdf })] });
  const analysis = await analyseFiles(await statementFile(), [new File([raw], "Notice.eml")]);
  const child = analysis.evidence[0].emailAttachments[0];
  const base = { ...analysis.candidates[0], evidenceId: analysis.evidence[0].id, included: true, confirmed: true };
  const omitted = await buildBundle(analysis, [{ ...base, emailAttachmentDispositions: { [child.identity]: "leave-out" } }]);
  const printed = await buildBundle(analysis, [{ ...base, emailAttachmentDispositions: { [child.identity]: "print-with-email" } }]);
  assert.ok(printed.records[0].endPage > omitted.records[0].endPage);
  assert.equal(printed.records[0].emailAttachments[0].disposition, "print-with-email");
  assert.equal(omitted.records[0].emailAttachments[0].disposition, "leave-out");
  assert.equal(printed.records[0].emailAttachments[0].sha256, child.sha256);
  const { formatBuildReportText, createBuildReportPayload } = await import("../app/lib/build-report.ts");
  const text = formatBuildReportText(createBuildReportPayload({
    projectName: "Email matter",
    build: printed,
    candidates: [{ ...base, emailAttachmentDispositions: { [child.identity]: "print-with-email" } }],
    analysis,
    preflight: [],
    resolutions: [],
  }));
  assert.match(text, /invoice\.pdf/);
  assert.match(text, /print-with-email/);
});

test("add-as-exhibit creates an uncited index row with parent provenance and no cited statement suggestion", async () => {
  const { analyseFiles, attachDerivedEmailEvidence, buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const txt = new TextEncoder().encode("Separate schedule body");
  const raw = makeEml({ parts: [attachmentPart({ name: "schedule.txt", mimeType: "text/plain", bytes: txt })] });
  const analysis = await analyseFiles(await statementFile(), [new File([raw], "Notice.eml")]);
  const child = analysis.evidence[0].emailAttachments[0];
  const parent = { ...analysis.candidates[0], evidenceId: analysis.evidence[0].id, included: true, confirmed: true, emailAttachmentDispositions: { [child.identity]: "add-as-exhibit" } };
  const hydrated = await attachDerivedEmailEvidence(analysis, [parent, {
    id: "manual-schedule",
    mark: "EX 1",
    provisionalNumber: 2,
    description: "Attached schedule",
    date: "Date not stated",
    paragraph: 0,
    citation: "",
    citationResolution: "none",
    discoverySignals: ["Manually added by reviewer"],
    evidenceId: "stale",
    confidence: 100,
    rationale: "Added attachment",
    included: true,
    confirmed: true,
    exhibitInitials: "EX",
    exhibitSequence: 1,
    witnessInitials: "EX",
    sequenceOrder: 2000,
    manualAddition: true,
    parentEmailProvenance: {
      parentName: "Notice.eml",
      parentSha256: analysis.evidence[0].sha256,
      childIdentity: child.identity,
      childSha256: child.sha256,
    },
  }]);
  const build = await buildBundle(hydrated.analysis, hydrated.candidates);
  assert.equal(build.records.length, 2);
  const added = build.records.find((record) => record.citationStatus === "not-cited-manual-addition");
  assert.equal(added.description, "Attached schedule");
  assert.equal(added.statementParagraph, null);
  assert.deepEqual(added.statementReferences, []);
  const suggestions = buildStatementUpdateSuggestions(build.records);
  const heading = suggestions.findIndex((item) => item.line === "Uncited exhibits — no statement reference");
  assert.ok(heading >= 0, "added attachments appear in the uncited suggestion section");
  assert.equal(suggestions.slice(0, heading).some((item) => item.line.includes("Attached schedule")), false);
  assert.ok(suggestions.slice(heading + 1).some((item) => item.line.includes("Attached schedule")));
  assert.ok(suggestions.slice(heading + 1).every((item) => !item.line.startsWith("Paragraph ") && !/\[AH\d|\[LV\d/i.test(item.line)));
  assert.equal(parent.emailAttachmentDispositions[child.identity], "add-as-exhibit");
  assert.equal(hydrated.candidates.find((candidate) => candidate.id === "manual-schedule")?.confirmed, true, "successful rederive keeps the journal confirmation");
});

test("review display keeps email-child exhibits out of statement card order", async () => {
  const { reviewCandidatesForDisplay } = await import("../app/lib/exhibit-groups.ts");
  const parent = { id: "email", included: true, evidenceId: "e1", paragraph: 44 };
  const other = { id: "other", included: true, evidenceId: "e3", paragraph: 6 };
  const child = {
    id: "child",
    included: true,
    evidenceId: "e2",
    paragraph: 0,
    parentEmailProvenance: { parentName: "Notice.eml", parentSha256: "aa", childIdentity: "id", childSha256: "bb" },
  };
  const displayed = reviewCandidatesForDisplay([parent, other, child], [
    { id: "candidate-email", canonical: parent },
    { id: "candidate-other", canonical: other },
    { id: "candidate-child", canonical: child },
  ]);
  assert.deepEqual(displayed.map((item) => item.id), ["email", "other"]);
});

test("unconfirmed email children pending on the parent card, not a hidden child id", async () => {
  const { deriveExhibitGroups, pendingReviewCandidateIds, reviewCandidatesForDisplay } = await import("../app/lib/exhibit-groups.ts");
  const { analyseFiles } = await import("../app/lib/bundle-engine.ts");
  const pdf = await pdfBytes("CHILD-BODY");
  const raw = makeEml({ parts: [attachmentPart({ name: "invoice.pdf", mimeType: "application/pdf", bytes: pdf })] });
  const analysis = await analyseFiles(await statementFile(), [new File([raw], "Notice.eml")]);
  const parentEvidence = analysis.evidence[0];
  const child = parentEvidence.emailAttachments[0];
  const parent = { ...analysis.candidates[0], id: "email-card", evidenceId: parentEvidence.id, included: true, confirmed: true };
  const childCandidate = {
    ...parent,
    id: "hidden-child",
    evidenceId: "derived-child",
    confirmed: false,
    manualAddition: true,
    parentEmailProvenance: {
      parentName: parentEvidence.name,
      parentSha256: parentEvidence.sha256,
      childIdentity: child.identity,
      childSha256: child.sha256,
    },
  };
  const groups = deriveExhibitGroups(analysis, [parent, childCandidate]);
  const pending = pendingReviewCandidateIds([parent, childCandidate], groups, analysis.evidence);
  assert.equal(pending.has("hidden-child"), false);
  assert.equal(pending.has("email-card"), true);
  const displayed = reviewCandidatesForDisplay([parent, childCandidate], groups).filter((candidate) => pending.has(candidate.id));
  assert.deepEqual(displayed.map((item) => item.id), ["email-card"]);
});

test("unsupported attachments can only be left out, and xlsx children require native Excel printing", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { runPreflight } = await import("../app/lib/preflight.ts");
  const workbook = await syntheticWorkbook();
  const raw = makeEml({
    parts: [
      attachmentPart({ name: "photo.png", mimeType: "image/png", bytes: new Uint8Array([137, 80, 78, 71]) }),
      attachmentPart({ name: "ledger.xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", bytes: workbook }),
    ],
  });
  const analysis = await analyseFiles(await statementFile(), [new File([raw], "Notice.eml")]);
  const photo = analysis.evidence[0].emailAttachments.find((child) => child.extension === "png");
  const sheet = analysis.evidence[0].emailAttachments.find((child) => child.extension === "xlsx");
  assert.equal(photo.supported, false);
  assert.ok(sheet.workbook);
  sheet.sheetSelections = (sheet.sheetSelections ?? []).map((selection) => ({ ...selection, included: false }));
  const candidate = {
    ...analysis.candidates[0],
    evidenceId: analysis.evidence[0].id,
    included: true,
    confirmed: true,
    emailAttachmentDispositions: { [photo.identity]: "print-with-email", [sheet.identity]: "print-with-email" },
  };
  const checks = runPreflight(analysis, [candidate], BUNDLE_PROFILES[0]);
  assert.ok(checks.some((check) => check.code === "email.attachment_unsupported"));
  assert.ok(checks.some((check) => check.label === "No worksheet selected" && check.fileName === "ledger.xlsx"));
  candidate.emailAttachmentDispositions = { [photo.identity]: "leave-out", [sheet.identity]: "print-with-email" };
  sheet.sheetSelections = sheet.sheetSelections.map((selection) => ({ ...selection, included: true }));
  await assert.rejects(buildBundle(analysis, [candidate]), /Microsoft Excel/);
  const built = await buildBundle(analysis, [candidate], { workbookExporter: workbookPrintTestDouble });
  assert.ok(built.records[0].endPage > 2);
});

test("nested eml grandchildren stay listed and are not silently printed with the nested message", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const innerPdf = await pdfBytes("GRANDCHILD");
  const nested = makeEml({
    subject: "Forwarded",
    parts: [attachmentPart({ name: "grandchild.pdf", mimeType: "application/pdf", bytes: innerPdf })],
  });
  const raw = makeEml({
    parts: [attachmentPart({ name: "forwarded.eml", mimeType: "message/rfc822", bytes: Buffer.from(nested, "utf8") })],
  });
  const analysis = await analyseFiles(await statementFile(), [new File([raw], "Notice.eml")]);
  const nestedEml = analysis.evidence[0].emailAttachments.find((child) => child.extension === "eml");
  const grandchild = analysis.evidence[0].emailAttachments.find((child) => child.name === "grandchild.pdf");
  assert.ok(nestedEml && grandchild);
  const candidate = {
    ...analysis.candidates[0],
    evidenceId: analysis.evidence[0].id,
    included: true,
    confirmed: true,
    emailAttachmentDispositions: { [nestedEml.identity]: "print-with-email", [grandchild.identity]: "leave-out" },
  };
  const built = await buildBundle(analysis, [candidate]);
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const document = await pdfjs.getDocument({ data: new Uint8Array(built.bytes) }).promise;
  let combined = "";
  for (let page = 1; page <= document.numPages; page += 1) {
    const text = await (await document.getPage(page)).getTextContent();
    combined += text.items.map((item) => ("str" in item ? item.str : "")).join(" ");
  }
  await document.destroy();
  assert.doesNotMatch(combined, /GRANDCHILD/);
});

test("extracted children count toward the 500-evidence cap", async () => {
  const { analyseEvidenceFiles } = await import("../app/lib/bundle-engine.ts");
  const { EMAIL_CHILD_LIMITS } = await import("../app/lib/email-attachments.ts");
  assert.equal(EMAIL_CHILD_LIMITS.children, 500);
  const pdf = await pdfBytes("ONE");
  const parts = Array.from({ length: 3 }, (_, index) => attachmentPart({ name: `file-${index}.pdf`, mimeType: "application/pdf", bytes: pdf }));
  const raw = makeEml({ parts });
  const records = await analyseEvidenceFiles([new File([raw], "Notice.eml")]);
  assert.equal(records[0].emailAttachments.length, 3);
});
