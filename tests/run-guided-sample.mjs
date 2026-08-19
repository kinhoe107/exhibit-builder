import assert from "node:assert/strict";
import { createHash, webcrypto } from "node:crypto";
import { File } from "node:buffer";
import { basename } from "node:path";
import { tmpdir } from "node:os";
import { createRequire } from "node:module";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import test from "node:test";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

globalThis.File = File;
globalThis.crypto ??= webcrypto;
const require = createRequire(import.meta.url);
const { exportWorkbookSheets } = require("../electron/workbook-export.cjs");

const sampleRoot = new URL("../public/guided-sample/", import.meta.url);
const stagedStatement = new URL("../tmp/guided-sample-staged/01_GUIDED_SAMPLE_Witness_Statement.docx", import.meta.url);
const outputRoot = new URL("../output/pdf/", import.meta.url);

async function fixture(name, overrideUrl) {
  const url = overrideUrl ?? new URL(name, sampleRoot);
  return new File([await readFile(url)], name);
}

function sha(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function reviewedPdfTemplate(slot, file, bytes) {
  const pdfSha256 = sha(bytes);
  const confirmation = { pdfSha256, confirmedAt: "2026-08-13T10:00:00.000Z" };
  return { slot, file, sha256: pdfSha256, sourceFormat: "pdf", pdfFile: file, pdfSha256, reviewState: { matterReview: { sourceName: file.name, pdfSha256, exactByteLength: bytes.byteLength, pageCount: 1, extractedCharacterCount: 0, textReliability: "none", requiresVisualConfirmation: true, notice: "Guided fixture visually reviewed.", matterNumbers: [], partyNames: [], forums: [], matterTitles: [], placeholders: [] }, matterConfirmation: confirmation } };
}

async function colourWorkbookPreview(_file, sheets) {
  return Promise.all(sheets.map(async (sheet) => {
    const pdf = await PDFDocument.create();
    const regular = await pdf.embedFont(StandardFonts.Helvetica);
    const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
    const page = pdf.addPage([841.89, 595.28]);
    page.drawRectangle({ x: 45, y: 520, width: 752, height: 30, color: rgb(0.95, 0.86, 0.81) });
    page.drawText("SAMPLE DOCUMENT - FOR EXHIBIT BUILDER DEMONSTRATION ONLY", { x: 210, y: 531, size: 9, font: bold, color: rgb(0.48, 0.21, 0.12) });
    page.drawRectangle({ x: 45, y: 460, width: 752, height: 38, color: rgb(0.09, 0.21, 0.36) });
    page.drawText(`SAMPLE COST WORKBOOK - ${sheet.name}`, { x: 58, y: 474, size: 16, font: bold, color: rgb(1, 1, 1) });
    page.drawText(`Selected cell range ${sheet.range}. This full-colour test page represents the native Excel print output.`, { x: 58, y: 432, size: 11, font: regular, color: rgb(0.08, 0.14, 0.22) });
    return { name: sheet.name, range: sheet.range, bytes: await pdf.save() };
  }));
}

async function nativeWorkbookExporter(file, sheets) {
  return exportWorkbookSheets(tmpdir(), { fileName: file.name, bytes: new Uint8Array(await file.arrayBuffer()), sheets });
}

test("builds the complete guided sample with mixed evidence, one selected worksheet and supplied templates", async () => {
  const { analyseBundleStatements, buildBundle, SAMPLE_EVIDENCE, SAMPLE_STATEMENT, SAMPLE_TEMPLATES } = await import("../app/lib/bundle-engine.ts");
  const statementUrl = process.env.EXHIBIT_GUIDED_STAGED_STATEMENT === "1" ? stagedStatement : new URL(SAMPLE_STATEMENT, sampleRoot);
  const statement = await fixture(SAMPLE_STATEMENT, statementUrl);
  const evidenceFiles = await Promise.all(SAMPLE_EVIDENCE.map((name) => fixture(name)));
  const workbookSourceBefore = process.env.EXHIBIT_BUILDER_NATIVE_EXCEL === "1" ? await readFile(new URL("05_SAMPLE_Cost_Workbook.xlsx", sampleRoot)) : null;
  const analysed = await analyseBundleStatements([{ id: "guided-sample", file: statement, witnessName: "Guided Sample", witnessInitials: "AH" }], evidenceFiles);
  const evidence = analysed.evidence.map((record) => record.name === "05_SAMPLE_Cost_Workbook.xlsx"
    ? { ...record, sheetSelections: record.sheetSelections.map((selection, index) => ({ ...selection, included: index === 0 })) }
    : record);
  const guidedFiles = [
    "01_SAMPLE_Agreement.pdf",
    "02_SAMPLE_Invoice.pdf",
    "03_SAMPLE_Project_Report.docx",
    "04_SAMPLE_Claimant_Email.eml",
    "05_SAMPLE_Cost_Workbook.xlsx",
    "01_SAMPLE_Agreement.pdf",
  ];
  const candidates = analysed.candidates.map((candidate, index) => {
    const record = evidence.find((item) => item.name === guidedFiles[index]);
    assert.ok(record, `missing guided mapping for reference ${index + 1}`);
    return {
      ...candidate,
      evidenceId: record.id,
      confidence: 100,
      included: true,
      confirmed: true,
      repeatDecision: index === 5 ? "same" : candidate.repeatDecision,
      ...(record.emailAttachments?.length ? {
        emailAttachmentDispositions: Object.fromEntries(record.emailAttachments.map((child) => [child.identity, "print-with-email"])),
      } : {}),
    };
  });
  const templates = await Promise.all(SAMPLE_TEMPLATES.map(async ({ slot, name }) => {
    const bytes = await readFile(new URL(name, sampleRoot));
    const file = new File([bytes], name, { type: "application/pdf" });
    return reviewedPdfTemplate(slot, file, bytes);
  }));
  const workbookExporter = process.env.EXHIBIT_BUILDER_NATIVE_EXCEL === "1" ? nativeWorkbookExporter : colourWorkbookPreview;
  const build = await buildBundle({ ...analysed, evidence, candidates }, candidates, {
    templates,
    profileId: "exhibit-neutral",
    workbookExporter,
  });
  assert.equal(build.records.length, 5);
  assert.equal(build.records.filter((record) => record.workbookSheet).length, 1);
  assert.equal(build.records.find((record) => record.workbookSheet)?.workbookSheet.name, "Summary - Include");
  assert.equal(build.records.find((record) => record.workbookSheet)?.workbookSheet.range, "A1:H15");
  assert.equal(build.records[0].statementReferences.length, 2);
  assert.equal(build.manifest.templates.length, 2);
  assert.ok(build.pageCount >= 7);
  await mkdir(outputRoot, { recursive: true });
  await writeFile(new URL("Guided_Sample_Bundle_QA.pdf", outputRoot), build.bytes);
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const rendered = await pdfjs.getDocument({ data: new Uint8Array(build.bytes) }).promise;
  const coverText = (await (await rendered.getPage(1)).getTextContent()).items.map((item) => item.str).join(" ");
  const indexText = (await (await rendered.getPage(2)).getTextContent()).items.map((item) => item.str).join(" ");
  assert.match(coverText, /GUIDED SAMPLE - PDF COVER TEMPLATE/);
  assert.match(indexText, /ITEM NO\./);
  for (const description of ["Sample agreement", "Sample invoice", "Sample project report", "Sample claimant email", "Sample cost workbook"]) assert.match(indexText, new RegExp(description, "i"));
  if (process.env.EXHIBIT_BUILDER_NATIVE_EXCEL === "1") {
    assert.equal(sha(await readFile(new URL("05_SAMPLE_Cost_Workbook.xlsx", sampleRoot))), sha(workbookSourceBefore), "native Excel printing leaves the source workbook byte-for-byte unchanged");
    const workbookRecord = build.records.find((record) => record.workbookSheet);
    assert.ok(workbookRecord, "native fidelity check requires the selected workbook exhibit");
    const workbookPage = await rendered.getPage(workbookRecord.startPage);
    const workbookText = (await workbookPage.getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ").replace(/\s+/g, " ");
    for (const expected of ["SAMPLE COST WORKBOOK", "C-101", "Workbook verification", "Choose worksheet tabs at the Sheets stage.", "What sheet selection means", "A workbook may contain several worksheet tabs"]) {
      assert.match(workbookText, new RegExp(expected.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i"), `native Excel output retains ${expected}`);
    }
    const operators = await workbookPage.getOperatorList();
    const colourOperations = new Set([pdfjs.OPS.setFillRGBColor, pdfjs.OPS.setStrokeRGBColor]);
    const nonGreyColour = operators.fnArray.some((operation, index) => {
      if (!colourOperations.has(operation)) return false;
      const raw = Array.from(operators.argsArray[index] ?? []);
      const hex = raw.find((value) => typeof value === "string" && /^#[0-9a-f]{6}$/i.test(value));
      const flattened = hex
        ? [Number.parseInt(hex.slice(1, 3), 16), Number.parseInt(hex.slice(3, 5), 16), Number.parseInt(hex.slice(5, 7), 16)]
        : raw.flatMap((value) => ArrayBuffer.isView(value) || Array.isArray(value) ? Array.from(value) : [value]).filter((value) => typeof value === "number");
      return flattened.length >= 3 && (Math.abs(flattened[0] - flattened[1]) > 0.005 || Math.abs(flattened[1] - flattened[2]) > 0.005);
    });
    assert.equal(nonGreyColour, true, "native Excel output retains at least one non-grey source colour");
  }
  await rendered.destroy();
});
