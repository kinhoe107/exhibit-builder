import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { basename } from "node:path";
import { File } from "node:buffer";
import { webcrypto } from "node:crypto";
import test from "node:test";
import {
  applyGuidedSampleMapping,
  analysisWithGuidedMapping,
  hasGuidedSampleEvidence,
  isGuidedSampleSelection,
  GUIDED_SAMPLE_MAPPING_RATIONALE,
} from "../app/lib/guided-sample.ts";
import { resolveTourStep, tourWorkspaceView } from "../app/lib/guided-tour.ts";
import { analyseBundleStatements, SAMPLE_EVIDENCE, SAMPLE_STATEMENT, samplePackRelativePath } from "../app/lib/bundle-engine.ts";

globalThis.File = File;
globalThis.crypto ??= webcrypto;

const idleTour = {
  active: true,
  openedFolder: false,
  statementSelected: false,
  evidenceSelected: false,
  analysisReady: false,
  bulkConfirmableCount: 0,
  repeatPending: false,
  attachmentPending: false,
  attachmentChoicesOpen: false,
  printWithEmailVisible: false,
  view: "home",
  includedWorkbook: false,
  hasBuild: false,
  downloaded: false,
  saved: false,
};

test("guided sample filenames are detected without rewriting sources", () => {
  assert.equal(isGuidedSampleSelection(SAMPLE_STATEMENT, SAMPLE_EVIDENCE), true);
  assert.equal(isGuidedSampleSelection(SAMPLE_STATEMENT, [
    "01_SAMPLE_Agreement.pdf",
    "02_SAMPLE_Invoice.pdf",
    "03_SAMPLE_Project_Report.docx",
    "04_SAMPLE_Claimant_Email.eml",
    "05_SAMPLE_Cost_Workbook.xlsx",
  ]), true);
  assert.equal(isGuidedSampleSelection("other.docx", SAMPLE_EVIDENCE), false);
  assert.equal(isGuidedSampleSelection(SAMPLE_STATEMENT, SAMPLE_EVIDENCE.slice(1)), false);
  assert.equal(hasGuidedSampleEvidence([...SAMPLE_EVIDENCE, "extra.pdf"]), true);
  assert.equal(analysisWithGuidedMapping({ candidates: [], evidence: [] }, "other.docx", SAMPLE_EVIDENCE).candidates.length, 0);
});

test("guided sample folder separates the witness statement from the exhibits", async () => {
  const { SAMPLE_EXHIBITS_FOLDER, SAMPLE_STATEMENT_FOLDER, SAMPLE_UNUSED } = await import("../app/lib/bundle-engine.ts");
  assert.equal(samplePackRelativePath(SAMPLE_STATEMENT), `${SAMPLE_STATEMENT_FOLDER}/${SAMPLE_STATEMENT}`);
  assert.equal(samplePackRelativePath("01_SAMPLE_Agreement.pdf"), `${SAMPLE_EXHIBITS_FOLDER}/01_SAMPLE_Agreement.pdf`);
  assert.equal(samplePackRelativePath(SAMPLE_UNUSED), SAMPLE_UNUSED);
  assert.equal(samplePackRelativePath("00_GUIDED_SAMPLE_Cover_Template.pdf"), "00_GUIDED_SAMPLE_Cover_Template.pdf");
  const root = new URL("../public/guided-sample/", import.meta.url);
  const [statement, agreement, unused, cover] = await Promise.all([
    readFile(new URL(samplePackRelativePath(SAMPLE_STATEMENT), root)),
    readFile(new URL(samplePackRelativePath("01_SAMPLE_Agreement.pdf"), root)),
    readFile(new URL(samplePackRelativePath(SAMPLE_UNUSED), root)),
    readFile(new URL(samplePackRelativePath("00_GUIDED_SAMPLE_Cover_Template.pdf"), root)),
  ]);
  assert.ok(statement.byteLength > 0);
  assert.ok(agreement.byteLength > 0);
  assert.ok(unused.byteLength > 0);
  assert.ok(cover.byteLength > 0);
  const tourCopy = await readFile(new URL("../app/lib/guided-tour.ts", import.meta.url), "utf8");
  assert.match(tourCopy, /from Witness statement/);
  assert.match(tourCopy, /files in the Exhibits folder/);
});

test("Analyse files mapping matches useSamplePack for the guided filenames", async () => {
  const root = new URL("../public/guided-sample/", import.meta.url);
  const statement = new File([await readFile(new URL(samplePackRelativePath(SAMPLE_STATEMENT), root))], SAMPLE_STATEMENT);
  const evidence = await Promise.all(SAMPLE_EVIDENCE.map(async (name) => new File([await readFile(new URL(samplePackRelativePath(name), root))], name)));
  const raw = await analyseBundleStatements([{ id: "guided", file: statement, witnessName: "Guided Sample", witnessInitials: "AH" }], evidence);
  const mapped = applyGuidedSampleMapping(raw);
  assert.notEqual(mapped, raw);
  const named = analysisWithGuidedMapping(raw, SAMPLE_STATEMENT, evidence.map((file) => file.name));
  assert.deepEqual(named.candidates.map((candidate) => candidate.evidenceId), mapped.candidates.map((candidate) => candidate.evidenceId));
  const other = analysisWithGuidedMapping(raw, "Matter_Statement.docx", evidence.map((file) => file.name));
  assert.equal(other, raw);
  assert.notEqual(raw.evidence.find((record) => record.name === "05_SAMPLE_Cost_Workbook.xlsx")?.sheetSelections?.map((sheet) => sheet.included).join(","), mapped.evidence.find((record) => record.name === "05_SAMPLE_Cost_Workbook.xlsx")?.sheetSelections?.map((sheet) => sheet.included).join(","));
  assert.equal(raw.candidates.slice(2, 5).every((candidate) => candidate.evidenceId === null), true);
  assert.deepEqual(mapped.candidates.map((candidate) => raw.evidence.find((record) => record.id === candidate.evidenceId)?.name), [
    "01_SAMPLE_Agreement.pdf",
    "02_SAMPLE_Invoice.pdf",
    "03_SAMPLE_Project_Report.docx",
    "04_SAMPLE_Claimant_Email.eml",
    "05_SAMPLE_Cost_Workbook.xlsx",
    "01_SAMPLE_Agreement.pdf",
  ]);
  assert.ok(mapped.candidates.every((candidate) => candidate.confidence === 100));
  assert.ok(mapped.candidates.every((candidate) => candidate.rationale === GUIDED_SAMPLE_MAPPING_RATIONALE));
  assert.equal(mapped.candidates[5].repeatDecision, "same");
  const workbook = mapped.evidence.find((record) => record.name === "05_SAMPLE_Cost_Workbook.xlsx");
  assert.deepEqual(workbook?.sheetSelections?.map((sheet) => sheet.included), [true, false, false]);
  assert.equal(basename(statement.name), SAMPLE_STATEMENT);
});

test("walkthrough starts on home, advances with state, and skip leaves files unloaded", () => {
  assert.equal(resolveTourStep({ ...idleTour, active: false }), null);
  assert.equal(resolveTourStep(idleTour), "open-folder");
  assert.equal(resolveTourStep({ ...idleTour, openedFolder: true }), "choose-statement");
  assert.equal(resolveTourStep({ ...idleTour, statementSelected: true }), "choose-evidence");
  assert.equal(resolveTourStep({ ...idleTour, statementSelected: true, evidenceSelected: true }), "analyse");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "review",
    bulkConfirmableCount: 6,
    includedWorkbook: true,
  }), "confirm-all");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "review",
    repeatPending: true,
    attachmentPending: true,
    printWithEmailVisible: true,
    includedWorkbook: true,
  }), "repeat-decision");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "review",
    attachmentPending: true,
    includedWorkbook: true,
  }), "attachments");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "review",
    attachmentPending: true,
    printWithEmailVisible: true,
    includedWorkbook: true,
  }), "print-with-email");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "review",
    attachmentPending: true,
    attachmentChoicesOpen: true,
    includedWorkbook: true,
  }), "print-with-email");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "review",
    includedWorkbook: true,
  }), "continue-sheets");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "sheets",
    includedWorkbook: true,
  }), "continue-finalise");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "build",
  }), "build");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "build",
    hasBuild: true,
  }), "download");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "build",
    hasBuild: true,
    downloaded: true,
  }), "save");
  assert.equal(resolveTourStep({
    ...idleTour,
    analysisReady: true,
    view: "build",
    hasBuild: true,
    downloaded: true,
    saved: true,
  }), null);
  assert.equal(tourWorkspaceView(false, "review"), "home");
  assert.equal(tourWorkspaceView(true, "reconcile"), "other");
});

test("walkthrough keeps skip, Escape, and labels without trapping native file dialogs", async () => {
  const [tour, component, styles] = await Promise.all([
    readFile(new URL("../app/GuidedSampleTour.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/BundleBuilder.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
  ]);
  assert.match(tour, /Skip walkthrough/);
  assert.match(tour, /event\.key !== "Escape"/);
  assert.match(tour, /escapeIsForOtherUi\(event, dialogOpen\)/);
  assert.match(tour, /ignoreFileDialogEscape/);
  assert.match(tour, /role="status"/);
  assert.match(tour, /aria-live="polite"/);
  assert.match(tour, /aria-atomic="true"/);
  assert.match(tour, /skipRef\.current\?\.focus\(\{ preventScroll: true \}\)/);
  assert.doesNotMatch(tour, /querySelector<HTMLElement>\(`\[data-tour="\$\{stepId\}"\]`\)[\s\S]{0,80}\.focus\(/);
  assert.doesNotMatch(tour, /event\.preventDefault\(\)/);
  assert.doesNotMatch(tour, /role="region"/);
  assert.match(component, /aria-describedby="guided-sample-note"/);
  assert.match(component, /id="guided-sample-note"/);
  assert.match(component, /ref=\{statementInput\}[\s\S]{0,220}tabIndex=\{-1\}/);
  assert.match(component, /ref=\{evidenceInput\}[\s\S]{0,280}tabIndex=\{-1\}/);
  assert.match(styles, /\.guided-tour-layer \{[\s\S]{0,80}pointer-events:\s*none/);
  assert.match(styles, /\.guided-tour-card \{[\s\S]{0,220}pointer-events:\s*none/);
  assert.match(styles, /\.guided-tour-skip \{[\s\S]{0,280}pointer-events:\s*auto/);
  assert.doesNotMatch(styles, /\.tour-current-target \{[\s\S]{0,80}pointer-events:\s*none/);
});
