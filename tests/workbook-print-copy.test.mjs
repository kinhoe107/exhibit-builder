import assert from "node:assert/strict";
import test from "node:test";

const SCALE = "Fidelity check failed: unbreakable worksheet content needs 77% scale to fit one landscape A4 printable page; the readable minimum is 85%.";
const MERGED_WIDE = "Fidelity check failed: merged cell A1:H1 is wider than one landscape A4 printable page at the readable 85% minimum scale.";
const MERGED_TALL = "Fidelity check failed: merged cell A1:A40 is taller than one portrait A4 printable page at the readable 85% minimum scale.";
const COLUMN_WIDE = "Fidelity check failed: column G is wider than one landscape A4 printable page at the readable 85% minimum scale.";
const DRAWINGS = "Fidelity check failed: worksheet drawings or charts are not supported because their printable anchors cannot yet be verified inside the approved range.";
const NOTE = "Print titles were ignored for this preview.";

test("combines a scale failure and a wide merge into one too-wide card", async () => {
  const { workbookPlanCheckCopy, WORKBOOK_TOO_WIDE_LABEL } = await import("../app/lib/workbook-print-copy.ts");
  const checks = workbookPlanCheckCopy("Programme", [SCALE, MERGED_WIDE]);
  assert.equal(checks.length, 1);
  assert.equal(checks[0].blocking, true);
  assert.equal(checks[0].idSuffix, "print-fit");
  assert.equal(checks[0].label, WORKBOOK_TOO_WIDE_LABEL);
  assert.doesNotMatch(checks[0].detail, /Fidelity check failed/i);
  assert.doesNotMatch(checks[0].detail, /Baseline_Programme/i);
  assert.match(checks[0].detail, /Programme is too wide for landscape A4/);
  assert.match(checks[0].detail, /A1:H1 would need the sheet printed at 77%/);
  assert.match(checks[0].detail, /stops shrinking at 85%/);
  assert.match(checks[0].detail, /split A1:H1/i);
  assert.match(checks[0].detail, /If this Excel file should not be in the bundle, leave it out/);
  assert.doesNotMatch(checks[0].detail, /workbook/i);
});

test("keeps a tall merge on its own card", async () => {
  const { workbookPlanCheckCopy, WORKBOOK_TOO_TALL_LABEL } = await import("../app/lib/workbook-print-copy.ts");
  const checks = workbookPlanCheckCopy("Programme", [SCALE, MERGED_WIDE, MERGED_TALL]);
  assert.equal(checks.length, 2);
  assert.equal(checks[0].label, "This Excel sheet is too wide to print on A4");
  assert.equal(checks[1].label, WORKBOOK_TOO_TALL_LABEL);
  assert.match(checks[1].detail, /A1:A40/);
  assert.doesNotMatch(checks.map((check) => check.detail).join("\n"), /Fidelity check failed/i);
});

test("explains other fidelity failures without the internal prefix", async () => {
  const { workbookPlanCheckCopy, WORKBOOK_CANNOT_PRINT_LABEL } = await import("../app/lib/workbook-print-copy.ts");
  const checks = workbookPlanCheckCopy("Chart evidence", [DRAWINGS, NOTE]);
  assert.equal(checks[0].label, WORKBOOK_CANNOT_PRINT_LABEL);
  assert.match(checks[0].detail, /drawings or charts whose print position cannot be checked/);
  assert.equal(checks[1].label, "Worksheet print note");
  assert.match(checks[1].detail, /Chart evidence: Print titles were ignored/);
  assert.doesNotMatch(checks.map((check) => `${check.label}\n${check.detail}`).join("\n"), /Fidelity check failed/i);
});

test("preflight emits one blocking check for the combined Programme warnings", async () => {
  const { File } = await import("node:buffer");
  const JSZip = (await import("jszip")).default;
  const { analyseFiles } = await import("../app/lib/bundle-engine.ts");
  const { runPreflight } = await import("../app/lib/preflight.ts");
  const { BUNDLE_PROFILES } = await import("../app/lib/bundle-types.ts");
  const { WORKBOOK_TOO_WIDE_LABEL } = await import("../app/lib/workbook-print-copy.ts");
  globalThis.File = File;
  const zip = new JSZip();
  zip.file("xl/workbook.xml", `<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Programme" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file("xl/worksheets/sheet1.xml", `<worksheet><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Programme</t></is></c></row></sheetData></worksheet>`);
  const workbook = new File([await zip.generateAsync({ type: "uint8array" })], "Baseline_Programme_RevC_2025-12-05.xlsx");
  const statement = new File([await (await import("node:fs/promises")).readFile(new URL("fixtures/core/01_Witness_Statement_Amelia_Hart.docx", import.meta.url))], "01_Witness_Statement_Amelia_Hart.docx");
  const analysis = await analyseFiles(statement, [workbook]);
  const evidence = analysis.evidence[0];
  evidence.sheetSelections = [{ name: "Programme", included: true }];
  evidence.workbook.sheets[0].renderPlan.warnings = [SCALE, MERGED_WIDE];
  const candidate = { ...analysis.candidates[0], evidenceId: evidence.id, included: true, confirmed: true };
  const fidelity = runPreflight(analysis, [candidate], BUNDLE_PROFILES[0]).filter((check) => check.code === "workbook.fidelity_failed");
  assert.equal(fidelity.length, 1);
  assert.equal(fidelity[0].label, WORKBOOK_TOO_WIDE_LABEL);
  assert.equal(fidelity[0].fileName, "Baseline_Programme_RevC_2025-12-05.xlsx");
  assert.doesNotMatch(fidelity[0].detail, /Baseline_Programme_RevC_2025-12-05\.xlsx/);
  assert.doesNotMatch(fidelity[0].detail, /Fidelity check failed/i);
  assert.match(fidelity[0].detail, /Programme is too wide for landscape A4/);
});

test("a wide column without a scale warning still uses the too-wide heading", async () => {
  const { workbookPlanCheckCopy, WORKBOOK_TOO_WIDE_LABEL } = await import("../app/lib/workbook-print-copy.ts");
  const checks = workbookPlanCheckCopy("Ledger", [COLUMN_WIDE]);
  assert.equal(checks.length, 1);
  assert.equal(checks[0].label, WORKBOOK_TOO_WIDE_LABEL);
  assert.match(checks[0].detail, /Column G on Ledger is wider/);
  assert.match(checks[0].detail, /leave this Excel file out of the bundle/);
  assert.doesNotMatch(checks[0].detail, /workbook/i);
});
