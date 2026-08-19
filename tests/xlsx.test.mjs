import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { basename } from "node:path";
import { File } from "node:buffer";
import test from "node:test";
import JSZip from "jszip";
globalThis.File = File;

async function syntheticWorkbook({ name = "Ledger", dimension = "A1:H80", printArea, titles, rows = "", merges = "", cols = "" } = {}) {
  const zip = new JSZip();
  zip.file("xl/workbook.xml", `<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${name}" sheetId="1" r:id="rId1"/></sheets><definedNames>${printArea ? `<definedName name="_xlnm.Print_Area" localSheetId="0">${printArea}</definedName>` : ""}${titles ? `<definedName name="_xlnm.Print_Titles" localSheetId="0">${titles}</definedName>` : ""}</definedNames></workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file("xl/worksheets/sheet1.xml", `<worksheet><dimension ref="${dimension}"/><cols>${cols}</cols><sheetData>${rows}</sheetData>${merges}</worksheet>`);
  return new File([await zip.generateAsync({ type: "uint8array" })], "synthetic.xlsx");
}

test("reads the supplied OOXML workbook without executing formulas", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const path = new URL("./fixtures/moorland/Exhibits/Payment_Ledger_to_2026-03-31.xlsx", import.meta.url);
  const file = new File([await readFile(path)], basename(path.pathname));
  const analysis = await analyseXlsx(file);
  assert.ok(analysis.sheets.length > 0);
  assert.ok(analysis.sheets.some((sheet) => sheet.state === "visible"));
  assert.ok(analysis.sheets.some((sheet) => sheet.cells.length > 0));
  for (const sheet of analysis.sheets) {
    assert.equal(sheet.renderPlan.sourceHash, analysis.sourceHash);
    assert.equal(sheet.renderPlan.path, sheet.path);
    assert.ok(sheet.renderPlan.predictedPageCount >= 1);
    assert.equal(sheet.renderPlan.fontSize, 7);
    assert.match(sheet.renderPlan.range, /^[A-Z]+\d+:[A-Z]+\d+$/);
  }
  assert.match(analysis.warnings[0], /cached displayed values/i);
});

test("rejects an unsupported extension before parsing", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  await assert.rejects(analyseXlsx(new File(["x"], "unsafe.xlsm")), /Only .xlsx/i);
});

test("uses a qualified contiguous Print_Area and rejects discontiguous areas", async () => {
  const { printAreaForSheet } = await import("../app/lib/xlsx.ts");
  assert.deepEqual(printAreaForSheet("'Ledger'!$B$2:$D$9", "Ledger").range, { left: 2, top: 2, right: 4, bottom: 9 });
  assert.equal(printAreaForSheet("'Ledger'!$B$2:$D$9", "Ledger").warning, undefined, "a valid Print_Area must not produce an invalid-area warning");
  assert.deepEqual(printAreaForSheet("'O''Brien'!$A$1:$B$2", "O'Brien").range, { left: 1, top: 1, right: 2, bottom: 2 });
  assert.match(printAreaForSheet("'Ledger'!$A$1:$B$2,'Ledger'!$D$1:$E$2", "Ledger").warning, /Discontiguous/i);
});

test("qualified print area wins over a larger worksheet dimension", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({ printArea: "'Ledger'!$B$2:$D$4", rows: `<row r="2"><c r="B2" t="inlineStr"><is><t>in range</t></is></c></row><row r="20"><c r="H20" t="inlineStr"><is><t>outside</t></is></c></row>` });
  const sheet = (await analyseXlsx(file)).sheets[0];
  assert.equal(sheet.range, "B2:D4");
  assert.deepEqual(sheet.renderPlan.bounds, { left: 2, top: 2, right: 4, bottom: 4 });
  assert.equal(sheet.renderPlan.tiles.length, 1);
});

test("physical plan paginates tall rows and retains print-title rows", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const rows = Array.from({ length: 24 }, (_, index) => { const r = index + 1; return `<row r="${r}" ht="42"><c r="A${r}" t="inlineStr"><is><t>${r === 1 ? "Title" : `wrapped evidence ${"x".repeat(140)}`}</t></is></c></row>`; }).join("");
  const file = await syntheticWorkbook({ dimension: "A1:A24", printArea: "'Ledger'!$A$1:$A$24", titles: "'Ledger'!$1:$1", rows });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.ok(plan.predictedPageCount > 1, "tall physical rows require multiple pages");
  assert.equal(plan.titleRows?.top, 1);
  assert.ok(plan.tiles.slice(1).every((tile) => tile.top > 1), "body tiles advance rather than overlap titles");
  assert.ok(plan.tiles.every((tile, index) => index === 0 || tile.top > plan.tiles[index - 1].bottom), "physical tiles do not overlap");
});

test("merge crossing a column tile boundary is retained for renderer repetition", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({ dimension: "A1:K2", printArea: "'Ledger'!$A$1:$K$2", rows: `<row r="1"><c r="I1" t="inlineStr"><is><t>Cross-tile merged evidence</t></is></c></row>`, merges: `<mergeCells count="1"><mergeCell ref="I1:J1"/></mergeCells>` });
  const sheet = (await analyseXlsx(file)).sheets[0];
  assert.deepEqual(sheet.merges[0], { left: 9, top: 1, right: 10, bottom: 1 });
  assert.ok(sheet.renderPlan.tiles.some((tile) => tile.left === 1 && tile.right === 9));
  assert.ok(sheet.renderPlan.tiles.some((tile) => tile.left === 10 && tile.right >= 10));
  assert.equal(sheet.cells.find((cell) => cell.row === 1 && cell.col === 9)?.value, "Cross-tile merged evidence");
});

test("automatic range expands to the complete bounds of a populated merged cell", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:H13",
    rows: `<row r="1"><c r="A1" t="inlineStr"><is><t>Title</t></is></c></row><row r="13"><c r="A13" t="inlineStr"><is><t>Complete guidance text</t></is></c></row>`,
    merges: `<mergeCells count="1"><mergeCell ref="A13:H15"/></mergeCells>`,
  });
  const sheet = (await analyseXlsx(file)).sheets[0];
  assert.equal(sheet.printArea, undefined);
  assert.equal(sheet.range, "A1:H15");
  assert.deepEqual(sheet.renderPlan.bounds, { left: 1, top: 1, right: 8, bottom: 15 });
});

test("guided workbook automatic range includes the complete explanatory note", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const path = new URL("../public/guided-sample/05_SAMPLE_Cost_Workbook.xlsx", import.meta.url);
  const file = new File([await readFile(path)], basename(path.pathname));
  const sheet = (await analyseXlsx(file)).sheets.find((item) => item.name === "Summary - Include");
  assert.ok(sheet);
  assert.equal(sheet.printArea, undefined);
  assert.equal(sheet.range, "A1:H15");
  assert.ok(sheet.merges.some((merge) => merge.left === 1 && merge.top === 13 && merge.right === 8 && merge.bottom === 15));
  assert.match(sheet.cells.find((cell) => cell.row === 13 && cell.col === 1)?.value ?? "", /A workbook may contain several worksheet tabs/);
});

test("a source print area that cuts through a merged cell creates a blocking fidelity issue", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:H15",
    printArea: "'Ledger'!$A$1:$H$13",
    rows: `<row r="13"><c r="A13" t="inlineStr"><is><t>Clipped note</t></is></c></row>`,
    merges: `<mergeCells count="1"><mergeCell ref="A13:H15"/></mergeCells>`,
  });
  const sheet = (await analyseXlsx(file)).sheets[0];
  assert.equal(sheet.range, "A1:H13");
  assert.ok(sheet.warnings.some((warning) => /Fidelity check failed/i.test(warning)));
});

test("workbook analysis still times out after 20 seconds when Worker is unavailable", async () => {
  const source = await readFile(new URL("../app/lib/xlsx.ts", import.meta.url), "utf8");
  const fallback = source.slice(source.indexOf("export function analyseXlsxInWorker"), source.indexOf("return new Promise((resolve, reject) => {"));
  assert.match(fallback, /typeof Worker === "undefined"/);
  assert.match(fallback, /Promise\.race/);
  assert.match(fallback, /20_000/);
});
