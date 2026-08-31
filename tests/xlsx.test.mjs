import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { basename } from "node:path";
import { File } from "node:buffer";
import test from "node:test";
import JSZip from "jszip";
globalThis.File = File;

async function syntheticWorkbook({ name = "Ledger", dimension = "A1:H80", printArea, titles, rows = "", merges = "", cols = "", sheetFormat = "", worksheetExtras = "", styles, sheetRelationships, extraEntries = {} } = {}) {
  const zip = new JSZip();
  zip.file("xl/workbook.xml", `<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${name}" sheetId="1" r:id="rId1"/></sheets><definedNames>${printArea ? `<definedName name="_xlnm.Print_Area" localSheetId="0">${printArea}</definedName>` : ""}${titles ? `<definedName name="_xlnm.Print_Titles" localSheetId="0">${titles}</definedName>` : ""}</definedNames></workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  zip.file("xl/worksheets/sheet1.xml", `<worksheet><sheetFormatPr ${sheetFormat}/><dimension ref="${dimension}"/><cols>${cols}</cols><sheetData>${rows}</sheetData>${merges}${worksheetExtras}</worksheet>`);
  if (styles) zip.file("xl/styles.xml", styles);
  if (sheetRelationships) zip.file("xl/worksheets/_rels/sheet1.xml.rels", sheetRelationships);
  for (const [path, content] of Object.entries(extraEntries)) zip.file(path, content);
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
    assert.ok(sheet.renderPlan.scalePercent >= 85 && sheet.renderPlan.scalePercent <= 100);
    assert.match(sheet.renderPlan.range, /^[A-Z]+\d+:[A-Z]+\d+$/);
  }
  assert.match(analysis.warnings[0], /cached displayed values/i);
});

test("rejects an unsupported extension before parsing", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  await assert.rejects(analyseXlsx(new File(["x"], "unsafe.xlsm")), /Only .xlsx/i);
});

test("rejects out-of-bounds structural worksheet geometry before planning loops", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  await assert.rejects(analyseXlsx(await syntheticWorkbook({ dimension: "A1:CC2" })), /dimension exceeds/i);
  await assert.rejects(analyseXlsx(await syntheticWorkbook({ dimension: "A1:A2", rows: `<row r="5001"/>` })), /row outside/i);
  await assert.rejects(analyseXlsx(await syntheticWorkbook({ dimension: "A1:A2", cols: `<col min="1" max="81" width="10"/>` })), /column definition outside/i);
  await assert.rejects(analyseXlsx(await syntheticWorkbook({ dimension: "A1:A2", printArea: "'Ledger'!$A$1:$CC$2" })), /print area exceeds/i);
  await assert.rejects(analyseXlsx(await syntheticWorkbook({ dimension: "A1:A2", merges: `<mergeCells><mergeCell ref="A1:CC1"/></mergeCells>` })), /merged range outside/i);
  const excessiveMerges = `<mergeCells>${Array.from({ length: 10_001 }, () => `<mergeCell ref="A1:A1"/>`).join("")}</mergeCells>`;
  await assert.rejects(analyseXlsx(await syntheticWorkbook({ dimension: "A1:A2", merges: excessiveMerges })), /10000-merge analysis limit/i);
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

test("models saved margins but plans and discloses canonical A4 margins", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:B2",
    rows: `<row r="1"><c r="A1" t="inlineStr"><is><t>Evidence</t></is></c></row>`,
    worksheetExtras: `<pageMargins left="1.25" right="0.9" top="1.1" bottom="1.2" header="0.4" footer="0.5"/>`,
  });
  const sheet = (await analyseXlsx(file)).sheets[0];
  assert.equal(sheet.pageMargins.left, 1.25, "analysis records the saved source geometry");
  assert.deepEqual(sheet.renderPlan.pageMargins, { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }, "the native plan uses disclosed canonical margins");
  assert.ok(sheet.warnings.some((warning) => /normalised.*canonical A4 margins/i.test(warning)));
});

test("blocks nonempty headers and footers before native Excel can clear them", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:B2",
    rows: `<row r="1"><c r="A1" t="inlineStr"><is><t>Evidence</t></is></c></row>`,
    worksheetExtras: `<headerFooter><oddHeader>&amp;LConfidential evidence</oddHeader><oddFooter>&amp;P</oddFooter></headerFooter>`,
  });
  assert.ok((await analyseXlsx(file)).sheets[0].warnings.some((warning) => /^Fidelity check failed:.*headers or footers/i.test(warning)));
});

test("blocks repeated title columns and retains only supported leading title rows", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const withColumns = await syntheticWorkbook({ dimension: "A1:H20", titles: "'Ledger'!$1:$2,'Ledger'!$A:$B" });
  const columnSheet = (await analyseXlsx(withColumns)).sheets[0];
  assert.equal(columnSheet.titleColumns, "$A:$B");
  assert.ok(columnSheet.warnings.some((warning) => /^Fidelity check failed: repeated print-title columns/i.test(warning)));

  const nonPrefix = await syntheticWorkbook({ dimension: "A1:H20", titles: "'Ledger'!$2:$3" });
  assert.ok((await analyseXlsx(nonPrefix)).sheets[0].warnings.some((warning) => /^Fidelity check failed: repeated print-title rows must be a leading prefix/i.test(warning)));
});

test("normalises over-then-down source order to deterministic down-then-over output", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({ dimension: "A1:T30", worksheetExtras: `<pageSetup pageOrder="overThenDown"/>` });
  const sheet = (await analyseXlsx(file)).sheets[0];
  assert.equal(sheet.pageOrder, "overThenDown", "the source setting remains visible in analysis");
  assert.equal(sheet.renderPlan.pageOrder, "downThenOver");
  assert.ok(sheet.warnings.some((warning) => /normalised.*down-then-over/i.test(warning)));
});

test("non-default Normal fonts and automatic wrapped rows defer to native geometry checks", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const styles = `<styleSheet><fonts count="1"><font><sz val="12"/><name val="Arial"/></font></fonts><cellStyleXfs count="1"><xf fontId="0"/></cellStyleXfs><cellXfs count="2"><xf/><xf><alignment wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0"/></cellStyles></styleSheet>`;
  const file = await syntheticWorkbook({ dimension: "A1:B3", rows: `<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Wrapped evidence that Excel sizes automatically</t></is></c></row>`, styles });
  const sheet = (await analyseXlsx(file)).sheets[0];
  assert.equal(sheet.cells[0].wrap, true);
  assert.equal(sheet.warnings.some((warning) => /^Fidelity check failed:/i.test(warning) && /Normal style|automatic row height/i.test(warning)), false, "OOXML estimates alone do not block native Excel");
  assert.ok(sheet.warnings.some((warning) => /Normal style uses Arial 12pt.*native Microsoft Excel dimensions/i.test(warning)));
  assert.ok(sheet.warnings.some((warning) => /wrapped text with automatic row heights.*native Microsoft Excel dimensions/i.test(warning)));
  assert.ok(sheet.renderPlan.geometryChecks.some((check) => check.axis === "vertical" && check.ranges.includes("A1:A1")), "native Excel receives an authoritative height check");
});

test("blocks printed headings and comment or note relationships", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:B2",
    worksheetExtras: `<printOptions headings="1"/><legacyDrawing r:id="note1"/>`,
    sheetRelationships: `<Relationships><Relationship Id="note1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>`,
  });
  const warnings = (await analyseXlsx(file)).sheets[0].warnings;
  assert.ok(warnings.some((warning) => /Fidelity check failed: printed row or column headings/i.test(warning)));
  assert.ok(warnings.some((warning) => /Fidelity check failed: worksheet comments, notes/i.test(warning)));
});

test("blocks worksheet drawings whose anchors may sit outside the approved range", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:B2",
    rows: `<row r="1"><c r="A1" t="inlineStr"><is><t>Evidence</t></is></c></row>`,
    worksheetExtras: `<drawing r:id="drawing1"/>`,
    sheetRelationships: `<Relationships><Relationship Id="drawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`,
    extraEntries: { "xl/drawings/drawing1.xml": `<wsDr><twoCellAnchor><from><col>4</col><row>6</row></from><to><col>8</col><row>12</row></to><graphicFrame/></twoCellAnchor></wsDr>` },
  });
  const sheet = (await analyseXlsx(file)).sheets[0];
  assert.ok(sheet.warnings.some((warning) => /^Fidelity check failed: worksheet drawings or charts/i.test(warning)));
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

test("a merged cell is never split between planned A4 pages", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({ dimension: "A1:K2", printArea: "'Ledger'!$A$1:$K$2", rows: `<row r="1"><c r="I1" t="inlineStr"><is><t>Cross-tile merged evidence</t></is></c></row>`, merges: `<mergeCells count="1"><mergeCell ref="I1:J1"/></mergeCells>` });
  const sheet = (await analyseXlsx(file)).sheets[0];
  assert.deepEqual(sheet.merges[0], { left: 9, top: 1, right: 10, bottom: 1 });
  for (const tile of sheet.renderPlan.tiles.filter((tile) => tile.left <= 10 && tile.right >= 9)) {
    assert.ok(tile.left <= 9 && tile.right >= 10, "the complete I:J merge stays on the same planned page");
  }
  assert.equal(sheet.cells.find((cell) => cell.row === 1 && cell.col === 9)?.value, "Cross-tile merged evidence");
});

test("moves an A4 boundary before a small merge rather than blocking or splitting it", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:Q2",
    rows: `<row r="1"><c r="P1" t="inlineStr"><is><t>Two-column heading</t></is></c></row>`,
    merges: `<mergeCells count="1"><mergeCell ref="P1:Q1"/></mergeCells>`,
  });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.equal(plan.warnings.some((warning) => /^Fidelity check failed:/i.test(warning)), false);
  const touching = plan.tiles.filter((tile) => tile.left <= 17 && tile.right >= 16);
  assert.ok(touching.length > 0);
  assert.ok(touching.every((tile) => tile.left <= 16 && tile.right >= 17), "no planned page boundary splits P1:Q1");
});

test("moves a row boundary before a tall vertical merge and keeps A9:A13 on one tile", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const rows = Array.from({ length: 20 }, (_, index) => {
    const row = index + 1;
    return `<row r="${row}" ht="70"><c r="A${row}" t="inlineStr"><is><t>${row === 9 ? "Vertically merged evidence" : `Row ${row}`}</t></is></c></row>`;
  }).join("");
  const file = await syntheticWorkbook({
    dimension: "A1:A20",
    rows,
    merges: `<mergeCells count="1"><mergeCell ref="A9:A13"/></mergeCells>`,
  });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  const verticalTiles = plan.tiles.filter((tile) => tile.left === plan.bounds.left && tile.right === plan.bounds.right);
  const rowBreaks = verticalTiles.slice(0, -1).map((tile) => tile.bottom);
  assert.equal(rowBreaks.some((row) => row >= 9 && row < 13), false, "no row break falls strictly inside A9:A13");
  assert.ok(rowBreaks.includes(8), "the proposed row break moves before the merged range");
  const containing = verticalTiles.filter((tile) => tile.top <= 9 && tile.bottom >= 13);
  assert.equal(containing.length, 1, "one planned tile contains the complete A9:A13 merge");
  assert.equal(plan.warnings.some((warning) => /^Fidelity check failed:/i.test(warning)), false);
});

test("an over-tall vertical merge stays unsplit and creates a blocking fidelity warning", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const rows = [1, 2, 3, 4].map((row) => `<row r="${row}" ht="300"><c r="A${row}" t="inlineStr"><is><t>${row === 1 ? "Over-tall merged evidence" : `Row ${row}`}</t></is></c></row>`).join("");
  const file = await syntheticWorkbook({
    dimension: "A1:A4",
    rows,
    merges: `<mergeCells count="1"><mergeCell ref="A1:A3"/></mergeCells>`,
  });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.ok(plan.warnings.some((warning) => /Fidelity check failed: merged cell A1:A3 is taller/i.test(warning)));
  assert.equal(plan.tiles.some((tile) => tile.bottom >= 1 && tile.bottom < 3), false, "the blocking plan still never proposes a row break inside the merge");
  assert.ok(plan.tiles.some((tile) => tile.top <= 1 && tile.bottom >= 3), "one tile retains the complete over-tall merge for review");
});

test("chooses the highest native scale that fits a 700pt portrait merged range", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:A3",
    rows: `<row r="1" ht="350"><c r="A1" t="inlineStr"><is><t>700 point merge</t></is></c></row><row r="2" ht="350"/><row r="3" ht="20"/>`,
    merges: `<mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells>`,
  });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.equal(plan.orientation, "portrait");
  assert.equal(plan.scalePercent, 96, "675 / 700 selects the highest satisfying whole-percent Excel zoom");
  assert.equal(plan.warnings.some((warning) => /^Fidelity check failed:/i.test(warning)), false);
  assert.ok(plan.tiles.some((tile) => tile.top <= 1 && tile.bottom >= 2));
});

test("evaluates both orientations before choosing portrait for tall content with horizontal pagination", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:K3",
    rows: `<row r="1" ht="300"><c r="A1" t="inlineStr"><is><t>Tall merged evidence</t></is></c></row><row r="2" ht="300"/><row r="3" ht="20"><c r="K3" t="inlineStr"><is><t>Final column</t></is></c></row>`,
    merges: `<mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells>`,
  });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.equal(plan.orientation, "portrait", "landscape's shorter page cannot contain the 600pt merge at the readable floor");
  assert.equal(plan.scalePercent, 100);
  assert.ok(new Set(plan.tiles.map((tile) => `${tile.left}:${tile.right}`)).size > 1, "portrait retains legibility by paginating horizontally");
  assert.equal(plan.warnings.some((warning) => /^Fidelity check failed:/i.test(warning)), false);
});

test("a genuinely wide horizontal merge selects landscape when it is the compliant orientation", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:N2",
    rows: `<row r="1"><c r="A1" t="inlineStr"><is><t>Wide merged heading</t></is></c></row>`,
    merges: `<mergeCells count="1"><mergeCell ref="A1:N1"/></mergeCells>`,
  });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.equal(plan.orientation, "landscape");
  assert.ok(plan.scalePercent >= 85);
  assert.equal(plan.warnings.some((warning) => /^Fidelity check failed:/i.test(warning)), false);
  assert.ok(plan.tiles.every((tile) => tile.left <= 1 && tile.right >= 14));
});

test("repeated title rows participate in the one global native scale and block only below 85 percent", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const fitting = await syntheticWorkbook({
    dimension: "A1:A5",
    titles: "'Ledger'!$1:$1",
    rows: `<row r="1" ht="100"><c r="A1" t="inlineStr"><is><t>Repeated title</t></is></c></row><row r="2" ht="325"><c r="A2" t="inlineStr"><is><t>Body merge</t></is></c></row><row r="3" ht="325"/><row r="4" ht="20"/><row r="5" ht="20"/>`,
    merges: `<mergeCells count="1"><mergeCell ref="A2:A3"/></mergeCells>`,
  });
  const fittingPlan = (await analyseXlsx(fitting)).sheets[0].renderPlan;
  assert.equal(fittingPlan.scalePercent, 90, "100pt repeated titles plus a 650pt merge fit at the highest satisfying scale");
  assert.equal(fittingPlan.warnings.some((warning) => /^Fidelity check failed:/i.test(warning)), false);

  const belowFloor = await syntheticWorkbook({
    dimension: "A1:A5",
    titles: "'Ledger'!$1:$1",
    rows: `<row r="1" ht="100"><c r="A1" t="inlineStr"><is><t>Repeated title</t></is></c></row><row r="2" ht="350"><c r="A2" t="inlineStr"><is><t>Body merge</t></is></c></row><row r="3" ht="350"/><row r="4" ht="20"/><row r="5" ht="20"/>`,
    merges: `<mergeCells count="1"><mergeCell ref="A2:A3"/></mergeCells>`,
  });
  const belowFloorPlan = (await analyseXlsx(belowFloor)).sheets[0].renderPlan;
  assert.equal(belowFloorPlan.scalePercent, 85);
  assert.ok(belowFloorPlan.warnings.some((warning) => /^Fidelity check failed:/i.test(warning)), "100pt titles plus a 700pt merge need 84%, below the readable floor");
});

test("title rows that exceed a full page at 85 percent create a blocking fidelity warning", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const rows = Array.from({ length: 8 }, (_, index) => {
    const row = index + 1;
    return `<row r="${row}" ht="${row <= 5 ? 200 : 20}"><c r="A${row}" t="inlineStr"><is><t>Row ${row}</t></is></c></row>`;
  }).join("");
  const file = await syntheticWorkbook({ dimension: "A1:A8", titles: "'Ledger'!$1:$5", rows });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.equal(plan.scalePercent, 85);
  assert.ok(plan.warnings.some((warning) => /Fidelity check failed: repeated title rows/i.test(warning)));
});

test("blocks a merged cell that cannot fit on one landscape A4 page", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:Z2",
    rows: `<row r="1"><c r="A1" t="inlineStr"><is><t>Unsplit merged evidence</t></is></c></row>`,
    merges: `<mergeCells count="1"><mergeCell ref="A1:Z1"/></mergeCells>`,
  });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.ok(plan.warnings.some((warning) => /^Fidelity check failed:/i.test(warning)));
});

test("blocks a worksheet whose saved-scale column cannot fit on an A4 page", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:A2",
    rows: `<row r="1"><c r="A1" t="inlineStr"><is><t>Extremely wide evidence</t></is></c></row>`,
    cols: `<col min="1" max="1" width="200" customWidth="1"/>`,
  });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.equal(plan.scalePercent, 85);
  assert.ok(plan.warnings.some((warning) => /^Fidelity check failed: column A/i.test(warning)));
});

test("blocks rather than allocating a native plan above twenty thousand pages", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const rows = Array.from({ length: 5000 }, (_, index) => `<row r="${index + 1}" ht="409"/>`).join("");
  const file = await syntheticWorkbook({ dimension: "A1:CB5000", rows, cols: `<col min="1" max="80" width="200" customWidth="1"/>` });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.ok(plan.predictedPageCount > 20_000);
  assert.equal(plan.tiles.length, 1, "the blocked plan does not allocate hundreds of thousands of page objects");
  assert.ok(plan.warnings.some((warning) => /Fidelity check failed: worksheet pagination predicts .*above the 20000-page/i.test(warning)));
});

test("honours a worksheet-wide default column width in the A4 fidelity gate", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const file = await syntheticWorkbook({
    dimension: "A1:A2",
    rows: `<row r="1"><c r="A1" t="inlineStr"><is><t>Default-width evidence</t></is></c></row>`,
    sheetFormat: `defaultColWidth="200" defaultRowHeight="42"`,
  });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.ok(plan.warnings.some((warning) => /^Fidelity check failed: column A/i.test(warning)));
});

test("uses the worksheet default row height for its A4 page estimate", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  const rows = Array.from({ length: 12 }, (_, index) => `<row r="${index + 1}"><c r="A${index + 1}" t="inlineStr"><is><t>Row ${index + 1}</t></is></c></row>`).join("");
  const file = await syntheticWorkbook({ dimension: "A1:A12", rows, sheetFormat: `defaultColWidth="8.43" defaultRowHeight="70"` });
  const plan = (await analyseXlsx(file)).sheets[0].renderPlan;
  assert.ok(plan.predictedPageCount > 1);
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
  const path = new URL("../public/guided-sample/Exhibits/05_SAMPLE_Cost_Workbook.xlsx", import.meta.url);
  const file = new File([await readFile(path)], basename(path.pathname));
  const sheet = (await analyseXlsx(file)).sheets.find((item) => item.name === "Summary - Include");
  assert.ok(sheet);
  assert.equal(sheet.printArea, undefined);
  assert.equal(sheet.range, "A1:H15");
  assert.ok(sheet.merges.some((merge) => merge.left === 1 && merge.top === 13 && merge.right === 8 && merge.bottom === 15));
  assert.match(sheet.cells.find((cell) => cell.row === 13 && cell.col === 1)?.value ?? "", /A workbook may contain several worksheet tabs/);
  assert.equal(sheet.renderPlan.orientation, "landscape");
  assert.equal(sheet.renderPlan.scalePercent, 88, "Carlito 11 column-width geometry selects the highest whole-percent scale that fits the 725pt canonical landscape budget");
  assert.equal(sheet.renderPlan.predictedPageCount, 1);
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

test("rejects entity-encoded external workbook relationships and active content", async () => {
  const { analyseXlsx } = await import("../app/lib/xlsx.ts");
  await assert.rejects(analyseXlsx(await syntheticWorkbook({
    dimension: "A1:B2",
    sheetRelationships: `<Relationships><Relationship Id="rId1" TargetMode="&#x45;xternal" Target="https://example.invalid"/></Relationships>`,
  })), /external relationship/i);
  await assert.rejects(analyseXlsx(await syntheticWorkbook({
    dimension: "A1:B2",
    extraEntries: { "[Content_Types].xml": `<Types ContentType="application/vnd.ms-excel.sheet.macr&#111;Enabled.main+xml"/>` },
  })), /active or externally connected/i);
});

test("workbook analysis still times out after 20 seconds when Worker is unavailable", async () => {
  const source = await readFile(new URL("../app/lib/xlsx.ts", import.meta.url), "utf8");
  const fallback = source.slice(source.indexOf("export function analyseXlsxInWorker"), source.indexOf("return new Promise((resolve, reject) => {"));
  assert.match(fallback, /typeof Worker === "undefined"/);
  assert.match(fallback, /Promise\.race/);
  assert.match(fallback, /20_000/);
});
