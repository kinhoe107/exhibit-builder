import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const appRoot = path.resolve(scriptDir, "..");
const outputPath = path.join(appRoot, "public", "guided-sample", "Exhibits", "05_SAMPLE_Cost_Workbook.xlsx");
const previewDir = path.join(appRoot, "tmp", "guided-sample-workbook-previews");

const navy = "#17365D";
const blue = "#1F4E79";
const paleBlue = "#EAF2F8";
const paleOrange = "#FFF0E8";
const orange = "#7A351F";
const border = "#B8C4D0";

function styleBanner(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.format.fill = paleOrange;
  range.format.font = { name: "Aptos", size: 10, bold: true, color: orange };
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
  range.format.rowHeight = 24;
  range.format.borders = { preset: "outside", style: "thin", color: "#CC8C75" };
}

function styleTitle(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.format.fill = navy;
  range.format.font = { name: "Aptos Display", size: 17, bold: true, color: "#FFFFFF" };
  range.format.verticalAlignment = "center";
  range.format.rowHeight = 32;
}

function styleHeader(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.format.fill = blue;
  range.format.font = { name: "Aptos", size: 10, bold: true, color: "#FFFFFF" };
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
  range.format.wrapText = true;
  range.format.rowHeight = 26;
  range.format.borders = { preset: "all", style: "thin", color: border };
}

function styleBody(sheet, rangeAddress) {
  const range = sheet.getRange(rangeAddress);
  range.format.font = { name: "Aptos", size: 10, color: "#182433" };
  range.format.verticalAlignment = "center";
  range.format.wrapText = true;
  range.format.borders = { preset: "all", style: "thin", color: "#D7E1EA" };
  range.format.rowHeight = 24;
}

const workbook = Workbook.create();
const summary = workbook.worksheets.add("Summary - Include");
const detail = workbook.worksheets.add("Detail - Optional");
const notes = workbook.worksheets.add("Working Notes - Exclude");

for (const sheet of [summary, detail, notes]) {
  sheet.showGridLines = false;
}

summary.getRange("A1:H1").merge();
summary.getRange("A1").values = [["SAMPLE DOCUMENT - FOR EXHIBIT BUILDER DEMONSTRATION ONLY"]];
styleBanner(summary, "A1:H1");
summary.getRange("A3:H3").merge();
summary.getRange("A3").values = [["SAMPLE COST WORKBOOK"]];
styleTitle(summary, "A3:H3");
summary.getRange("A4:H4").merge();
summary.getRange("A4").values = [["XLSX evidence example | 5 August 2026 | Reference DEMO-WORKBOOK-05"]];
summary.getRange("A4:H4").format.font = { name: "Aptos", size: 10, italic: true, color: "#4F5968" };
summary.getRange("A4:H4").format.fill = paleBlue;
summary.getRange("A6:H6").values = [["Cost code", "Description", "Quantity", "Unit", "Unit cost", "Net cost", "Status", "Guide note"]];
styleHeader(summary, "A6:H6");
summary.getRange("A7:H11").values = [
  ["C-101", "Document review", 6, "hours", 120, null, "Approved", "This sheet is selected for the guided exhibit."],
  ["C-102", "Index preparation", 4, "hours", 110, null, "Approved", "Colours, borders and number formats should be preserved."],
  ["C-103", "PDF quality checks", 3, "hours", 135, null, "Pending", "Saved formula results are used without changing the workbook."],
  ["C-104", "Workbook verification", 2, "hours", 145, null, "Approved", "The source file remains read-only."],
  ["TOTAL", "Guided sample total", null, null, null, null, "", "Choose worksheet tabs at the Sheets stage."],
];
summary.getRange("F7").formulas = [["=C7*E7"]];
summary.getRange("F7:F10").fillDown();
summary.getRange("F11").formulas = [["=SUM(F7:F10)"]];
styleBody(summary, "A7:H11");
summary.getRange("E7:F11").setNumberFormat("£#,##0.00");
summary.getRange("A11:H11").format.fill = paleBlue;
summary.getRange("A11:H11").format.font = { name: "Aptos", size: 10, bold: true, color: navy };
summary.getRange("A13:H15").merge();
summary.getRange("A13").values = [["What sheet selection means\nA workbook may contain several worksheet tabs. Tick the tabs that belong in the exhibit. Only the selected tabs and their shown cell ranges are printed into A4 bundle pages; unticked tabs are omitted. The workbook itself is not edited."]];
summary.getRange("A13:H15").format.fill = "#F4F6F9";
summary.getRange("A13:H15").format.font = { name: "Aptos", size: 10, color: navy };
summary.getRange("A13:H15").format.wrapText = true;
summary.getRange("A13:H15").format.verticalAlignment = "center";
summary.getRange("A13:H15").format.borders = { preset: "outside", style: "thin", color: border };
for (const [range, width] of [["A:A", 13], ["B:B", 24], ["C:C", 11], ["D:D", 11], ["E:F", 13], ["G:G", 13], ["H:H", 38]]) summary.getRange(range).format.columnWidth = width;
summary.freezePanes.freezeRows(6);

detail.getRange("A1:F1").merge();
detail.getRange("A1").values = [["SAMPLE DOCUMENT - FOR EXHIBIT BUILDER DEMONSTRATION ONLY"]];
styleBanner(detail, "A1:F1");
detail.getRange("A3:F3").merge();
detail.getRange("A3").values = [["SUPPORTING COST DETAIL - OPTIONAL SHEET"]];
styleTitle(detail, "A3:F3");
detail.getRange("A4:F4").merge();
detail.getRange("A4").values = [["This worksheet tab is valid evidence but is left unselected in the guided project. A reviewer can include it if required."]];
detail.getRange("A4:F4").format.fill = paleBlue;
detail.getRange("A4:F4").format.font = { name: "Aptos", size: 10, italic: true, color: "#4F5968" };
detail.getRange("A6:F6").values = [["Entry", "Date", "Work item", "Owner", "Hours", "Comment"]];
styleHeader(detail, "A6:F6");
detail.getRange("A7:F10").values = [
  [1, new Date("2026-08-01T12:00:00Z"), "Document review", "Sample team", 3, "Supporting row"],
  [2, new Date("2026-08-02T12:00:00Z"), "Document review", "Sample team", 3, "Supporting row"],
  [3, new Date("2026-08-03T12:00:00Z"), "Index preparation", "Sample team", 4, "Supporting row"],
  [4, new Date("2026-08-04T12:00:00Z"), "Quality checks", "Sample team", 3, "Supporting row"],
];
styleBody(detail, "A7:F10");
detail.getRange("B7:B10").setNumberFormat("dd mmm yyyy");
for (const [range, width] of [["A:A", 9], ["B:B", 15], ["C:C", 24], ["D:D", 17], ["E:E", 10], ["F:F", 28]]) detail.getRange(range).format.columnWidth = width;
detail.freezePanes.freezeRows(6);

notes.getRange("A1:E1").merge();
notes.getRange("A1").values = [["SAMPLE DOCUMENT - FOR EXHIBIT BUILDER DEMONSTRATION ONLY"]];
styleBanner(notes, "A1:E1");
notes.getRange("A3:E3").merge();
notes.getRange("A3").values = [["WORKING NOTES - EXCLUDE FROM THE EXHIBIT"]];
styleTitle(notes, "A3:E3");
notes.getRange("A5:E8").merge();
notes.getRange("A5").values = [["This worksheet demonstrates omission. It remains part of the source workbook but is not printed into the guided exhibit because its checkbox is cleared at the Sheets stage. Sheet selection never deletes or edits a source tab."]];
notes.getRange("A5:E8").format.fill = "#FFF4CE";
notes.getRange("A5:E8").format.font = { name: "Aptos", size: 12, bold: true, color: "#6B5200" };
notes.getRange("A5:E8").format.wrapText = true;
notes.getRange("A5:E8").format.verticalAlignment = "center";
notes.getRange("A5:E8").format.horizontalAlignment = "center";
notes.getRange("A5:E8").format.borders = { preset: "outside", style: "medium", color: "#D6B656" };
notes.getRange("A:E").format.columnWidth = 20;

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.mkdir(previewDir, { recursive: true });
const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);

const inspection = await workbook.inspect({ kind: "sheet,region,formula", maxChars: 7000, tableMaxRows: 18, tableMaxCols: 10 });
console.log(inspection.ndjson ?? inspection);
for (const sheetName of ["Summary - Include", "Detail - Optional", "Working Notes - Exclude"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1.4, format: "png" });
  const safeName = sheetName.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
  await fs.writeFile(path.join(previewDir, `${safeName}.png`), new Uint8Array(await preview.arrayBuffer()));
}

console.log(`Saved ${outputPath}`);
