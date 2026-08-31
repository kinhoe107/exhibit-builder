import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { EventEmitter } from "node:events";
import { PassThrough } from "node:stream";
import test from "node:test";
import JSZip from "jszip";

const require = createRequire(import.meta.url);
const { DEFAULT_EXPORT_TIMEOUT_MS, EXPORT_SCRIPT, assertSafeWorkbookArchive, runPowerShell, validateRequest } = require("../electron/workbook-export.cjs");

async function workbookArchive(extra = {}) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", extra.contentTypes ?? '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
  zip.file("_rels/.rels", extra.relationships ?? '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
  zip.file("xl/workbook.xml", "<workbook/>");
  for (const [name, content] of Object.entries(extra.entries ?? {})) zip.file(name, content);
  return Buffer.from(await zip.generateAsync({ type: "uint8array" }));
}

test("native workbook export is read-only and disables executable workbook behaviour", () => {
  assert.match(EXPORT_SCRIPT, /AutomationSecurity = 3/);
  assert.match(EXPORT_SCRIPT, /EnableEvents = \$false/);
  assert.match(EXPORT_SCRIPT, /AskToUpdateLinks = \$false/);
  assert.match(EXPORT_SCRIPT, /Workbooks\.Open\(\$WorkbookPath, 0, \$true/);
  assert.match(EXPORT_SCRIPT, /BlackAndWhite = \$false/);
  assert.match(EXPORT_SCRIPT, /Zoom = \[int\]\$requested\.scalePercent/);
  assert.match(EXPORT_SCRIPT, /FitToPagesWide = \$false/);
  assert.match(EXPORT_SCRIPT, /ResetAllPageBreaks/);
  assert.ok(EXPORT_SCRIPT.indexOf("PageSetup.Zoom = [int]$requested.scalePercent") > EXPORT_SCRIPT.indexOf("ResetAllPageBreaks"), "Excel zoom is applied after reset because ResetAllPageBreaks restores it to 100%");
  assert.match(EXPORT_SCRIPT, /did not retain the approved worksheet zoom/);
  assert.match(EXPORT_SCRIPT, /VPageBreaks\.Add/);
  assert.match(EXPORT_SCRIPT, /HPageBreaks\.Add/);
  assert.match(EXPORT_SCRIPT, /PageSetup\.Order = 1/);
  assert.match(EXPORT_SCRIPT, /PageSetup\.PrintTitleRows/);
  assert.match(EXPORT_SCRIPT, /PageSetup\.PrintTitleColumns = ''/);
  assert.match(EXPORT_SCRIPT, /PageSetup\.LeftMargin = \$excel\.InchesToPoints/);
  assert.match(EXPORT_SCRIPT, /PageSetup\.PrintHeadings = \$false/);
  assert.match(EXPORT_SCRIPT, /PageSetup\.PrintComments = -4142/);
  assert.match(EXPORT_SCRIPT, /nativeRange\.Width/);
  assert.match(EXPORT_SCRIPT, /nativeRange\.Height/);
  assert.match(EXPORT_SCRIPT, /actualColumns/);
  assert.match(EXPORT_SCRIPT, /actualRows/);
  assert.match(EXPORT_SCRIPT, /relocated, added, or removed a planned vertical page break/);
  assert.match(EXPORT_SCRIPT, /ExportAsFixedFormat/);
  assert.match(EXPORT_SCRIPT, /excel\.pid/);
  assert.equal(DEFAULT_EXPORT_TIMEOUT_MS, 10 * 60 * 1000);
});

test("native workbook export times out and requests exact-process cleanup", async () => {
  const child = new EventEmitter();
  child.pid = 4242;
  child.stderr = new PassThrough();
  let terminated = null;
  const startedAt = Date.now();
  await assert.rejects(
    runPowerShell("export.ps1", "source.xlsx", "request.json", "C:\\export", {
      timeoutMs: 10,
      spawnProcess: () => child,
      terminate: async (processId, directory) => { terminated = { processId, directory }; child.emit("exit", 1); },
    }),
    /timed out/i,
  );
  assert.deepEqual(terminated, { processId: 4242, directory: "C:\\export" });
  assert.ok(Date.now() - startedAt < 1_000);
});

test("native workbook export accepts only bounded xlsx requests and safe ranges", () => {
  const valid = validateRequest("source.xlsx", new Uint8Array([1, 2, 3]), [{ name: "Programme", range: "A1:H12", scalePercent: 94, columnBreaks: [6, 6, -1], rowBreaks: [20, 20, 0], titleRows: "$1:$2", pageOrder: "overThenDown", margins: { left: 9 }, geometryChecks: [{ axis: "horizontal", ranges: ["A1:B1"] }], expectedPageCount: 4 }]);
  assert.equal(valid.cleanSheets[0].range, "A1:H12");
  assert.equal(valid.cleanSheets[0].scalePercent, 94);
  assert.deepEqual(valid.cleanSheets[0].columnBreaks, [6]);
  assert.deepEqual(valid.cleanSheets[0].rowBreaks, [20]);
  assert.equal(valid.cleanSheets[0].titleRows, "$1:$2");
  assert.equal(valid.cleanSheets[0].pageOrder, "downThenOver", "untrusted source order cannot alter deterministic pagination");
  assert.deepEqual(valid.cleanSheets[0].margins, { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 });
  assert.deepEqual(valid.cleanSheets[0].geometryChecks, [{ axis: "horizontal", ranges: ["A1:B1"] }]);
  assert.equal(valid.cleanSheets[0].expectedPageCount, 4);
  assert.throws(() => validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Programme", range: "A1:H12;Remove-Item" }]), /print range/i);
  assert.throws(() => validateRequest("source.xlsm", new Uint8Array([1]), [{ name: "Sheet1", range: "A1:B2" }]), /Only \.xlsx/i);
  assert.throws(() => validateRequest("source.xlsx", new Uint8Array(), [{ name: "Sheet1", range: "A1:B2" }]), /empty/i);
  assert.throws(() => validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Programme", range: "" }]), /print range/i, "native export never inherits an implicit source print range");
  assert.throws(() => validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Programme", range: "A1:B2", titleColumns: "$A:$B" }]), /title columns/i);
  assert.throws(() => validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Programme", range: "A1:B2", geometryChecks: [{ axis: "vertical", ranges: ["A1:B2;bad"] }], expectedPageCount: 1 }]), /unsafe cell range/i);
  assert.throws(() => validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Programme", range: "A1:B2" }]), /expected page count.*missing/i);
  assert.throws(() => validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Programme", range: "A1:B2", expectedPageCount: 20_001 }]), /outside.*20000/i);
  assert.throws(() => validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Programme", range: "A1:B20", rowBreaks: [10], expectedPageCount: 1 }]), /does not match.*2-page break grid/i);
  assert.doesNotThrow(() => validateRequest("source.xlsx", new Uint8Array([1]), Array.from({ length: 129 }, (_, index) => ({ name: `Sheet ${index + 1}`, range: "A1:B2", expectedPageCount: 1 }))));
  assert.throws(() => validateRequest("source.xlsx", new Uint8Array([1]), Array.from({ length: 513 }, (_, index) => ({ name: `Sheet ${index + 1}`, range: "A1:B2" }))), /print jobs/i);

  const moreThanLegacyCap = Array.from({ length: 250 }, (_, index) => index + 2);
  const completeBreakPlan = validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Long programme", range: "A1:A5000", rowBreaks: moreThanLegacyCap, expectedPageCount: 251 }]);
  assert.equal(completeBreakPlan.cleanSheets[0].rowBreaks.length, moreThanLegacyCap.length, "valid unique row breaks are never silently truncated at 200");
  assert.deepEqual(completeBreakPlan.cleanSheets[0].rowBreaks, moreThanLegacyCap);
  assert.throws(
    () => validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Impossible programme", range: "A1:A5000", rowBreaks: Array.from({ length: 1_027 }, (_, index) => index + 2) }]),
    /more than 1026 unique valid page breaks/i,
  );
});

test("native workbook safety inspection is dependency-free and rejects active OOXML content", async () => {
  await assert.doesNotReject(assertSafeWorkbookArchive(await workbookArchive()));
  await assert.rejects(assertSafeWorkbookArchive(await workbookArchive({ entries: { "xl/macrosheets/sheet1.xml": "<worksheet/>" } })), /macros|embedded|external/i);
  await assert.rejects(assertSafeWorkbookArchive(await workbookArchive({ relationships: '<Relationships><Relationship TargetMode="External" Target="https://example.invalid"/></Relationships>' })), /external relationship/i);
  await assert.rejects(assertSafeWorkbookArchive(await workbookArchive({ relationships: '<Relationships><Relationship TargetMode="&#x45;xternal" Target="https://example.invalid"/></Relationships>' })), /external relationship/i);
  await assert.rejects(assertSafeWorkbookArchive(await workbookArchive({ contentTypes: '<Types ContentType="application/vnd.ms-excel.sheet.macr&#111;Enabled.main+xml"/>' })), /active or externally connected/i);
  await assert.rejects(assertSafeWorkbookArchive(Buffer.from("not-a-zip")), /malformed|safe OOXML/i);
});

test("renderer and Electron workbook screens decode XML before matching external relationships", async () => {
  const { readFile } = await import("node:fs/promises");
  const renderer = await readFile(new URL("../app/lib/xlsx.ts", import.meta.url), "utf8");
  const exporter = await readFile(new URL("../electron/workbook-export.cjs", import.meta.url), "utf8");
  const marker = /const WORKBOOK_EXTERNAL_RELATIONSHIP = \/TargetMode\\s\*=\\s\*\(\?:\["'\]\\s\*External\\s\*\["'\]\|External\\b\)\/i;/;
  assert.match(renderer, marker);
  assert.match(exporter, marker);
  assert.match(renderer, /function decodeWorkbookXml/);
  assert.match(exporter, /function decodeWorkbookXml/);
});
