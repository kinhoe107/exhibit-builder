import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { EventEmitter } from "node:events";
import { PassThrough } from "node:stream";
import test from "node:test";

const require = createRequire(import.meta.url);
const { DEFAULT_EXPORT_TIMEOUT_MS, EXPORT_SCRIPT, runPowerShell, validateRequest } = require("../electron/workbook-export.cjs");

test("native workbook export is read-only and disables executable workbook behaviour", () => {
  assert.match(EXPORT_SCRIPT, /AutomationSecurity = 3/);
  assert.match(EXPORT_SCRIPT, /EnableEvents = \$false/);
  assert.match(EXPORT_SCRIPT, /AskToUpdateLinks = \$false/);
  assert.match(EXPORT_SCRIPT, /Workbooks\.Open\(\$WorkbookPath, 0, \$true/);
  assert.match(EXPORT_SCRIPT, /BlackAndWhite = \$false/);
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
  const valid = validateRequest("source.xlsx", new Uint8Array([1, 2, 3]), [{ name: "Programme", range: "A1:H12" }]);
  assert.equal(valid.cleanSheets[0].range, "A1:H12");
  const ignoredRange = validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Programme", range: "A1:H12;Remove-Item" }]);
  assert.equal(ignoredRange.cleanSheets[0].range, "");
  assert.throws(() => validateRequest("source.xlsm", new Uint8Array([1]), [{ name: "Sheet1", range: "A1:B2" }]), /Only \.xlsx/i);
  assert.throws(() => validateRequest("source.xlsx", new Uint8Array(), [{ name: "Sheet1", range: "A1:B2" }]), /empty/i);
  const automatic = validateRequest("source.xlsx", new Uint8Array([1]), [{ name: "Programme", range: "" }]);
  assert.equal(automatic.cleanSheets[0].range, "", "an empty range lets Microsoft Excel use its native printable area");
});
