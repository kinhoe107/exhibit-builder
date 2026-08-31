const { spawn } = require("node:child_process");
const { randomUUID } = require("node:crypto");
const { mkdir, readFile, rm, writeFile } = require("node:fs/promises");
const path = require("node:path");
const { inflateRawSync } = require("node:zlib");

const MAX_WORKBOOK_BYTES = 512 * 1024 * 1024;
const MAX_PRINT_JOBS = 512;
const MAX_EXCEL_PAGE_BREAKS = 1_026;
const MAX_GEOMETRY_CHECKS = 20_000;
const CANONICAL_MARGINS = Object.freeze({ left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 });
const DEFAULT_EXPORT_TIMEOUT_MS = 10 * 60 * 1000;
const WINDOWS_ROOT = process.env.SystemRoot || process.env.WINDIR || "C:\\Windows";
const POWERSHELL_PATH = path.join(WINDOWS_ROOT, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
const TASKKILL_PATH = path.join(WINDOWS_ROOT, "System32", "taskkill.exe");
const activeExports = new Map();
const DISALLOWED_WORKBOOK_ENTRY = /^(?:xl\/(?:vbaProject\.bin|macrosheets\/|intlmacrosheets\/|activeX\/|embeddings\/|externalLinks\/|connections\.xml)|customXml\/)/i;
const DISALLOWED_WORKBOOK_CONTENT = /(?:macroEnabled|macrosheet|intlmacrosheet|activeX|oleObject|externalLink|connections?)/i;
const WORKBOOK_EXTERNAL_RELATIONSHIP = /TargetMode\s*=\s*(?:["']\s*External\s*["']|External\b)/i;
function decodeWorkbookXml(value) {
  let current = String(value ?? "");
  for (let pass = 0; pass < 4; pass += 1) {
    const next = current
      .replace(/&#x([0-9a-f]+);/gi, (_, n) => {
        const code = parseInt(n, 16);
        return Number.isInteger(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : "";
      })
      .replace(/&#(\d+);/g, (_, n) => {
        const code = Number(n);
        return Number.isInteger(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : "";
      })
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, "&");
    if (next === current) return current;
    current = next;
  }
  return current;
}
const MAX_ARCHIVE_ENTRIES = 20_000;
const MAX_INSPECTED_XML_BYTES = 16 * 1024 * 1024;

// Excel is used only as a local, read-only print engine. Macros, events,
// external-link updates and recalculation are disabled before the workbook is
// opened. The source file is never opened directly: the renderer supplies
// bytes which are written to the disposable Electron session directory.
const EXPORT_SCRIPT = String.raw`
param([string]$WorkbookPath, [string]$SpecificationPath, [string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {
  $specification = Get-Content -LiteralPath $SpecificationPath -Raw | ConvertFrom-Json
  $excel = New-Object -ComObject Excel.Application
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExhibitBuilderNativeMethods {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
  [uint32]$excelProcessId = 0
  [void][ExhibitBuilderNativeMethods]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelProcessId)
  if ($excelProcessId -gt 0) { Set-Content -LiteralPath (Join-Path $OutputDirectory 'excel.pid') -Value $excelProcessId -NoNewline }
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.EnableEvents = $false
  $excel.AskToUpdateLinks = $false
  $excel.AutomationSecurity = 3
  try { $excel.Calculation = -4135 } catch {}
  $workbook = $excel.Workbooks.Open($WorkbookPath, 0, $true, 5, '', '', $true, 2, '', $false, $false, $false, $false, $false, 0)
	  foreach ($requested in $specification.sheets) {
	    $sheet = $workbook.Worksheets.Item([string]$requested.name)
	    $sheet.PageSetup.PrintArea = [string]$requested.range
	    $sheet.PageSetup.BlackAndWhite = $false
	    $sheet.PageSetup.PaperSize = 9
	    $sheet.PageSetup.Orientation = if ([string]$requested.orientation -eq 'landscape') { 2 } else { 1 }
	    $sheet.PageSetup.Order = 1
	    $sheet.PageSetup.LeftMargin = $excel.InchesToPoints([double]$requested.margins.left)
	    $sheet.PageSetup.RightMargin = $excel.InchesToPoints([double]$requested.margins.right)
	    $sheet.PageSetup.TopMargin = $excel.InchesToPoints([double]$requested.margins.top)
	    $sheet.PageSetup.BottomMargin = $excel.InchesToPoints([double]$requested.margins.bottom)
	    $sheet.PageSetup.HeaderMargin = $excel.InchesToPoints([double]$requested.margins.header)
	    $sheet.PageSetup.FooterMargin = $excel.InchesToPoints([double]$requested.margins.footer)
	    $sheet.PageSetup.PrintTitleRows = [string]$requested.titleRows
	    $sheet.PageSetup.PrintTitleColumns = ''
	    $sheet.PageSetup.LeftHeader = ''
	    $sheet.PageSetup.CenterHeader = ''
	    $sheet.PageSetup.RightHeader = ''
	    $sheet.PageSetup.LeftFooter = ''
	    $sheet.PageSetup.CenterFooter = ''
	    $sheet.PageSetup.RightFooter = ''
	    $sheet.PageSetup.PrintHeadings = $false
	    $sheet.PageSetup.PrintComments = -4142
	    $sheet.PageSetup.CenterHorizontally = $false
	    $sheet.PageSetup.CenterVertically = $false
	    $sheet.PageSetup.Draft = $false
	    $sheet.PageSetup.FirstPageNumber = -4105
	    try { $sheet.PageSetup.OddAndEvenPagesHeaderFooter = $false } catch {}
	    try { $sheet.PageSetup.DifferentFirstPageHeaderFooter = $false } catch {}
	    $sheet.PageSetup.FitToPagesWide = $false
	    $sheet.PageSetup.FitToPagesTall = $false
	    foreach ($check in $requested.geometryChecks) {
	      [double]$points = 0
	      foreach ($range in $check.ranges) {
	        $nativeRange = $sheet.Range([string]$range)
	        $points += if ([string]$check.axis -eq 'horizontal') { [double]$nativeRange.Width } else { [double]$nativeRange.Height }
	        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($nativeRange)
	      }
	      [double]$effectivePoints = $points * ([double]$requested.scalePercent / 100.0)
	      [double]$limit = if ([string]$check.axis -eq 'horizontal') { [double]$requested.printableWidthPoints } else { [double]$requested.printableHeightPoints }
	      if ($effectivePoints -gt ($limit + 0.5)) { throw "Microsoft Excel measured unbreakable worksheet content at $([Math]::Round($points, 2)) raw points and $([Math]::Round($effectivePoints, 2)) points at the approved $([int]$requested.scalePercent)% scale, beyond the approved $([Math]::Round($limit, 2))-point $([string]$check.axis) A4 budget." }
	    }
	    $sheet.ResetAllPageBreaks()
	    foreach ($column in $requested.columnBreaks) { [void]$sheet.VPageBreaks.Add($sheet.Cells.Item(1, [int]$column)) }
	    foreach ($row in $requested.rowBreaks) { [void]$sheet.HPageBreaks.Add($sheet.Cells.Item([int]$row, 1)) }
	    $sheet.PageSetup.Zoom = [int]$requested.scalePercent
	    if ([int]$sheet.PageSetup.Zoom -ne [int]$requested.scalePercent) { throw "Microsoft Excel did not retain the approved worksheet zoom." }
	    $sheet.Activate()
	    $sheet.DisplayPageBreaks = $false
	    $sheet.DisplayPageBreaks = $true
	    $outputPath = Join-Path $OutputDirectory ([string]$requested.outputName)
	    $sheet.ExportAsFixedFormat(0, $outputPath, 0, $true, $false)
	    $printRange = $sheet.Range([string]$requested.range)
	    [int]$leftColumn = $printRange.Column
	    [int]$rightColumn = $leftColumn + $printRange.Columns.Count - 1
	    [int]$topRow = $printRange.Row
	    [int]$bottomRow = $topRow + $printRange.Rows.Count - 1
	    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($printRange)
	    $actualColumns = @($sheet.VPageBreaks | ForEach-Object { [int]$_.Location.Column } | Where-Object { $_ -gt $leftColumn -and $_ -le $rightColumn } | Sort-Object -Unique)
	    $actualRows = @($sheet.HPageBreaks | ForEach-Object { [int]$_.Location.Row } | Where-Object { $_ -gt $topRow -and $_ -le $bottomRow } | Sort-Object -Unique)
	    $expectedColumns = @($requested.columnBreaks | ForEach-Object { [int]$_ } | Sort-Object -Unique)
	    $expectedRows = @($requested.rowBreaks | ForEach-Object { [int]$_ } | Sort-Object -Unique)
	    if ([string]::Join(',', $actualColumns) -ne [string]::Join(',', $expectedColumns)) { throw "Microsoft Excel relocated, added, or removed a planned vertical page break (expected: $([string]::Join(',', $expectedColumns)); actual: $([string]::Join(',', $actualColumns)); zoom: $($sheet.PageSetup.Zoom); orientation: $($sheet.PageSetup.Orientation); left margin: $([Math]::Round([double]$sheet.PageSetup.LeftMargin, 2)); right margin: $([Math]::Round([double]$sheet.PageSetup.RightMargin, 2)))." }
	    if ([string]::Join(',', $actualRows) -ne [string]::Join(',', $expectedRows)) { throw "Microsoft Excel relocated, added, or removed a planned horizontal page break (expected: $([string]::Join(',', $expectedRows)); actual: $([string]::Join(',', $actualRows)))." }
	  }
} finally {
  if ($workbook -ne $null) { try { $workbook.Close($false) } catch {} }
  if ($excel -ne $null) { try { $excel.Quit() } catch {} }
  if ($sheet -ne $null) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
  if ($workbook -ne $null) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
  if ($excel -ne $null) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
`;

function boundedText(value, label, maximum) {
  if (typeof value !== "string" || !value.trim() || value.length > maximum) throw new Error(`${label} is invalid.`);
  return value;
}

function validateRequest(fileName, bytes, sheets) {
  const safeName = boundedText(fileName, "Workbook filename", 1_024);
  if (!/\.xlsx$/i.test(safeName)) throw new Error("Only .xlsx workbook exhibits can be exported through Microsoft Excel.");
  const input = Buffer.from(bytes ?? []);
  if (!input.byteLength || input.byteLength > MAX_WORKBOOK_BYTES) throw new Error("The workbook is empty or exceeds the 512 MB export limit.");
  if (!Array.isArray(sheets) || !sheets.length || sheets.length > MAX_PRINT_JOBS) throw new Error("Choose at least one worksheet print job and no more than 512 print jobs.");
  const boundedBreaks = (values, maximum, label) => {
    if (!Array.isArray(values)) return [];
    const unique = [...new Set(values.filter((value) => Number.isInteger(value) && value > 0 && value <= maximum))];
    if (unique.length > MAX_EXCEL_PAGE_BREAKS) {
      throw new Error(`${label} contains more than ${MAX_EXCEL_PAGE_BREAKS} unique valid page breaks and cannot be printed faithfully by Microsoft Excel.`);
    }
    // Never silently truncate a render plan: every admitted planned break is
    // serialized into the native Excel request in its original order.
    return unique;
  };
  const safeRange = (value) => typeof value === "string" && /^\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$/i.test(value.trim()) ? value.trim() : "";
  const safeTitleRows = (value) => typeof value === "string" && /^\$?\d+:\$?\d+$/.test(value.trim()) ? value.trim() : "";
  const cleanGeometryChecks = (checks) => {
    if (!Array.isArray(checks)) return [];
    if (checks.length > MAX_GEOMETRY_CHECKS) throw new Error(`Worksheet geometry contains more than ${MAX_GEOMETRY_CHECKS} checks.`);
    return checks.map((check) => {
      if (check?.axis !== "horizontal" && check?.axis !== "vertical") throw new Error("Worksheet geometry check axis is invalid.");
      if (!Array.isArray(check.ranges) || !check.ranges.length || check.ranges.length > 2) throw new Error("Worksheet geometry check ranges are invalid.");
      const ranges = check.ranges.map(safeRange);
      if (ranges.some((range) => !range)) throw new Error("Worksheet geometry check contains an unsafe cell range.");
      return { axis: check.axis, ranges };
    });
  };
  const cleanSheets = sheets.map((sheet, index) => {
    const range = safeRange(sheet?.range);
    if (!range) throw new Error("Worksheet print range is invalid or missing.");
    if (typeof sheet?.titleColumns === "string" && sheet.titleColumns.trim()) throw new Error("Repeated worksheet title columns are not supported.");
    const orientation = sheet?.orientation === "landscape" ? "landscape" : "portrait";
    const columnBreaks = boundedBreaks(sheet?.columnBreaks, 16_384, "Worksheet column-break plan");
    const rowBreaks = boundedBreaks(sheet?.rowBreaks, 1_048_576, "Worksheet row-break plan");
    if (!Number.isInteger(sheet?.expectedPageCount) || sheet.expectedPageCount < 1 || sheet.expectedPageCount > 20_000) throw new Error("Worksheet expected page count is missing or outside the 1–20000 page limit.");
    const breakGridPageCount = (columnBreaks.length + 1) * (rowBreaks.length + 1);
    if (sheet.expectedPageCount !== breakGridPageCount) throw new Error(`Worksheet expected page count ${sheet.expectedPageCount} does not match its ${breakGridPageCount}-page break grid.`);
    return {
      name: boundedText(sheet?.name, "Worksheet name", 256),
      range,
      orientation,
      scalePercent: Number.isInteger(sheet?.scalePercent) && sheet.scalePercent >= 85 && sheet.scalePercent <= 100 ? sheet.scalePercent : 100,
      columnBreaks,
      rowBreaks,
      titleRows: safeTitleRows(sheet?.titleRows),
      titleColumns: "",
      pageOrder: "downThenOver",
      margins: { ...CANONICAL_MARGINS },
      printableWidthPoints: orientation === "landscape" ? 725 : 480,
      printableHeightPoints: orientation === "landscape" ? 430 : 675,
      geometryChecks: cleanGeometryChecks(sheet?.geometryChecks),
      expectedPageCount: sheet.expectedPageCount,
      outputName: `sheet-${String(index + 1).padStart(3, "0")}.pdf`,
    };
  });
  return { input, cleanSheets };
}

function workbookArchiveEntries(input) {
  const minimumEocd = 22;
  const searchStart = Math.max(0, input.length - 65_557);
  let eocd = -1;
  for (let offset = input.length - minimumEocd; offset >= searchStart; offset -= 1) {
    if (input.readUInt32LE(offset) === 0x06054b50) { eocd = offset; break; }
  }
  if (eocd < 0 || eocd + minimumEocd > input.length) throw new Error("Missing workbook archive directory.");
  const disk = input.readUInt16LE(eocd + 4);
  const centralDisk = input.readUInt16LE(eocd + 6);
  const count = input.readUInt16LE(eocd + 10);
  const centralSize = input.readUInt32LE(eocd + 12);
  const centralOffset = input.readUInt32LE(eocd + 16);
  if (disk !== 0 || centralDisk !== 0 || count === 0xffff || centralOffset === 0xffffffff || centralSize === 0xffffffff) {
    throw new Error("Unsupported workbook archive layout.");
  }
  if (!count || count > MAX_ARCHIVE_ENTRIES || centralOffset + centralSize > eocd) throw new Error("Invalid workbook archive directory.");
  const entries = [];
  let offset = centralOffset;
  for (let index = 0; index < count; index += 1) {
    if (offset + 46 > input.length || input.readUInt32LE(offset) !== 0x02014b50) throw new Error("Invalid workbook archive entry.");
    const flags = input.readUInt16LE(offset + 8);
    const method = input.readUInt16LE(offset + 10);
    const compressedSize = input.readUInt32LE(offset + 20);
    const uncompressedSize = input.readUInt32LE(offset + 24);
    const nameLength = input.readUInt16LE(offset + 28);
    const extraLength = input.readUInt16LE(offset + 30);
    const commentLength = input.readUInt16LE(offset + 32);
    const localOffset = input.readUInt32LE(offset + 42);
    const end = offset + 46 + nameLength + extraLength + commentLength;
    if ((flags & 1) !== 0 || ![0, 8].includes(method) || end > input.length || localOffset === 0xffffffff) throw new Error("Unsupported workbook archive entry.");
    const name = input.subarray(offset + 46, offset + 46 + nameLength).toString("utf8").replace(/\\/g, "/");
    if (!name || name.includes("../") || name.startsWith("/") || name.includes("\0")) throw new Error("Unsafe workbook archive path.");
    entries.push({ name, method, compressedSize, uncompressedSize, localOffset });
    offset = end;
  }
  if (offset !== centralOffset + centralSize) throw new Error("Workbook archive directory does not reconcile.");
  return entries;
}

function workbookArchiveText(input, entry) {
  if (entry.uncompressedSize > MAX_INSPECTED_XML_BYTES) throw new Error("Workbook safety metadata is too large to inspect.");
  const offset = entry.localOffset;
  if (offset + 30 > input.length || input.readUInt32LE(offset) !== 0x04034b50) throw new Error("Invalid workbook archive content.");
  const nameLength = input.readUInt16LE(offset + 26);
  const extraLength = input.readUInt16LE(offset + 28);
  const start = offset + 30 + nameLength + extraLength;
  const end = start + entry.compressedSize;
  if (end > input.length) throw new Error("Truncated workbook archive content.");
  const compressed = input.subarray(start, end);
  const data = entry.method === 0
    ? compressed
    : inflateRawSync(compressed, { maxOutputLength: MAX_INSPECTED_XML_BYTES });
  if (data.length !== entry.uncompressedSize) throw new Error("Workbook archive content size does not reconcile.");
  return data.toString("utf8");
}

async function assertSafeWorkbookArchive(input) {
  let entries;
  try {
    entries = workbookArchiveEntries(input);
  } catch {
    throw new Error("The workbook is malformed or is not a safe OOXML archive.");
  }
  if (entries.some((entry) => DISALLOWED_WORKBOOK_ENTRY.test(entry.name))) {
    throw new Error("Workbook contains macros, embedded objects, external links or connections and cannot be opened for printing.");
  }
  const contentTypeEntry = entries.find((entry) => entry.name.toLowerCase() === "[content_types].xml");
  if (!contentTypeEntry) throw new Error("The workbook has no OOXML content-type manifest.");
  const contentTypes = decodeWorkbookXml(workbookArchiveText(input, contentTypeEntry));
  if (DISALLOWED_WORKBOOK_CONTENT.test(contentTypes)) {
    throw new Error("Workbook contains active or externally connected content and cannot be opened for printing.");
  }
  for (const relationship of entries.filter((entry) => /\.rels$/i.test(entry.name))) {
    const text = decodeWorkbookXml(workbookArchiveText(input, relationship));
    if (WORKBOOK_EXTERNAL_RELATIONSHIP.test(text)) {
      throw new Error("Workbook contains an external relationship and cannot be opened for printing.");
    }
  }
}

function stopProcess(processId, spawnProcess = spawn) {
  if (!Number.isSafeInteger(processId) || processId <= 0) return Promise.resolve();
  return new Promise((resolve) => {
    let settled = false;
    const finish = () => { if (!settled) { settled = true; resolve(); } };
    const killer = spawnProcess(TASKKILL_PATH, ["/PID", String(processId), "/T", "/F"], { windowsHide: true, stdio: "ignore" });
    const timer = setTimeout(() => { try { killer.kill(); } catch {} finish(); }, 15_000);
    timer.unref?.();
    killer.on("error", () => { clearTimeout(timer); finish(); });
    killer.on("exit", () => { clearTimeout(timer); finish(); });
  });
}

async function terminateExportProcesses(powerShellPid, outputDirectory, spawnProcess = spawn) {
  let excelPid = 0;
  try {
    const value = (await readFile(path.join(outputDirectory, "excel.pid"), "utf8")).trim();
    if (/^\d+$/.test(value)) excelPid = Number(value);
  } catch {}
  // Kill the recorded Excel COM server as well as the PowerShell process tree.
  // These calls are deliberately exact-PID only; no unrelated Excel session is touched.
  if (excelPid && excelPid !== powerShellPid) await stopProcess(excelPid, spawnProcess);
  await stopProcess(powerShellPid, spawnProcess);
}

function runPowerShell(scriptPath, workbookPath, specificationPath, outputDirectory, options = {}) {
  const spawnProcess = options.spawnProcess ?? spawn;
  const timeoutMs = options.timeoutMs ?? DEFAULT_EXPORT_TIMEOUT_MS;
  const terminate = options.terminate ?? ((processId, directory) => terminateExportProcesses(processId, directory, spawnProcess));
  return new Promise((resolve, reject) => {
    const child = spawnProcess(POWERSHELL_PATH, [
      "-NoLogo",
      "-NoProfile",
      "-NonInteractive",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      scriptPath,
      workbookPath,
      specificationPath,
      outputDirectory,
    ], { windowsHide: true, stdio: ["ignore", "pipe", "pipe"] });
    if (Number.isSafeInteger(child.pid) && child.pid > 0) activeExports.set(child.pid, outputDirectory);
    let errorText = "";
    let settled = false;
    let terminating = false;
    const finish = (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (Number.isSafeInteger(child.pid)) activeExports.delete(child.pid);
      if (error) reject(error);
      else resolve();
    };
    child.stderr.setEncoding("utf8");
    child.stderr.on("data", (chunk) => { if (errorText.length < 16_000) errorText += chunk; });
    const timer = setTimeout(() => {
      if (settled || terminating) return;
      terminating = true;
      void Promise.resolve(terminate(child.pid, outputDirectory)).finally(() => {
        finish(new Error(`Microsoft Excel export timed out after ${Math.ceil(timeoutMs / 60_000)} minutes. The dedicated export processes were stopped.`));
      });
    }, timeoutMs);
    timer.unref?.();
    child.on("error", (error) => { if (!terminating) finish(error); });
    child.on("exit", (code) => {
      if (settled || terminating) return;
      if (code === 0) finish();
      else {
        terminating = true;
        clearTimeout(timer);
        void Promise.resolve(terminate(child.pid, outputDirectory)).finally(() => finish(new Error(errorText.trim() || `Microsoft Excel export exited with code ${code}.`)));
      }
    });
  });
}

async function shutdownAllExports() {
  const active = Array.from(activeExports.entries());
  await Promise.all(active.map(async ([processId, directory]) => {
    await terminateExportProcesses(processId, directory);
    activeExports.delete(processId);
    await rm(directory, { recursive: true, force: true, maxRetries: 2, retryDelay: 100 }).catch(() => {});
  }));
  return { stopped: active.length };
}

async function exportWorkbookSheets(sessionRoot, { fileName, bytes, sheets }) {
  if (!path.isAbsolute(sessionRoot)) throw new Error("Workbook export requires an absolute session directory.");
  const { input, cleanSheets } = validateRequest(fileName, bytes, sheets);
  await assertSafeWorkbookArchive(input);
  const exportRoot = path.join(sessionRoot, `workbook-export-${randomUUID()}`);
  const workbookPath = path.join(exportRoot, "source.xlsx");
  const specificationPath = path.join(exportRoot, "request.json");
  const scriptPath = path.join(exportRoot, "export.ps1");
  await mkdir(exportRoot, { recursive: true });
  try {
    await writeFile(workbookPath, input, { flag: "wx" });
    await writeFile(specificationPath, JSON.stringify({ sheets: cleanSheets }), { flag: "wx" });
    await writeFile(scriptPath, EXPORT_SCRIPT, { flag: "wx" });
    try {
      await runPowerShell(scriptPath, workbookPath, specificationPath, exportRoot);
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      throw new Error(`Microsoft Excel could not print this workbook faithfully. No simplified copy was substituted. Open the workbook in Excel, confirm it prints correctly, then retry. ${detail}`);
    }
    return Promise.all(cleanSheets.map(async (sheet) => ({
      name: sheet.name,
      range: sheet.range,
      orientation: sheet.orientation,
      bytes: new Uint8Array(await readFile(path.join(exportRoot, sheet.outputName))),
    })));
  } finally {
    await rm(exportRoot, { recursive: true, force: true, maxRetries: 2, retryDelay: 100 }).catch(() => {});
  }
}

module.exports = { DEFAULT_EXPORT_TIMEOUT_MS, EXPORT_SCRIPT, assertSafeWorkbookArchive, exportWorkbookSheets, runPowerShell, shutdownAllExports, terminateExportProcesses, validateRequest };
