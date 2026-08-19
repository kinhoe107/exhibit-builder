const { spawn } = require("node:child_process");
const { randomUUID } = require("node:crypto");
const { mkdir, readFile, rm, writeFile } = require("node:fs/promises");
const path = require("node:path");

const MAX_WORKBOOK_BYTES = 512 * 1024 * 1024;
const MAX_SHEETS = 128;
const DEFAULT_EXPORT_TIMEOUT_MS = 10 * 60 * 1000;
const WINDOWS_ROOT = process.env.SystemRoot || process.env.WINDIR || "C:\\Windows";
const POWERSHELL_PATH = path.join(WINDOWS_ROOT, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
const TASKKILL_PATH = path.join(WINDOWS_ROOT, "System32", "taskkill.exe");
const activeExports = new Map();

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
    if ([string]$requested.range) { $sheet.PageSetup.PrintArea = [string]$requested.range }
    $sheet.PageSetup.BlackAndWhite = $false
    $sheet.PageSetup.PaperSize = 9
    $sheet.PageSetup.Orientation = if ([string]$requested.orientation -eq 'landscape') { 2 } else { 1 }
    $sheet.PageSetup.Zoom = $false
    $sheet.PageSetup.FitToPagesWide = 1
    $sheet.PageSetup.FitToPagesTall = $false
    $outputPath = Join-Path $OutputDirectory ([string]$requested.outputName)
    $sheet.ExportAsFixedFormat(0, $outputPath, 0, $true, $false)
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
  if (!Array.isArray(sheets) || !sheets.length || sheets.length > MAX_SHEETS) throw new Error("Choose at least one worksheet and no more than 128 worksheets.");
  const cleanSheets = sheets.map((sheet, index) => ({
    name: boundedText(sheet?.name, "Worksheet name", 256),
    range: typeof sheet?.range === "string" && /^\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$/i.test(sheet.range.trim()) ? sheet.range.trim() : "",
    orientation: sheet?.orientation === "landscape" ? "landscape" : "portrait",
    outputName: `sheet-${String(index + 1).padStart(3, "0")}.pdf`,
  }));
  return { input, cleanSheets };
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

module.exports = { DEFAULT_EXPORT_TIMEOUT_MS, EXPORT_SCRIPT, exportWorkbookSheets, runPowerShell, shutdownAllExports, terminateExportProcesses, validateRequest };
