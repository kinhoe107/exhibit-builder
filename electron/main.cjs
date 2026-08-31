const { app, BrowserWindow, clipboard, dialog, ipcMain, session, shell } = require("electron");
const { createServer } = require("node:http");
const { randomUUID } = require("node:crypto");
const { appendFileSync, existsSync, mkdirSync, mkdtempSync, readdirSync, rmSync, statSync } = require("node:fs");
const { readFile, rename, rm, stat, writeFile, mkdir } = require("node:fs/promises");
const { tmpdir } = require("node:os");
const path = require("node:path");
const { RecoveryStore } = require("./recovery-journal.cjs");
const { PreferenceStore } = require("./preferences.cjs");
const { exportWorkbookSheets, shutdownAllExports } = require("./workbook-export.cjs");
const { installDiagnosticStreamErrorHandlers } = require("./diagnostic-stream.cjs");
const {
  SourcePathCollisionError,
  parseProtectedSourcePaths,
  parseOptionalAbsolutePath,
  assertDestinationAllowed,
} = require("./protected-save-paths.cjs");

const DIST_ROOT = path.join(__dirname, "..", "dist");
const SMOKE_TEST = process.argv.includes("--smoke-test");
const ANALYSIS_SMOKE_TEST = process.argv.includes("--analysis-smoke-test");
const BUILD_SMOKE_TEST = process.argv.includes("--build-smoke-test");
const GUIDED_SMOKE_TEST = ANALYSIS_SMOKE_TEST || BUILD_SMOKE_TEST;
const SESSION_PREFIX = "bundle-builder-session-";
const DIAGNOSTIC_LOG = process.env.EXHIBIT_BUILDER_DIAGNOSTIC_LOG;
const WORKSPACE_LOAD_TIMEOUT_MS = 15_000;
const STABLE_APP_DATA_ROOT = app.getPath("appData");
const STABLE_DIAGNOSTIC_DIRECTORY = path.join(STABLE_APP_DATA_ROOT, "Exhibit Builder");
const STREAM_ERROR_LOG = path.join(STABLE_DIAGNOSTIC_DIRECTORY, "diagnostic-stream-errors.log");
const recoveryStore = new RecoveryStore(path.join(STABLE_APP_DATA_ROOT, "Exhibit Builder", "recovery"));
const preferenceStore = new PreferenceStore(path.join(STABLE_APP_DATA_ROOT, "Exhibit Builder"));


function diagnostic(stage, detail = {}) {
  if (!DIAGNOSTIC_LOG) return;
  try {
    appendFileSync(DIAGNOSTIC_LOG, `${JSON.stringify({ at: new Date().toISOString(), pid: process.pid, stage, ...detail })}\n`);
  } catch {
    // Diagnostics must never affect application startup.
  }
}

// Automated launchers may close their output pipe after collecting a smoke
// result. Only EPIPE is benign. Every other stream failure is written to an
// always-available local diagnostic and terminates through a controlled path.
installDiagnosticStreamErrorHandlers([process.stdout, process.stderr], {
  record(error) {
    mkdirSync(STABLE_DIAGNOSTIC_DIRECTORY, { recursive: true });
    appendFileSync(STREAM_ERROR_LOG, `${JSON.stringify({ at: new Date().toISOString(), pid: process.pid, code: error?.code ?? "unknown", message: error instanceof Error ? error.message : String(error) })}\n`);
    diagnostic("diagnostic-stream-error", { code: error?.code ?? "unknown" });
  },
  fatal(error) {
    process.exitCode = 1;
    if (!SMOKE_TEST && !GUIDED_SMOKE_TEST) {
      dialog.showErrorBox(
        "Exhibit Builder diagnostic failure",
        `A local application output stream failed (${error?.code ?? "unknown"}). Details were saved to ${STREAM_ERROR_LOG}. Exhibit Builder will close.`,
      );
    }
    app.exit(1);
  },
});

diagnostic("module-loaded", { smoke: SMOKE_TEST, analysisSmoke: ANALYSIS_SMOKE_TEST, buildSmoke: BUILD_SMOKE_TEST });

function cleanupStaleSessions() {
  diagnostic("stale-session-cleanup-start");
  const staleBefore = Date.now() - 60 * 60 * 1000;
  try {
    for (const entry of readdirSync(tmpdir(), { withFileTypes: true })) {
      if (!entry.isDirectory() || !entry.name.startsWith(SESSION_PREFIX)) continue;
      const directory = path.join(tmpdir(), entry.name);
      try {
        if (statSync(directory).mtimeMs < staleBefore) rmSync(directory, { recursive: true, force: true, maxRetries: 2, retryDelay: 100 });
      } catch {
        // A currently running or just-closed Chromium session may still hold
        // a lock. It will be reconsidered on the next launch.
      }
    }
  } catch {
    // Temporary-directory cleanup is best effort and must never block launch.
  }
  diagnostic("stale-session-cleanup-finished");
}

cleanupStaleSessions();
const SESSION_ROOT = mkdtempSync(path.join(tmpdir(), "bundle-builder-session-"));
const APP_PARTITION = "bundle-builder-memory";

// This app has no account or credential store. Avoid Chromium attempting a
// DPAPI-backed encryption operation in the disposable per-launch profile.
app.commandLine.appendSwitch("password-store", "basic");
// The affected host cannot launch Electron's sandboxed renderer
// (launch-failed/exit 49), and its GPU child exits with a missing-DLL status
// unless GPU work is kept in-process. The sole window below therefore keeps
// the narrow preload bridge, context isolation and disabled Node integration;
// it does not silently retry with a second renderer or security mode.
app.commandLine.appendSwitch("disable-gpu");
app.commandLine.appendSwitch("in-process-gpu");
app.disableHardwareAcceleration();
app.setPath("userData", SESSION_ROOT);
app.setPath("sessionData", SESSION_ROOT);
const MIME_TYPES = {
  ".css": "text/css; charset=utf-8",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".eml": "message/rfc822",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

// Electron renders PDF iframe/blob previews through Chromium's bundled PDF
// viewer. Its own scripts and styles use this fixed extension origin; the PDF
// bytes themselves remain on the already-allowed blob: URL. Allowing only this
// origin keeps every remote http(s) destination blocked.
const CHROMIUM_PDF_VIEWER_ORIGIN = "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai";

let localServer;
let localOrigin;
let mainWindow;
let sessionCleaned = false;
let unresponsiveDialogOpen = false;
let exportShutdownStarted = false;

const singleInstanceLock = app.requestSingleInstanceLock();
if (!singleInstanceLock) app.quit();
app.on("second-instance", () => {
  if (!mainWindow || mainWindow.isDestroyed()) return;
  if (mainWindow.isMinimized()) mainWindow.restore();
  mainWindow.focus();
});

function cleanupSession() {
  if (sessionCleaned) return;
  sessionCleaned = true;
  diagnostic("session-cleanup-start");
  try {
    rmSync(SESSION_ROOT, {
      recursive: true,
      force: true,
      maxRetries: 2,
      retryDelay: 100,
    });
  } catch (error) {
    // Windows can briefly retain Chromium file handles during process
    // teardown. A locked temporary shell is safe for Windows cleanup and
    // must never become a second application failure.
    if (!["EBUSY", "ENOTEMPTY", "EPERM"].includes(error?.code)) {
      console.error("Temporary session cleanup could not complete.", error);
    }
  }
  diagnostic("session-cleanup-finished");
}

function showUnresponsiveDialog(window) {
  if (unresponsiveDialogOpen || window.isDestroyed() || SMOKE_TEST || GUIDED_SMOKE_TEST) return;
  unresponsiveDialogOpen = true;
  void dialog.showMessageBox(window, {
    type: "warning",
    title: "Exhibit Builder is busy",
    message: "Exhibit Builder is taking longer than expected to respond.",
    detail: "If you have just started analysis or bundle assembly, the operation may still be working. You can keep waiting, reload the workspace, or close the application.",
    buttons: ["Keep waiting", "Reload workspace", "Close Exhibit Builder"],
    defaultId: 0,
    cancelId: 0,
  }).then(({ response }) => {
    if (response === 1 && !window.isDestroyed()) window.reload();
    if (response === 2 && !window.isDestroyed()) window.close();
  }).catch(() => {}).finally(() => {
    unresponsiveDialogOpen = false;
  });
}

function attachWindowDiagnostics(window) {
  window.webContents.on("unresponsive", () => {
    diagnostic("renderer-unresponsive");
    console.error("Exhibit Builder renderer became unresponsive.");
    showUnresponsiveDialog(window);
  });
  window.webContents.on("responsive", () => {
    diagnostic("renderer-responsive");
    console.error("Exhibit Builder renderer became responsive again.");
  });
  window.webContents.on("render-process-gone", (_event, details) => {
    if (details.reason === "clean-exit") return;
    diagnostic("renderer-gone", { reason: details.reason, exitCode: details.exitCode });
    if (SMOKE_TEST || GUIDED_SMOKE_TEST || window.isDestroyed()) return;
    console.error(`Exhibit Builder renderer exited: ${details.reason || "unknown"} (${details.exitCode ?? "n/a"}).`);
    void dialog.showMessageBox(window, {
      type: "error",
      title: "Exhibit Builder stopped responding",
      message: "The workspace renderer stopped unexpectedly.",
      detail: "No source statement is rewritten by this event. Close and reopen Exhibit Builder before retrying the analysis.",
      buttons: ["Close Exhibit Builder"],
    }).finally(() => {
      if (!window.isDestroyed()) window.close();
    });
  });
  window.webContents.on("did-fail-load", (_event, errorCode, errorDescription, validatedURL, isMainFrame) => {
    if (isMainFrame && errorCode !== -3) {
      diagnostic("load-failed", { errorCode, errorDescription });
      console.error(`Exhibit Builder failed to load the local workspace: ${errorCode} ${errorDescription} ${validatedURL}`);
    }
  });
}

function resolveAsset(requestUrl) {
  const pathname = decodeURIComponent(new URL(requestUrl, localOrigin).pathname);
  const relative = pathname === "/" ? "index.html" : pathname.replace(/^\/+/, "");
  const candidate = path.resolve(DIST_ROOT, relative);
  if (!candidate.startsWith(`${path.resolve(DIST_ROOT)}${path.sep}`)) {
    return null;
  }
  return candidate;
}

async function startLocalServer() {
  diagnostic("local-server-start");
  localServer = createServer(async (request, response) => {
    try {
      if (!request.url || !["GET", "HEAD"].includes(request.method || "GET")) {
        response.writeHead(405, { Allow: "GET, HEAD" }).end("Method not allowed");
        return;
      }
      let filePath = resolveAsset(request.url || "/");
      if (!filePath) {
        response.writeHead(403).end("Forbidden");
        return;
      }
      try {
        const details = await stat(filePath);
        if (details.isDirectory()) filePath = path.join(filePath, "index.html");
      } catch {
        const requestedPath = decodeURIComponent(new URL(request.url || "/", localOrigin).pathname);
        const looksLikeAsset = path.extname(requestedPath) !== "";
        if (looksLikeAsset) {
          response.writeHead(404).end("File not found");
          return;
        }
        filePath = path.join(DIST_ROOT, "index.html");
      }
      const body = await readFile(filePath);
      response.writeHead(200, {
        "Cache-Control": "no-store",
        "Content-Type":
          MIME_TYPES[path.extname(filePath).toLowerCase()] ||
          "application/octet-stream",
        "Cross-Origin-Opener-Policy": "same-origin",
        "X-Content-Type-Options": "nosniff",
      });
      response.end(request.method === "HEAD" ? undefined : body);
    } catch {
      response.writeHead(500).end("Exhibit Builder could not load this file.");
    }
  });

  await new Promise((resolve, reject) => {
    localServer.once("error", reject);
    localServer.listen(0, "127.0.0.1", () => resolve());
  });
  const address = localServer.address();
  localOrigin = `http://127.0.0.1:${address.port}`;
  diagnostic("local-server-ready");
}

function installNetworkBlock(appSession) {
  appSession.setPermissionRequestHandler(
    (_contents, _permission, callback) => callback(false),
  );
  appSession.webRequest.onBeforeRequest((details, callback) => {
    const allowed =
      details.url === localOrigin ||
      details.url.startsWith(`${localOrigin}/`) ||
      details.url === CHROMIUM_PDF_VIEWER_ORIGIN ||
      details.url.startsWith(`${CHROMIUM_PDF_VIEWER_ORIGIN}/`) ||
      details.url.startsWith("blob:") ||
      details.url.startsWith("data:");
    callback({ cancel: !allowed });
  });
}

function assertTrustedIpcSender(event, channel) {
  const senderUrl = event?.senderFrame?.url || event?.sender?.getURL?.();
  let trusted = false;
  try {
    trusted = Boolean(localOrigin) && new URL(senderUrl).origin === localOrigin;
  } catch {
    trusted = false;
  }
  if (trusted) return;
  diagnostic("ipc-rejected", { channel, senderUrl: typeof senderUrl === "string" ? senderUrl : "unknown" });
  throw new Error("This action is available only from the local Exhibit Builder workspace.");
}

function resolveAppIcon() {
  const packaged = path.join(path.dirname(process.execPath), "icon.ico");
  const unpackaged = path.join(__dirname, "..", "build", "icon.ico");
  if (app.isPackaged && existsSync(packaged)) return packaged;
  if (existsSync(unpackaged)) return unpackaged;
}

function resolveGuidedSampleDirectory() {
  const packaged = path.join(path.dirname(process.execPath), "Guided Sample");
  const unpackaged = path.join(__dirname, "..", "public", "guided-sample");
  if (app.isPackaged && existsSync(packaged)) return packaged;
  if (existsSync(unpackaged)) return unpackaged;
}

function makeWindow() {
  diagnostic("window-create-start", { sandbox: true });
  const window = new BrowserWindow({
    width: 1440,
    height: 940,
    minWidth: 980,
    minHeight: 680,
    backgroundColor: "#f7f4ee",
    show: false,
    title: "Exhibit Builder",
    icon: resolveAppIcon(),
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, "preload.cjs"),
      partition: APP_PARTITION,
      sandbox: true,
      spellcheck: false,
    },
  });

  window.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
  attachWindowDiagnostics(window);
  if (GUIDED_SMOKE_TEST) window.webContents.on("console-message", (_event, level, message, line, sourceId) => console.error(`[analysis-smoke renderer ${level}] ${message} (${sourceId}:${line})`));
  if (GUIDED_SMOKE_TEST) window.webContents.on("did-fail-load", (_event, code, description, url) => console.error(`[analysis-smoke load failure] ${code} ${description} ${url}`));
  window.webContents.on("will-navigate", (event, url) => {
    if (url !== localOrigin && !url.startsWith(`${localOrigin}/`)) event.preventDefault();
  });
  diagnostic("window-create-finished", { sandbox: true });
  return window;
}

const presentedMainWindows = new WeakSet();

function presentMainWindow(window) {
  if (window.isDestroyed()) return;
  const alreadyPresented = presentedMainWindows.has(window) && window.isVisible() && !window.isMinimized();
  if (alreadyPresented) return;
  presentedMainWindows.add(window);
  if (window.isFullScreen()) window.setFullScreen(false);
  if (window.isMinimized()) window.restore();
  window.show();
  window.maximize();
  window.focus();
}

function windowPresentationState(window) {
  return {
    windowVisible: window.isVisible(),
    windowMaximized: window.isMaximized(),
    windowFullScreen: window.isFullScreen(),
    windowMinimized: window.isMinimized(),
  };
}

async function createWindow() {
  diagnostic("window-load-start");
  // Keep one predictable, sandboxed configuration. A renderer launch failure
  // is surfaced as a bounded startup error; do not retry with weaker security.
  const window = makeWindow();
  const presented = new Promise((resolve) => {
    window.once("ready-to-show", () => {
      presentMainWindow(window);
      resolve();
    });
  });
  try {
    await loadWorkspace(window);
    await Promise.race([presented, new Promise((resolve) => setTimeout(resolve, 5000))]);
    presentMainWindow(window);
    diagnostic("window-load-finished");
  } catch (error) {
    diagnostic("window-load-failed", { message: error instanceof Error ? error.message : String(error), sandbox: true });
    if (!window.isDestroyed()) window.destroy();
    throw error;
  }
  mainWindow = window;
  diagnostic("window-ready", { smoke: SMOKE_TEST, analysisSmoke: ANALYSIS_SMOKE_TEST, buildSmoke: BUILD_SMOKE_TEST });
  if (SMOKE_TEST || GUIDED_SMOKE_TEST) {
    diagnostic("smoke-check-start");
    const rendererResult = GUIDED_SMOKE_TEST
      ? await window.webContents.executeJavaScript(`
          (async () => {
            const BUILD_SMOKE = ${BUILD_SMOKE_TEST ? "true" : "false"};
            const wait = (milliseconds) =>
              new Promise((resolve) => setTimeout(resolve, milliseconds));
            // Offline OCR may take longer than the UI itself.  Keep the
            // packaged analysis smoke test generous without changing normal
            // application behaviour.
            const deadline = Date.now() + 90000;
            let sampleButton = null;
            while (!sampleButton && Date.now() < deadline) {
              const discardRecovery = Array.from(document.querySelectorAll("button")).find(
                (button) => button.textContent.trim() === "Discard recovery",
              );
              discardRecovery?.click();
              const showSampleButton = Array.from(document.querySelectorAll("button")).find(
                (button) => button.textContent.trim() === "Show guided sample",
              );
              showSampleButton?.click();
              sampleButton = document.querySelector('[data-testid="guided-sample-button"]');
              if (!sampleButton) await wait(50);
            }
            // React may replace the first rendered button while startup and
            // recovery state settle. Re-query after a short settling period so
            // the smoke test never clicks a detached pre-hydration element.
            await wait(500);
            Array.from(document.querySelectorAll("button")).find(
              (button) => button.textContent.trim() === "Discard recovery",
            )?.click();
            sampleButton = document.querySelector('[data-testid="guided-sample-button"]');
            window.dispatchEvent(new Event("exhibit-builder:analyse-guided-sample"));
            let lastStartAttempt = Date.now();
            while (Date.now() < deadline) {
              // Successful analysis intentionally lands on the Review stage;
              // wait for the candidate rows before inspecting the manifest
              // table, just as a reviewer would.
              const continueButton = Array.from(document.querySelectorAll("button")).find(
                (button) => button.textContent.includes("Continue to review"),
              );
              if (continueButton) continueButton.click();
              const error = document.querySelector(".error-toast span")?.textContent;
              const candidateRows =
                document.querySelectorAll(".exhibit-review-card").length;
              const repeatRows =
                document.querySelectorAll(".repeat-reference").length;
              const reviewStatus = document.querySelector(".review-sticky-bar span")?.textContent?.trim() ?? "";
              const citedReferenceCount = Number(reviewStatus.match(/of\\s+(\\d+)\\s+statement references/i)?.[1] ?? 0);
              const referenceRows = candidateRows + repeatRows;
              const candidateLabels = Array.from(
                document.querySelectorAll(".exhibit-review-card .review-card-identity > strong"),
              ).map((label) => label.textContent.trim());
              const referenceFormat = document.querySelector(".reference-format-note > strong")?.textContent?.trim() ?? null;
              const citationChips = Array.from(document.querySelectorAll(".exhibit-review-card .review-card-chips span")).map((chip) => chip.textContent.trim());
              const guidedIdentity =
                candidateLabels.length === 5 &&
                candidateLabels.every((label) => label.length > 0 && !/^Item \\d+ - /.test(label)) &&
                citationChips.length === 5 &&
                citationChips.every((chip) => chip.startsWith("Cited at paragraph") || chip === "Not cited in the statement") &&
                citationChips.every((chip) => !/Bundle mark/.test(chip)) &&
                referenceFormat === "[AH1/page]";
              if (error || (candidateRows === 5 && citedReferenceCount === 6)) {
                const settings = document.querySelector('details[aria-label="Optional project and bundle settings"]');
                if (settings) settings.open = true;
                const prefixInput = document.querySelector('[data-testid="page-number-prefix"]');
                const suffixInput = document.querySelector('[data-testid="page-number-suffix"]');
                const setInputValue = (input, value) => {
                  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
                  setter?.call(input, value);
                  input.dispatchEvent(new Event("input", { bubbles: true }));
                  input.dispatchEvent(new Event("change", { bubbles: true }));
                };
                if (prefixInput && suffixInput) {
                  setInputValue(prefixInput, "QA-");
                  await wait(0);
                  setInputValue(suffixInput, "-END");
                  await wait(0);
                }
                const previewBeforeApply = document.querySelector(".numbering-preview span")?.textContent?.trim() ?? null;
                const warningBeforeApply = Boolean(document.querySelector('[aria-labelledby="numbering-confirmation-title"]'));
                const applyNumbering = document.querySelector('[data-testid="apply-page-numbering"]');
                const applyWasEnabled = Boolean(applyNumbering && !applyNumbering.disabled);
                applyNumbering?.click();
                let useNumbering = null;
                const numberingDialogDeadline = Date.now() + 2000;
                while (Date.now() < numberingDialogDeadline) {
                  useNumbering = Array.from(document.querySelectorAll("button")).find(
                    (button) => button.textContent.trim() === "Use this numbering",
                  );
                  if (useNumbering) break;
                  await wait(25);
                }
                const warningOnApply = Boolean(useNumbering);
                useNumbering?.click();
                const numberingClosedDeadline = Date.now() + 2000;
                while (Date.now() < numberingClosedDeadline && document.querySelector('[aria-labelledby="numbering-confirmation-title"]')) {
                  await wait(25);
                }
                const numberingInputPassed =
                  prefixInput?.value === "QA-" &&
                  suffixInput?.value === "-END" &&
                  previewBeforeApply?.includes("QA-1-END") &&
                  !warningBeforeApply &&
                  applyWasEnabled &&
                  warningOnApply &&
                  document.querySelector('[data-testid="page-number-prefix"]')?.value === "QA-" &&
                  document.querySelector('[data-testid="page-number-suffix"]')?.value === "-END" &&
                  !document.querySelector('[aria-labelledby="numbering-confirmation-title"]');
                const standardCoverInput = document.querySelector("#template-file-cover");
                const standardCoverControl = standardCoverInput?.closest(".template-control");
                const standardTemplateStatePassed = Boolean(
                  standardCoverInput &&
                  !standardCoverInput.hidden &&
                  standardCoverInput.value === "" &&
                  !standardCoverControl?.querySelector(".template-selected-name, .template-review-status") &&
                  !Array.from(standardCoverControl?.querySelectorAll("button") ?? []).some((button) => button.textContent.includes("Review")),
                );
                let templateSelectionPassed = false;
                let templatePreviewPassed = false;
                let templateResetPassed = false;
                if (standardCoverInput) {
                  const templateName = "00_GUIDED_SAMPLE_Cover_Template.pdf";
                  const templateResponse = await fetch("/guided-sample/" + encodeURIComponent(templateName), { cache: "no-store" });
                  if (templateResponse.ok) {
                    const transfer = new DataTransfer();
                    transfer.items.add(new File([await templateResponse.blob()], templateName, { type: "application/pdf" }));
                    const filesSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "files")?.set;
                    filesSetter?.call(standardCoverInput, transfer.files);
                    standardCoverInput.dispatchEvent(new Event("change", { bubbles: true }));
                    const templateSelectionDeadline = Date.now() + 10000;
                    while (!document.querySelector(".template-selected-name") && !document.querySelector(".error-toast") && Date.now() < templateSelectionDeadline) await wait(25);
                    const selectedInput = document.querySelector("#template-file-cover");
                    const selectedControl = selectedInput?.closest(".template-control");
                    const selectedActions = Array.from(selectedControl?.querySelectorAll("button") ?? []).map((button) => button.textContent.trim());
                    templateSelectionPassed = Boolean(
                      selectedInput?.hidden &&
                      selectedInput.value === "" &&
                      selectedControl?.querySelector(".template-selected-name")?.textContent.trim() === templateName &&
                      selectedControl?.querySelector(".template-review-status") &&
                      selectedActions.includes("Change") &&
                      selectedActions.includes("Review this PDF") &&
                      selectedActions.includes("Use standard design"),
                    );
                    const templateReviewButton = Array.from(selectedControl?.querySelectorAll("button") ?? []).find((button) => button.textContent.trim() === "Review this PDF");
                    templateReviewButton?.click();
                    const templatePreviewDeadline = Date.now() + 10000;
                    let templatePreviewDialog = null;
                    let templatePreviewRoot = null;
                    let templatePreviewCanvas = null;
                    let templateConfirmation = null;
                    while (Date.now() < templatePreviewDeadline) {
                      templatePreviewDialog = document.querySelector("dialog.template-preview-dialog[open]");
                      templatePreviewRoot = templatePreviewDialog?.querySelector('.original-pdf-review[data-preview-purpose="template-preview"]') ?? null;
                      templatePreviewCanvas = templatePreviewRoot?.querySelector("canvas") ?? null;
                      templateConfirmation = templatePreviewDialog?.querySelector(".template-confirmation-actions .primary-button") ?? null;
                      if (
                        templatePreviewRoot?.getAttribute("data-preview-status") === "ready" &&
                        templatePreviewCanvas?.width > 10 &&
                        templatePreviewCanvas?.height > 10 &&
                        templateConfirmation &&
                        !templateConfirmation.disabled
                      ) break;
                      if (templatePreviewRoot?.getAttribute("data-preview-status") === "error") break;
                      await wait(50);
                    }
                    const templateConfirmationWasEnabled = Boolean(templateConfirmation && !templateConfirmation.disabled);
                    templateConfirmation?.click();
                    await wait(0);
                    templatePreviewPassed = Boolean(
                      templateReviewButton &&
                      templatePreviewRoot?.getAttribute("data-preview-status") === "ready" &&
                      templatePreviewCanvas?.width > 10 &&
                      templatePreviewCanvas?.height > 10 &&
                      templateConfirmationWasEnabled &&
                      templateConfirmation?.disabled,
                    );
                    Array.from(templatePreviewDialog?.querySelectorAll("button") ?? []).find((button) => button.textContent.trim() === "Close preview")?.click();
                    const templateDialogCloseDeadline = Date.now() + 3000;
                    while (document.querySelector("dialog.template-preview-dialog[open]") && Date.now() < templateDialogCloseDeadline) await wait(25);
                    const selectedControlAfterPreview = document.querySelector("#template-file-cover")?.closest(".template-control");
                    Array.from(selectedControlAfterPreview?.querySelectorAll("button") ?? []).find((button) => button.textContent.trim() === "Use standard design")?.click();
                    const templateResetDeadline = Date.now() + 3000;
                    while ((document.querySelector(".template-selected-name") || document.querySelector(".readiness-item.template")) && Date.now() < templateResetDeadline) await wait(25);
                    const resetInput = document.querySelector("#template-file-cover");
                    const resetControl = resetInput?.closest(".template-control");
                    templateResetPassed = Boolean(
                      resetInput &&
                      !resetInput.hidden &&
                      resetInput.value === "" &&
                      !resetControl?.querySelector(".template-selected-name, .template-review-status") &&
                      !Array.from(resetControl?.querySelectorAll("button") ?? []).some((button) => button.textContent.includes("Review")) &&
                      !document.querySelector(".readiness-item.template"),
                    );
                  }
                }
                const pickers = Array.from(document.querySelectorAll("details.document-picker"));
                const secondPicker = pickers[1];
                secondPicker?.querySelector("summary")?.click();
                const pickerDeadline = Date.now() + 3000;
                while (secondPicker && !secondPicker.querySelector('.document-picker-options button[aria-pressed]') && Date.now() < pickerDeadline) {
                  await wait(50);
                }
                const pickerOptionCount = secondPicker?.querySelectorAll('.document-picker-options button[aria-pressed]').length ?? 0;
                const replacement = Array.from(secondPicker?.querySelectorAll('.document-picker-options button[aria-pressed]') ?? []).find(
                  (option) => option.textContent.includes("06_SAMPLE_Unreferenced_Checklist.pdf"),
                );
                replacement?.click();
                await wait(0);
                const labelsAfterReplacement = Array.from(
                  document.querySelectorAll(".exhibit-review-card .review-card-identity > strong"),
                ).map((label) => label.textContent.trim());
                const secondCard = document.querySelectorAll(".exhibit-review-card")[1];
                const documentPickerPassed =
                  pickers.length === 5 &&
                  pickerOptionCount === 7 &&
                  Boolean(replacement) &&
                  labelsAfterReplacement.length === 5 &&
                  labelsAfterReplacement.every((label) => label.length > 0 && !/^Item \\d+ - /.test(label)) &&
                  secondCard?.querySelector(".review-card-identity > strong")?.textContent.trim() === "Sample invoice" &&
                  (secondCard?.querySelector(".review-card-chips span")?.textContent ?? "").includes("Cited at paragraph") &&
                  (secondCard?.querySelector("details.document-picker summary span")?.textContent ?? "").includes("06_SAMPLE_Unreferenced_Checklist.pdf") &&
                  !secondCard?.querySelector("details.document-picker")?.open;
                const repeatPanel = document.querySelector(".repeat-panel.needs-decision");
                const repeatPanelButtons = Array.from(repeatPanel?.querySelectorAll("button") ?? []).map((button) => button.textContent.trim());
                const repeatControlsPresent =
                  repeatRows === 1 &&
                  repeatPanelButtons.includes("Same exhibit") &&
                  repeatPanelButtons.includes("Separate exhibit") &&
                  Boolean(repeatPanel?.querySelector(".repeat-confirm-button")) &&
                  Boolean(document.querySelector(".needs-review-status"));
                const sameExhibitButton = Array.from(repeatPanel?.querySelectorAll("button") ?? []).find(
                  (button) => button.textContent.trim().startsWith("Same exhibit"),
                );
                sameExhibitButton?.click();
                const repeatConfirmationDeadline = Date.now() + 3000;
                let repeatConfirmation = null;
                while (Date.now() < repeatConfirmationDeadline) {
                  repeatConfirmation = document.querySelector(".repeat-panel.needs-decision .repeat-confirm-button");
                  if (repeatConfirmation && !repeatConfirmation.disabled) break;
                  await wait(25);
                }
                if (repeatConfirmation && !repeatConfirmation.disabled) repeatConfirmation.click();
                const repeatDecisionDeadline = Date.now() + 3000;
                while (!document.querySelector(".repeat-panel.repeat-resolved") && Date.now() < repeatDecisionDeadline) await wait(25);
                const repeatDecisionResolved = repeatControlsPresent && Boolean(document.querySelector(".repeat-panel.repeat-resolved"));
                const repeatSelectionApplied = document.querySelector('.repeat-reference button[aria-pressed="true"]')?.textContent.trim().startsWith("Same exhibit") ?? false;
                const repeatConfirmationEnabled = Boolean(repeatConfirmation && !repeatConfirmation.disabled);
                const originalPdfButton = Array.from(document.querySelectorAll("button")).find(
                  (button) => button.textContent.trim() === "Open original PDF",
                );
                originalPdfButton?.click();
                const originalPreviewDeadline = Date.now() + 10000;
                let originalPreviewDialog = null;
                let originalPreviewCanvas = null;
                while (Date.now() < originalPreviewDeadline) {
                  originalPreviewDialog = document.querySelector("dialog.source-review-dialog[open]");
                  originalPreviewCanvas = originalPreviewDialog?.querySelector("canvas") ?? null;
                  if (
                    originalPreviewDialog?.querySelector('[data-preview-status="ready"]') &&
                    originalPreviewCanvas?.width > 10 &&
                    originalPreviewCanvas?.height > 10
                  ) break;
                  if (originalPreviewDialog?.querySelector('[data-preview-status="error"]')) break;
                  await wait(50);
                }
                const originalPdfPreviewPassed = Boolean(
                  originalPdfButton &&
                  originalPreviewDialog?.querySelector('[data-preview-status="ready"]') &&
                  originalPreviewCanvas?.width > 10 &&
                  originalPreviewCanvas?.height > 10 &&
                  !originalPreviewDialog?.querySelector(".original-pdf-error"),
                );
                Array.from(originalPreviewDialog?.querySelectorAll("button") ?? []).find(
                  (button) => button.textContent.trim() === "Close preview",
                )?.click();
                const dialogCloseDeadline = Date.now() + 3000;
                while (document.querySelector("dialog[open]") && Date.now() < dialogCloseDeadline) await wait(25);
                const confirmationDeadline = Date.now() + 45000;
                const confirmScrollSamples = [];
                let confirmScrollPassed = true;
                let confirmFocusPassed = true;
                const pendingConfirmButtons = () => Array.from(document.querySelectorAll("button.confirm-document-button"))
                  .filter((button) => button.textContent.trim() === "Confirm this document" && !button.disabled);
                let confirmationButtonCount = 0;
                while (Date.now() < confirmationDeadline && pendingConfirmButtons().length) {
                  const button = pendingConfirmButtons()[0];
                  const currentCard = button.closest("[data-candidate-id]");
                  const currentCandidateId = currentCard?.getAttribute("data-candidate-id") ?? "";
                  const cardsBeforeConfirmation = Array.from(document.querySelectorAll(".exhibit-review-card[data-candidate-id]"));
                  const currentCardIndex = cardsBeforeConfirmation.indexOf(currentCard);
                  const nextPendingCard = currentCardIndex < 0 ? null : cardsBeforeConfirmation.slice(currentCardIndex + 1).find((card) => (
                    !card.classList.contains("duplicate-source-card") &&
                    !card.classList.contains("subordinate-citation-card") &&
                    card.getAttribute("data-included") !== "false" &&
                    card.getAttribute("data-confirmed") !== "true" &&
                    card.getAttribute("data-confirmable") !== "false"
                  ));
                  const expectedNextCandidateId = nextPendingCard?.getAttribute("data-candidate-id") ?? "";
                  const unresolvedEmailStaysOpen = Boolean(
                    currentCard?.querySelector(".email-attachments-panel") &&
                    Array.from(currentCard.querySelectorAll(".email-attachment-list > li")).some((row) => !row.querySelector('.email-attachment-actions button[aria-pressed="true"]')),
                  );
                  confirmationButtonCount += 1;
                  button.focus();
                  button.scrollIntoView({ block: "center" });
                  const before = window.scrollY;
                  button.click();
                  await wait(0);
                  for (let frame = 0; frame < 6; frame += 1) {
                    await new Promise((resolve) => requestAnimationFrame(resolve));
                  }
                  const after = window.scrollY;
                  const maxAfter = Math.max(0, document.documentElement.scrollHeight - window.innerHeight);
                  const browserClamped = before > maxAfter && Math.abs(after - maxAfter) <= 1;
                  const scrollHeld = Math.abs(after - before) <= 1 || browserClamped;
                  const focused = document.activeElement;
                  const focusedCandidateId = focused?.closest("[data-candidate-id]")?.getAttribute("data-candidate-id") ?? "";
                  const currentCardStillPresent = Array.from(document.querySelectorAll("[data-candidate-id]")).some((card) => card.getAttribute("data-candidate-id") === currentCandidateId);
                  const focusHeld = unresolvedEmailStaysOpen
                    ? focusedCandidateId === currentCandidateId && Boolean(focused?.matches('[data-confirm-document][data-confirm-action="undo"]'))
                    : expectedNextCandidateId
                      ? focusedCandidateId === expectedNextCandidateId && Boolean(focused?.matches('[data-confirm-document][data-confirm-action="confirm"]:not(:disabled)'))
                      : Boolean(
                        (focusedCandidateId === currentCandidateId && focused?.matches("[data-confirm-focus]")) ||
                        (!currentCardStillPresent && focused?.matches(".pending-empty button, .exhibit-card-list")),
                      );
                  confirmScrollSamples.push({ before, after, maxAfter, browserClamped, scrollHeld, focusHeld, currentCandidateId, expectedNextCandidateId, focusedCandidateId, focusText: focused?.textContent?.trim() ?? "" });
                  if (!scrollHeld) confirmScrollPassed = false;
                  if (!focusHeld) confirmFocusPassed = false;
                }
                while (Date.now() < confirmationDeadline && !document.querySelector(".review-sticky-bar span")?.textContent.includes("6 of 6")) await wait(25);
                const emailCard = Array.from(document.querySelectorAll(".exhibit-review-card")).find((card) => card.querySelector("details.email-attachments-panel"));
                const attachmentJumpButton = Array.from(document.querySelectorAll(".readiness-controls button")).find(
                  (button) => button.textContent.trim() === "Open attachment choices on this email",
                );
                attachmentJumpButton?.scrollIntoView({ block: "center" });
                attachmentJumpButton?.click();
                for (let frame = 0; frame < 8; frame += 1) await new Promise((resolve) => requestAnimationFrame(resolve));
                const jumpedAttachmentPanel = emailCard?.querySelector(".email-attachments-panel");
                const jumpedAttachmentBounds = jumpedAttachmentPanel?.getBoundingClientRect();
                const attachmentJumpPassed = Boolean(
                  attachmentJumpButton &&
                  jumpedAttachmentPanel?.open &&
                  jumpedAttachmentPanel.contains(document.activeElement) &&
                  jumpedAttachmentBounds &&
                  jumpedAttachmentBounds.bottom > 0 &&
                  jumpedAttachmentBounds.top < window.innerHeight,
                );
                const attachmentRowCount = emailCard?.querySelectorAll(".email-attachment-list > li").length ?? 0;
                let emailChoicesMade = 0;
                for (let rowIndex = 0; rowIndex < attachmentRowCount; rowIndex += 1) {
                  const row = emailCard?.querySelectorAll(".email-attachment-list > li")[rowIndex];
                  const buttons = Array.from(row?.querySelectorAll(".email-attachment-actions button") ?? []);
                  const choice = buttons.find((button) => button.textContent.trim() === "Print with this email")
                    ?? buttons.find((button) => button.textContent.trim() === "Leave out");
                  if (choice && choice.getAttribute("aria-pressed") !== "true") {
                    choice.click();
                    emailChoicesMade += 1;
                    await wait(0);
                  }
                }
                for (let frame = 0; frame < 4; frame += 1) await new Promise((resolve) => requestAnimationFrame(resolve));
                const emailCompact = Boolean(emailCard?.classList.contains("compact"));
                const emailMinimise = Boolean(Array.from(emailCard?.querySelectorAll("button") ?? []).some((button) => button.textContent.trim() === "Minimise"));
                const emailOpenDetails = Array.from(emailCard?.querySelectorAll("button") ?? []).find((button) => button.textContent.trim() === "View or change");
                emailOpenDetails?.click();
                await wait(0);
                emailCard?.querySelector(".email-attachments-panel > summary")?.click();
                await wait(0);
                const emailAttachmentsOpenBeforeMinimise = Boolean(emailCard?.querySelector(".email-attachments-panel")?.open);
                const emailMinimiseButton = Array.from(emailCard?.querySelectorAll("button") ?? []).find((button) => button.textContent.trim() === "Minimise");
                const emailMinimiseAfterReview = Boolean(emailMinimiseButton);
                const emailDecisionsRetained = () => {
                  const rows = Array.from(emailCard?.querySelectorAll(".email-attachment-list > li") ?? []);
                  return rows.length > 0 && rows.every((row) => row.querySelector('button[aria-pressed="true"]'));
                };
                const emailDecisionsBeforeMinimise = emailDecisionsRetained();
                emailMinimiseButton?.click();
                await wait(0);
                const emailCompactAfterMinimise = Boolean(emailCard?.classList.contains("compact"));
                Array.from(emailCard?.querySelectorAll("button") ?? []).find((button) => button.textContent.trim() === "View or change")?.click();
                await wait(0);
                const emailAttachmentsClosedAfterMinimise = !emailCard?.querySelector(".email-attachments-panel")?.open;
                const emailDecisionsAfterMinimise = emailDecisionsRetained();
                Array.from(emailCard?.querySelectorAll("button") ?? []).find((button) => button.textContent.trim() === "Undo confirmation")?.click();
                await wait(0);
                const emailDecisionsAfterUndo = emailDecisionsRetained();
                Array.from(emailCard?.querySelectorAll("button") ?? []).find((button) => button.textContent.trim() === "Confirm this document")?.click();
                await wait(0);
                for (let frame = 0; frame < 4; frame += 1) await new Promise((resolve) => requestAnimationFrame(resolve));
                const emailCompactAfterAttachmentsFirstConfirm = Boolean(emailCard?.classList.contains("compact"));
                const compactCards = Array.from(document.querySelectorAll(".exhibit-review-card.compact"));
                const confirmationWordingPassed = compactCards.length > 0 && compactCards.every((card) => {
                  const status = card.querySelector(".review-card-status")?.textContent ?? "";
                  const openDetails = Array.from(card.querySelectorAll("button")).some((button) => button.textContent.trim() === "View or change");
                  return status.includes("Confirmed") && !/\bNeeds review\b/.test(status) && !status.includes("Possible suggested match") && openDetails;
                });
                const reviewStatusAfterConfirmation = document.querySelector(".review-sticky-bar span")?.textContent?.trim() ?? "";
                const progressStrip = document.querySelector(".review-progress-strip")?.textContent ?? "";
                const repeatProgressPassed = /repeat exhibit/.test(reviewStatusAfterConfirmation) && /repeat exhibit/.test(progressStrip);
                const continueFromReview = Array.from(document.querySelectorAll("button")).find(
                  (button) => button.textContent.trim() === "Continue to workbook sheets" || button.textContent.trim() === "Continue to finalise",
                );
                // Template review is an independent build gate. The repeat-path
                // smoke is complete when all six statement references are confirmed
                // and the repeat resolves.
                const repeatReviewPassed = repeatDecisionResolved && reviewStatusAfterConfirmation.includes("6 of 6");
                let buildPassed = !BUILD_SMOKE;
                let buildHeading = null;
                let buildSha = null;
                let buildPageCount = 0;
                let optionalPagesSummary = null;
                let buildSaved = false;
                let buildError = null;
                if (BUILD_SMOKE) {
                  const continueButton = Array.from(document.querySelectorAll("button")).find(
                    (button) => button.textContent.trim() === "Continue to workbook sheets" || button.textContent.trim() === "Continue to finalise",
                  );
                  continueButton?.click();
                  const sheetsDeadline = Date.now() + 20000;
                  while (Date.now() < sheetsDeadline && !document.querySelector("main h1")?.textContent.includes("Choose which Excel sheets")) await wait(50);
                  Array.from(document.querySelectorAll("button")).find((button) => button.textContent.trim() === "Continue to finalise")?.click();
                  const finaliseDeadline = Date.now() + 20000;
                  while (Date.now() < finaliseDeadline && !Array.from(document.querySelectorAll("button")).some((button) => button.textContent.trim() === "Build exhibit bundle")) await wait(50);
                  Array.from(document.querySelectorAll("button")).find((button) => button.textContent.trim() === "Build exhibit bundle")?.click();
                  const buildDeadline = Date.now() + 600000;
                  while (Date.now() < buildDeadline) {
                    buildError = document.querySelector(".error-toast span")?.textContent ?? null;
                    buildHeading = document.querySelector("main h1")?.textContent?.trim() ?? null;
                    if (buildHeading === "Your bundle is ready" || buildError) break;
                    await wait(250);
                  }
                  buildSha = document.querySelector(".hash-line code")?.textContent?.trim() ?? null;
                  const summary = document.querySelector('[data-testid="bundle-output-summary"]')?.textContent ?? "";
                  buildPageCount = Number(summary.match(/(\\d+)\\s+(?:total\\s+)?pages/)?.[1] ?? 0);
                  optionalPagesSummary = Array.from(document.querySelectorAll(".output-metrics span")).find((item) => /optional pages included/.test(item.textContent))?.textContent?.trim() ?? null;
                  const downloadPdf = Array.from(document.querySelectorAll("button")).find((button) => button.textContent.trim() === "Download bundle PDF");
                  downloadPdf?.click();
                  await wait(3000);
                  buildSaved = Boolean(downloadPdf);
                  buildPassed = Boolean(
                    continueButton &&
                    buildHeading === "Your bundle is ready" &&
                    document.body.textContent.includes("The finished PDF has been reopened and checked.") &&
                    buildSha &&
                    buildPageCount >= 8 &&
                    optionalPagesSummary === "No optional pages included in exhibit references" &&
                    buildSaved &&
                    !buildError,
                  );
                }
                return {
                  title: document.title,
                  hasWorkspace: Boolean(document.querySelector("main")),
                  remoteBlocked: await Promise.race([
                    fetch("https://example.com").then(() => false).catch(() => true),
                    wait(2000).then(() => true),
                  ]),
                  analysisPassed: candidateRows === 5 && citedReferenceCount === 6 && guidedIdentity && numberingInputPassed && standardTemplateStatePassed && templateSelectionPassed && templateResetPassed && documentPickerPassed && repeatReviewPassed && originalPdfPreviewPassed && templatePreviewPassed && confirmScrollPassed && confirmFocusPassed && attachmentJumpPassed && confirmationWordingPassed && attachmentRowCount > 0 && emailChoicesMade === attachmentRowCount && emailCompact && emailAttachmentsOpenBeforeMinimise && emailMinimiseAfterReview && emailDecisionsBeforeMinimise && emailCompactAfterMinimise && emailAttachmentsClosedAfterMinimise && emailDecisionsAfterMinimise && emailDecisionsAfterUndo && emailCompactAfterAttachmentsFirstConfirm && repeatProgressPassed && buildPassed && !error,
                  candidateRows,
                  repeatRows,
                  referenceRows,
                  citedReferenceCount,
                  reviewStatus,
                  candidateLabels,
                  referenceFormat,
                  guidedIdentity,
                  prefixValue: prefixInput?.value ?? null,
                  suffixValue: suffixInput?.value ?? null,
                  previewBeforeApply,
                  warningBeforeApply,
                  warningOnApply,
                  numberingInputPassed,
                  standardTemplateStatePassed,
                  templateSelectionPassed,
                  templateResetPassed,
                  pickerCount: pickers.length,
                  pickerOptionCount,
                  labelsAfterReplacement,
                  documentPickerPassed,
                  repeatReviewPassed,
                  repeatControlsPresent,
                  repeatSelectionApplied,
                  repeatConfirmationEnabled,
                  repeatDecisionResolved,
                  continueFromReviewEnabled: Boolean(continueFromReview && !continueFromReview.disabled),
                  confirmationButtonCount,
                  reviewStatusAfterConfirmation,
                  originalPdfPreviewPassed,
                  templatePreviewPassed,
                  confirmScrollPassed,
                  confirmFocusPassed,
                  confirmScrollSamples,
                  repeatProgressPassed,
                  attachmentJumpPassed,
                  confirmationWordingPassed,
                  emailCompact,
                  attachmentRowCount,
                  emailChoicesMade,
                  emailMinimise,
                  emailAttachmentsOpenBeforeMinimise,
                  emailMinimiseAfterReview,
                  emailDecisionsBeforeMinimise,
                  emailCompactAfterMinimise,
                  emailAttachmentsClosedAfterMinimise,
                  emailDecisionsAfterMinimise,
                  emailDecisionsAfterUndo,
                  emailCompactAfterAttachmentsFirstConfirm,
                  buildPassed,
                  buildHeading,
                  buildSha,
                  buildPageCount,
                  optionalPagesSummary,
                  buildSaved,
                  buildError,
                  error: error || buildError || null,
                };
              }
              const currentSampleButton = document.querySelector('[data-testid="guided-sample-button"]');
              if (
                currentSampleButton &&
                !currentSampleButton.disabled &&
                Date.now() - lastStartAttempt >= 1000
              ) {
                window.dispatchEvent(new Event("exhibit-builder:analyse-guided-sample"));
                lastStartAttempt = Date.now();
              }
              await wait(100);
            }
            return {
              title: document.title,
              hasWorkspace: Boolean(document.querySelector("main")),
              remoteBlocked: true,
              analysisPassed: false,
              candidateRows: document.querySelectorAll(".exhibit-review-card").length,
              citedReferenceCount: Number((document.querySelector(".review-sticky-bar span")?.textContent ?? "").match(/of\\s+(\\d+)\\s+statement references/i)?.[1] ?? 0),
              reviewStatus: document.querySelector(".review-sticky-bar span")?.textContent?.trim() ?? null,
              sampleButtonPresent: Boolean(document.querySelector('[data-testid="guided-sample-button"]')),
              sampleButtonText: document.querySelector('[data-testid="guided-sample-button"]')?.textContent?.trim() ?? null,
              sampleButtonDisabled: document.querySelector('[data-testid="guided-sample-button"]')?.disabled ?? null,
              visibleHeading: document.querySelector("main h1")?.textContent?.trim() ?? null,
              rendererError: document.querySelector(".error-toast span")?.textContent ?? null,
              error: "Timed out while analysing the sample pack after 90 seconds",
            };
          })()
        `, true)
      : await window.webContents.executeJavaScript(`
          (async () => {
            const deadline = Date.now() + 30000;
            while (!document.querySelector("main") && Date.now() < deadline) {
              await new Promise((resolve) => setTimeout(resolve, 50));
            }
            return {
              title: document.title,
              hasWorkspace: Boolean(document.querySelector("main")),
              remoteBlocked: await Promise.race([
                fetch("https://example.com").then(() => false).catch(() => true),
                new Promise((resolve) => setTimeout(() => resolve(true), 2000)),
              ]),
            };
          })()
        `, true);
    const windowState = windowPresentationState(window);
    const result = { ...rendererResult, ...windowState };
    const windowPresentationPassed = windowState.windowVisible && windowState.windowMaximized && !windowState.windowFullScreen && !windowState.windowMinimized;
    const passed =
      result.title === "Exhibit Builder | Offline Desktop" &&
      result.hasWorkspace &&
      result.remoteBlocked &&
      windowPresentationPassed &&
      (!GUIDED_SMOKE_TEST || result.analysisPassed) &&
      (!BUILD_SMOKE_TEST || result.buildPassed);
    if (process.env.BUNDLE_BUILDER_SMOKE_LOG) {
      await writeFile(
        process.env.BUNDLE_BUILDER_SMOKE_LOG,
        JSON.stringify({ passed, ...result }, null, 2),
      );
    }
    console.log(passed ? "DESKTOP_SMOKE_OK" : "DESKTOP_SMOKE_FAILED", result);
    diagnostic("smoke-check-finished", { passed, hasWorkspace: result.hasWorkspace, analysisPassed: result.analysisPassed, buildPassed: result.buildPassed });
    // Make the smoke contract meaningful to CI and packaging checks. A
    // failed renderer/load check must not look like a successful process exit.
    process.exitCode = passed ? 0 : 1;
    if (!window.isDestroyed()) window.destroy();
    localServer?.close();
    // Smoke checks need a trustworthy process exit code. Perform the normal
    // synchronous session cleanup explicitly, then exit with that result.
    cleanupSession();
    app.exit(process.exitCode);
    return;
  }
}

function loadWorkspace(window) {
  return new Promise((resolve, reject) => {
    let settled = false;
    let timeout;

    const finish = (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      window.webContents.removeListener("render-process-gone", onRendererGone);
      window.webContents.removeListener("did-fail-load", onLoadFailed);
      if (error) reject(error);
      else resolve();
    };

    const onRendererGone = (_event, details) => {
      if (details.reason === "clean-exit") return;
      finish(new Error(`The workspace renderer could not start (${details.reason || "unknown"}, exit ${details.exitCode ?? "n/a"}).`));
    };

    const onLoadFailed = (_event, errorCode, errorDescription, _validatedURL, isMainFrame) => {
      if (isMainFrame && errorCode !== -3) {
        finish(new Error(`The offline workspace could not load (${errorCode}: ${errorDescription}).`));
      }
    };

    timeout = setTimeout(() => {
      diagnostic("window-load-timeout", { timeoutMs: WORKSPACE_LOAD_TIMEOUT_MS });
      finish(new Error(`The offline workspace did not load within ${WORKSPACE_LOAD_TIMEOUT_MS}ms.`));
    }, WORKSPACE_LOAD_TIMEOUT_MS);

    window.webContents.once("render-process-gone", onRendererGone);
    window.webContents.once("did-fail-load", onLoadFailed);
    window.loadURL(`${localOrigin}/`).then(() => finish()).catch(finish);
  });
}

ipcMain.handle(
  "bundle-builder:save-file",
  async (event, payload) => {
    assertTrustedIpcSender(event, "bundle-builder:save-file");
    const bytes = payload?.bytes;
    const fileName = payload?.fileName;
    const mediaType = payload?.mediaType;
    const protectedSourcePaths = parseProtectedSourcePaths(payload?.protectedSourcePaths);
    const allowedOverwritePath = parseOptionalAbsolutePath(payload?.allowedOverwritePath, "Allowed overwrite path");
    const extension = path.extname(fileName).slice(1).toUpperCase() || "FILE";
    let destination;
    if (BUILD_SMOKE_TEST) {
      const smokeDir = process.env.EXHIBIT_BUILDER_SMOKE_SAVE_DIR;
      if (!smokeDir || typeof smokeDir !== "string") throw new Error("Build smoke test requires EXHIBIT_BUILDER_SMOKE_SAVE_DIR.");
      const root = path.resolve(smokeDir);
      await mkdir(root, { recursive: true });
      destination = path.join(root, path.basename(String(fileName || "bundle.pdf")));
      if (path.dirname(destination) !== root) throw new Error("Build smoke save path escaped the smoke directory.");
    } else {
      const result = await dialog.showSaveDialog({
        title: "Save Exhibit Builder output",
        defaultPath: fileName,
        filters: [
          {
            name: mediaType || extension,
            extensions: [extension.toLowerCase()],
          },
        ],
      });
      if (result.canceled || !result.filePath) return { saved: false };
      destination = result.filePath;
    }
    try {
      await assertDestinationAllowed(destination, protectedSourcePaths, { fileName: String(fileName || ""), allowedOverwritePath });
    } catch (error) {
      if (error instanceof SourcePathCollisionError || error?.code === "SOURCE_PATH_COLLISION") {
        if (BUILD_SMOKE_TEST) throw error;
        await dialog.showMessageBox({
          type: "warning",
          title: "Source file not replaced",
          message: error.message,
        });
        return { saved: false };
      }
      throw error;
    }
    const temporary = path.join(path.dirname(destination), `.${path.basename(destination)}.${randomUUID()}.tmp`);
    try {
      await writeFile(temporary, Buffer.from(bytes), { flag: "wx" });
      await rename(temporary, destination);
    } finally {
      await rm(temporary, { force: true }).catch(() => {});
    }
    return { saved: true, filePath: destination };
  },
);

ipcMain.handle("bundle-builder:clipboard-write", async (event, text) => {
  assertTrustedIpcSender(event, "bundle-builder:clipboard-write");
  if (typeof text !== "string") throw new TypeError("Clipboard text must be a string.");
  if (text.length > 1_000_000) throw new Error("The text is too long to copy in one step. Download the .txt file instead.");
  clipboard.writeText(text);
  return { copied: true };
});

ipcMain.handle("bundle-builder:export-workbook", async (event, request) => {
  assertTrustedIpcSender(event, "bundle-builder:export-workbook");
  return exportWorkbookSheets(SESSION_ROOT, request);
});

ipcMain.handle("bundle-builder:recovery-status", async (event) => {
  assertTrustedIpcSender(event, "bundle-builder:recovery-status");
  return recoveryStore.status();
});

ipcMain.handle("bundle-builder:preferences-read", async (event) => {
  assertTrustedIpcSender(event, "bundle-builder:preferences-read");
  return preferenceStore.read();
});

ipcMain.handle("bundle-builder:guided-sample-hidden", async (event, hidden) => {
  assertTrustedIpcSender(event, "bundle-builder:guided-sample-hidden");
  if (typeof hidden !== "boolean") throw new TypeError("The guided-sample preference must be true or false.");
  return preferenceStore.write({ hideGuidedSample: hidden });
});

ipcMain.handle("bundle-builder:open-guided-sample-folder", async (event) => {
  assertTrustedIpcSender(event, "bundle-builder:open-guided-sample-folder");
  const directory = resolveGuidedSampleDirectory();
  if (!directory) throw new Error("The guided sample folder is not available.");
  const opened = await shell.openPath(directory);
  if (opened) throw new Error("The guided sample folder could not be opened.");
  return { opened: true };
});

ipcMain.handle("bundle-builder:recovery-begin", async (event) => {
  assertTrustedIpcSender(event, "bundle-builder:recovery-begin");
  return recoveryStore.begin();
});

ipcMain.handle("bundle-builder:recovery-write", async (event, { recoveryId, revision, payload }) => {
  assertTrustedIpcSender(event, "bundle-builder:recovery-write");
  return recoveryStore.write(recoveryId, revision, payload);
});

ipcMain.handle("bundle-builder:recovery-load", async (event, { recoveryId }) => {
  assertTrustedIpcSender(event, "bundle-builder:recovery-load");
  return recoveryStore.load(recoveryId);
});

ipcMain.handle("bundle-builder:recovery-read-source", async (event, { recoveryId, sourceId }) => {
  assertTrustedIpcSender(event, "bundle-builder:recovery-read-source");
  return recoveryStore.readSource(recoveryId, sourceId);
});

ipcMain.handle("bundle-builder:recovery-discard", async (event, { recoveryId }) => {
  assertTrustedIpcSender(event, "bundle-builder:recovery-discard");
  return recoveryStore.discard(recoveryId);
});

ipcMain.handle("bundle-builder:recovery-clear-all", async (event) => {
  assertTrustedIpcSender(event, "bundle-builder:recovery-clear-all");
  return recoveryStore.clearAll();
});

ipcMain.handle("bundle-builder:recovery-mark-clean", async (event, { recoveryId, revision, savedArchive }) => {
  assertTrustedIpcSender(event, "bundle-builder:recovery-mark-clean");
  return recoveryStore.markClean(recoveryId, revision, savedArchive ?? null);
});

ipcMain.handle(
  "bundle-builder:convert-template",
  async (event, { html, sourceName }) => {
    assertTrustedIpcSender(event, "bundle-builder:convert-template");
    if (typeof html !== "string" || html.length === 0 || html.length > 20_000_000) {
      throw new Error("The Word template preview was empty or too large to convert locally.");
    }
    const preview = new BrowserWindow({
      show: false,
      width: 794,
      height: 1123,
      backgroundColor: "#ffffff",
      webPreferences: {
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
        spellcheck: false,
        partition: APP_PARTITION,
      },
    });
    try {
      preview.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
      preview.webContents.on("will-navigate", (event, url) => {
        if (!url.startsWith("data:")) event.preventDefault();
      });
      const dataUrl = `data:text/html;charset=utf-8,${encodeURIComponent(html)}`;
      let conversionTimer;
      const pdf = await Promise.race([
        (async () => {
          await preview.loadURL(dataUrl);
          await preview.webContents.executeJavaScript(
            `Promise.race([document.fonts?.ready, new Promise((resolve) => setTimeout(resolve, 300))])`,
          );
          return preview.webContents.printToPDF({
            pageSize: "A4",
            printBackground: true,
            preferCSSPageSize: true,
            margins: { marginType: "none" },
          });
        })(),
        new Promise((_resolve, reject) => { conversionTimer = setTimeout(() => reject(new Error("Template conversion exceeded the two-minute safety limit.")), 2 * 60 * 1000); }),
      ]).finally(() => clearTimeout(conversionTimer));
      return new Uint8Array(pdf);
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      throw new Error(`The Word template ${sourceName || "file"} could not be converted offline: ${detail}`);
    } finally {
      if (!preview.isDestroyed()) preview.destroy();
    }
  },
);

app.whenReady().then(async () => {
  diagnostic("app-ready");
  app.setAppUserModelId("com.exhibitbuilder.offline");
  await startLocalServer();
  const appSession = session.fromPartition(APP_PARTITION);
  installNetworkBlock(appSession);
  await createWindow();
}).catch((error) => {
  diagnostic("startup-failed", { message: error instanceof Error ? error.message : String(error) });
  console.error("Exhibit Builder failed during startup.", error);
  process.exitCode = 1;
  cleanupSession();
  dialog.showErrorBox("Exhibit Builder could not start", "The offline workspace could not be opened. Close and retry the application.");
  app.quit();
});

app.on("window-all-closed", () => {
  diagnostic("window-all-closed");
  localServer?.close();
  app.quit();
});

app.on("before-quit", (event) => {
  if (exportShutdownStarted) return;
  event.preventDefault();
  exportShutdownStarted = true;
  void Promise.race([
    shutdownAllExports(),
    new Promise((resolve) => setTimeout(resolve, 20_000)),
  ]).finally(() => app.quit());
});

app.on("child-process-gone", (_event, details) => {
  if (details.reason !== "clean-exit") {
    console.error(`Exhibit Builder child process exited: ${details.type || "unknown"} / ${details.reason || "unknown"} (${details.exitCode ?? "n/a"}).`);
  }
});

app.on("will-quit", () => {
  diagnostic("will-quit");
  localServer?.close();
  cleanupSession();
});
