const { createHash } = require("node:crypto");
const { existsSync, readFileSync, rmSync, writeFileSync } = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const { assertExactFile, assertNonEmptyFile } = require("./release-verification-helpers.cjs");

function fail(message) {
  throw new Error(`Installed release verification failed: ${message}`);
}

function run(file, args, timeout, environment = process.env) {
  const result = spawnSync(file, args, { env: environment, encoding: "utf8", timeout, windowsHide: true });
  if (result.error) fail(`${path.basename(file)} could not run: ${result.error.message}`);
  if (result.status !== 0) fail(`${path.basename(file)} exited with code ${result.status}. ${result.stderr || result.stdout || ""}`.trim());
  return result;
}

function safeRemove(directory, allowedRoot) {
  const resolved = path.resolve(directory);
  const root = `${path.resolve(allowedRoot)}${path.sep}`.toLowerCase();
  if (!`${resolved}${path.sep}`.toLowerCase().startsWith(root)) fail("temporary installation path escaped the system temporary directory");
  rmSync(resolved, { recursive: true, force: true, maxRetries: 3, retryDelay: 150 });
}

const installerPath = path.resolve(process.argv[2] || "");
const version = process.argv[3] || "";
const reportPath = path.resolve(process.argv[4] || path.join(process.cwd(), "release", `installed-verification-${version}.json`));
if (!version || !existsSync(installerPath)) fail("provide an existing installer path and version");

const temporaryRoot = path.resolve(os.tmpdir());
const sourceRoot = path.resolve(__dirname, "..");
const installRoot = path.join(temporaryRoot, `ExhibitBuilderReleaseVerify-${version}-${process.pid}`);
if (existsSync(installRoot)) fail("the temporary verification directory already exists");
const diagnosticLog = path.join(temporaryRoot, `ExhibitBuilderReleaseVerify-${version}-${process.pid}.log`);
const checks = [];
let installed = false;

try {
  run(installerPath, ["/S", "/RELEASEVERIFY", `/D=${installRoot}`], 10 * 60 * 1000);
  installed = true;
  checks.push("temporary installation completed");

  const executable = path.join(installRoot, "Exhibit Builder.exe");
  const uninstaller = path.join(installRoot, "Uninstall Exhibit Builder.exe");
  const eula = path.join(installRoot, "EULA.txt");
  const thirdPartyLicences = path.join(installRoot, "THIRD_PARTY_LICENSES.txt");
  const appArchive = path.join(installRoot, "resources", "app.asar");
  const electronLicence = path.join(installRoot, "LICENSE.electron.txt");
  const chromiumLicences = path.join(installRoot, "LICENSES.chromium.html");
  for (const required of [executable, uninstaller, eula, thirdPartyLicences, electronLicence, chromiumLicences, appArchive, path.join(installRoot, ".exhibit-builder-verification-install")]) {
    if (!existsSync(required)) fail(`installed file is missing: ${path.relative(installRoot, required)}`);
  }
  const eulaText = assertExactFile(eula, path.join(sourceRoot, "build", "EULA.txt"), "the installed EULA");
  if (!/EXHIBIT BUILDER FREEWARE END USER LICENCE AGREEMENT/i.test(eulaText)) fail("the installed EULA is not the approved Exhibit Builder licence");
  const thirdPartyText = assertExactFile(thirdPartyLicences, path.join(sourceRoot, "THIRD_PARTY_LICENSES.txt"), "the installed third-party licence notices");
  if (!/THIRD-PARTY LICENCE NOTICES/i.test(thirdPartyText) || !/tesseract/i.test(thirdPartyText)) fail("the installed third-party licence notices are incomplete");
  const electronRuntimeNotices = {
    electron: assertNonEmptyFile(electronLicence, "the installed Electron licence"),
    chromium: assertNonEmptyFile(chromiumLicences, "the installed Chromium notices"),
  };
  checks.push("installed application, EULA, project notices, Electron/Chromium runtime notices and packaged archive present");

  const smokeEnvironment = { ...process.env, EXHIBIT_BUILDER_DIAGNOSTIC_LOG: diagnosticLog };
  run(executable, ["--smoke-test"], 3 * 60 * 1000, smokeEnvironment);
  checks.push("installed shell smoke test passed");
  run(executable, ["--analysis-smoke-test"], 10 * 60 * 1000, smokeEnvironment);
  checks.push("installed guided-sample analysis smoke test passed");

  run(uninstaller, ["/S"], 5 * 60 * 1000);
  const deadline = Date.now() + 10_000;
  while (existsSync(installRoot) && Date.now() < deadline) Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 100);
  if (existsSync(installRoot)) fail("the verification installation was not removed by its uninstaller");
  installed = false;
  checks.push("verification uninstall completed without touching normal user installation data");

  const installerBytes = readFileSync(installerPath);
  const report = {
    product: "Exhibit Builder",
    version,
    verifiedAt: new Date().toISOString(),
    installer: { fileName: path.basename(installerPath), sha256: createHash("sha256").update(installerBytes).digest("hex"), size: installerBytes.byteLength },
    electronRuntimeNotices,
    checks,
    result: "pass",
  };
  writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`, { flag: "w" });
  process.stdout.write(`Installed release verification passed. Report: ${reportPath}\n`);
} finally {
  if (installed && existsSync(installRoot)) safeRemove(installRoot, temporaryRoot);
  if (existsSync(diagnosticLog)) rmSync(diagnosticLog, { force: true });
}
