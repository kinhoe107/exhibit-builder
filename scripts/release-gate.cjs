const { existsSync, lstatSync, readFileSync, readdirSync, rmSync, symlinkSync } = require("node:fs");
const { createHash } = require("node:crypto");
const os = require("node:os");
const path = require("node:path");
const { spawnSync } = require("node:child_process");

const root = path.resolve(__dirname, "..");
const packageJson = JSON.parse(readFileSync(path.join(root, "package.json"), "utf8"));
const version = packageJson.version;
const node = process.execPath;

function run(label, file, args, timeout, environment = process.env, workingDirectory = root) {
  process.stdout.write(`\n[release gate] ${label}\n`);
  const result = spawnSync(file, args, { cwd: workingDirectory, env: environment, encoding: "utf8", stdio: "inherit", timeout, windowsHide: true });
  if (result.error) throw new Error(`${label} could not run: ${result.error.message}`);
  if (result.status !== 0) throw new Error(`${label} failed with exit code ${result.status}. Release promotion is blocked.`);
}

function findFile(directory, targetName, predicate = () => true) {
  if (!existsSync(directory)) return null;
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const candidate = path.join(directory, entry.name);
    if (entry.isFile() && entry.name.toLowerCase() === targetName.toLowerCase() && predicate(candidate)) return candidate;
    if (entry.isDirectory()) {
      const nested = findFile(candidate, targetName, predicate);
      if (nested) return nested;
    }
  }
  return null;
}

function pinnedNsisCompiler() {
  const version = "nsis-3.0.4.1";
  const expectedSha256 = "e277b7378931b74392015f5ad6b1d744dcd8a347baa4480350a75ebeab8d8e3d";
  const versionRoot = path.join(process.env.LOCALAPPDATA || "", "electron-builder", "Cache", version);
  const compiler = findFile(versionRoot, "makensis.exe", (candidate) => candidate.toLowerCase().includes(`${path.sep}bin${path.sep}`));
  if (!compiler) throw new Error(`Pinned NSIS ${version} is not installed. Provision that exact electron-builder toolchain before running the release gate. Release promotion is blocked.`);
  const actual = createHash("sha256").update(readFileSync(compiler)).digest("hex");
  if (actual !== expectedSha256) throw new Error(`Pinned NSIS compiler hash mismatch: expected ${expectedSha256}, received ${actual}. Release promotion is blocked.`);
  return compiler;
}

const tests = readdirSync(path.join(root, "tests"))
  .filter((name) => name.endsWith(".test.mjs") || name === "run-guided-sample.mjs")
  .map((name) => path.join("tests", name));

// Check the exact installer prerequisite before starting any lengthy release
// work. Deep Windows paths cannot reliably run electron-builder's generated
// NSIS script, so the final assembly uses the audited short-path wrapper.
const makensis = pinnedNsisCompiler();

run("third-party licence notice generation", node, [path.join("scripts", "generate-third-party-notices.cjs")], 60 * 1000);
for (const directory of ["electron", "scripts"]) {
  for (const name of readdirSync(path.join(root, directory)).filter((entry) => entry.endsWith(".cjs"))) {
    run(`Node syntax: ${directory}/${name}`, node, ["--check", path.join(directory, name)], 60 * 1000);
  }
}
run("strict TypeScript compilation", node, [path.join("node_modules", "typescript", "bin", "tsc"), "--noEmit"], 5 * 60 * 1000);
run("TypeScript-aware lint", node, [path.join("node_modules", "eslint", "bin", "eslint.js"), "app", "src", "vite.config.ts"], 5 * 60 * 1000);
run("critical bundle-logic coverage", node, [
  "--experimental-strip-types",
  "--experimental-test-coverage",
  "--test-coverage-lines=95",
  "--test-coverage-branches=80",
  "--test-coverage-functions=95",
  "--test-coverage-include=app/lib/bundle-arrangement.ts",
  "--test-coverage-include=app/lib/build-plan.ts",
  "--test-coverage-include=app/lib/index-layout.ts",
  "--test-coverage-include=app/lib/template-persistence.ts",
  "--test",
  path.join("tests", "bundle-arrangement.test.mjs"),
  path.join("tests", "build-plan.test.mjs"),
  path.join("tests", "index-layout.test.mjs"),
  path.join("tests", "template-persistence.test.mjs"),
], 10 * 60 * 1000);
run("legal-integrity coverage", node, [
  "--experimental-strip-types",
  "--experimental-test-coverage",
  "--test-coverage-lines=80",
  "--test-coverage-branches=70",
  "--test-coverage-functions=80",
  "--test-coverage-include=app/lib/preflight.ts",
  "--test-coverage-include=app/lib/project-archive.ts",
  "--test-coverage-include=app/lib/xlsx.ts",
  "--test-coverage-include=app/lib/exhibit-groups.ts",
  "--test",
  path.join("tests", "project-features.test.mjs"),
  path.join("tests", "xlsx.test.mjs"),
  path.join("tests", "candidate-review-change.test.mjs"),
  path.join("tests", "engine.test.mjs"),
], 10 * 60 * 1000);
run("source build", node, [path.join("node_modules", "vite", "bin", "vite.js"), "build"], 5 * 60 * 1000);
run("complete automated test suite", node, ["--experimental-strip-types", "--test", ...tests], 30 * 60 * 1000);
run("native Microsoft Excel guided-sample fidelity test", node, ["--experimental-strip-types", "--test", path.join("tests", "run-guided-sample.mjs")], 15 * 60 * 1000, { ...process.env, EXHIBIT_BUILDER_NATIVE_EXCEL: "1" });
const privateStressPack = path.resolve(root, "..", "Sample_Pack_ICC_Adversarial");
if (!existsSync(privateStressPack)) {
  throw new Error("The required private ICC adversarial pack is missing. Release promotion is blocked.");
}
run("private ICC adversarial fidelity test", node, ["--experimental-strip-types", "--test", path.join("tests", "run-adversarial-pack.mjs")], 30 * 60 * 1000, { ...process.env, EXHIBIT_BUILDER_NATIVE_EXCEL: "1" });

const unpackedExecutable = path.join(root, "release", "win-unpacked", "Exhibit Builder.exe");
const installer = path.join(root, "release", `Exhibit-Builder-${version}-setup.exe`);
run("Windows unpacked application packaging", node, [path.join("node_modules", "electron-builder", "cli.js"), "--win", "dir", "--x64"], 30 * 60 * 1000);
if (!existsSync(unpackedExecutable)) throw new Error("Packaging did not produce the unpacked application. Release promotion is blocked.");
const shortRoot = path.join(os.tmpdir(), `EBRelease-${process.pid}`);
if (existsSync(shortRoot)) throw new Error("The short release-build path already exists. Release promotion is blocked.");
symlinkSync(root, shortRoot, "junction");
try {
  rmSync(installer, { force: true });
  run("audited installer assembly", makensis, [
    `/DAPP_VERSION=${version}`,
    `/DSOURCE_DIR=${path.join(shortRoot, "release", "win-unpacked")}`,
    `/DOUT_FILE=${path.join(shortRoot, "release", `Exhibit-Builder-${version}-setup.exe`)}`,
    path.join(shortRoot, "build", "standalone-installer.nsi"),
  ], 30 * 60 * 1000);
} finally {
  if (existsSync(shortRoot)) {
    if (!lstatSync(shortRoot).isSymbolicLink()) throw new Error("The temporary short path is not a junction; it was not removed.");
    rmSync(shortRoot, { recursive: true, force: true });
  }
}
if (!existsSync(installer)) throw new Error("Audited installer assembly did not produce the release installer. Release promotion is blocked.");
const smokeLog = path.join(root, "release", `unpacked-smoke-${version}.log`);
const smokeSave = path.join(os.tmpdir(), `ExhibitBuilderUnpackedSmoke-save-${process.pid}`);
const smokeEnvironment = { ...process.env, EXHIBIT_BUILDER_DIAGNOSTIC_LOG: smokeLog };
run("unpacked shell smoke test", unpackedExecutable, ["--smoke-test"], 3 * 60 * 1000, smokeEnvironment);
run("unpacked guided-sample finished-bundle smoke test", unpackedExecutable, ["--build-smoke-test"], 15 * 60 * 1000, { ...smokeEnvironment, EXHIBIT_BUILDER_SMOKE_SAVE_DIR: smokeSave });
if (!existsSync(smokeSave) || !readdirSync(smokeSave).some((name) => /\.pdf$/i.test(name))) {
  throw new Error("Unpacked finished-bundle smoke test did not save a PDF. Release promotion is blocked.");
}
rmSync(smokeSave, { recursive: true, force: true });
run("temporary installation, installed smoke tests and uninstall", node, [path.join("scripts", "verify-installed-release.cjs"), installer, version, path.join(root, "release", `release-verification-${version}.json`)], 30 * 60 * 1000);

process.stdout.write(`\nRelease gate passed for Exhibit Builder ${version}. The verification report must accompany release promotion.\n`);
