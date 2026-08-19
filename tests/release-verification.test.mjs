import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

const require = createRequire(import.meta.url);
const { assertExactFile, assertNonEmptyFile } = require("../scripts/release-verification-helpers.cjs");

test("pnpm test lists every public automated test file", async () => {
  const { readdir } = await import("node:fs/promises");
  const packageJson = JSON.parse(await readFile(new URL("../package.json", import.meta.url), "utf8"));
  const listed = `${packageJson.scripts.test}\n${packageJson.scripts["test:coverage:critical"]}`;
  const names = (await readdir(new URL("../tests", import.meta.url)))
    .filter((name) => name.endsWith(".test.mjs") || name === "run-guided-sample.mjs");
  for (const name of names) {
    assert.ok(listed.includes(`tests/${name}`), `${name} is missing from pnpm test`);
  }
  assert.doesNotMatch(packageJson.scripts.test, /run-adversarial-pack/, "the private adversarial pack stays on its own script");
});

test("release promotion is gated by source, package, unpacked and installed verification", async () => {
  const gate = await readFile(new URL("../scripts/release-gate.cjs", import.meta.url), "utf8");
  const installed = await readFile(new URL("../scripts/verify-installed-release.cjs", import.meta.url), "utf8");
  const installer = await readFile(new URL("../build/standalone-installer.nsi", import.meta.url), "utf8");
  assert.match(gate, /complete automated test suite/);
  assert.match(gate, /TypeScript-aware lint/);
  assert.match(gate, /critical bundle-logic coverage/);
  assert.match(gate, /--test-coverage-lines=95/);
  assert.match(gate, /--test-coverage-branches=80/);
  assert.match(gate, /--test-coverage-functions=95/);
  assert.match(gate, /native Microsoft Excel guided-sample fidelity test/);
  assert.match(gate, /EXHIBIT_BUILDER_NATIVE_EXCEL/);
  assert.match(gate, /private ICC adversarial fidelity test/);
  assert.match(gate, /EXHIBIT_BUILDER_REQUIRE_PRIVATE_STRESS/);
  assert.match(gate, /unpacked shell smoke test/);
  assert.match(gate, /unpacked guided-sample analysis smoke test/);
  assert.match(gate, /temporary installation, installed smoke tests and uninstall/);
  assert.match(gate, /Release promotion is blocked/);
  assert.match(installed, /--smoke-test/);
  assert.match(installed, /--analysis-smoke-test/);
  assert.match(installed, /EULA\.txt/);
  assert.match(installed, /THIRD_PARTY_LICENSES\.txt/);
  assert.match(installed, /LICENSE\.electron\.txt/);
  assert.match(installed, /LICENSES\.chromium\.html/);
  assert.match(installed, /assertExactFile/);
  assert.match(installed, /assertNonEmptyFile/);
  assert.match(installed, /app\.asar/);
  assert.match(installed, /Uninstall Exhibit Builder\.exe/);
  assert.match(installer, /\/RELEASEVERIFY/);
  assert.match(installer, /\.exhibit-builder-verification-install/);
});

test("installed Electron runtime notices must be non-empty and are hashed", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "exhibit-runtime-notice-verification-"));
  try {
    const notice = path.join(root, "notice.txt");
    await writeFile(notice, "runtime licence\n");
    assert.deepEqual(assertNonEmptyFile(notice, "runtime notice"), {
      sha256: "54000155d2109a9e0a886a99cd02d3c0e6d3580b1ebba485f3162720a71cb3f8",
      size: 16,
    });
    await writeFile(notice, "");
    assert.throws(() => assertNonEmptyFile(notice, "runtime notice"), /runtime notice is empty/);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("installed legal files must match reviewed bytes rather than only titles or keywords", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "exhibit-legal-verification-"));
  try {
    const reviewed = path.join(root, "reviewed.txt");
    const installed = path.join(root, "installed.txt");
    await writeFile(reviewed, "LICENCE TITLE\nVersion 1.1\nApproved clause\n");
    await writeFile(installed, "LICENCE TITLE\nVersion 1.1\nApproved clause\n");
    assert.match(assertExactFile(installed, reviewed, "licence"), /Approved clause/);
    await writeFile(installed, "LICENCE TITLE\nVersion 1.0\nApproved clause\n");
    assert.throws(() => assertExactFile(installed, reviewed, "licence"), /differs from the reviewed source file/i);
    await writeFile(installed, "LICENCE TITLE\nVersion 1.1\nChanged clause\n");
    assert.throws(() => assertExactFile(installed, reviewed, "licence"), /differs from the reviewed source file/i);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});
