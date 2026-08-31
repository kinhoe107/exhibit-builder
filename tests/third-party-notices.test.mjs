import assert from "node:assert/strict";
import { mkdtempSync, mkdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

const require = createRequire(import.meta.url);
const { generateThirdPartyNotices } = require("../scripts/generate-third-party-notices.cjs");
const root = path.resolve(import.meta.dirname, "..");

function manifestFor(importerManifest, packageName, optional = false) {
  const resolver = createRequire(importerManifest);
  try {
    return realpathSync(resolver.resolve(`${packageName}/package.json`));
  } catch {
    try {
      let current = path.dirname(resolver.resolve(packageName));
      while (path.dirname(current) !== current) {
        const candidate = path.join(current, "package.json");
        try {
          if (JSON.parse(readFileSync(candidate, "utf8")).name === packageName) return realpathSync(candidate);
        } catch {}
        current = path.dirname(current);
      }
    } catch {}
    if (optional) return null;
    throw new Error(`Test could not independently resolve ${packageName}.`);
  }
}

function independentProductionClosure(projectRoot) {
  const rootManifestPath = path.join(projectRoot, "package.json");
  const rootManifest = JSON.parse(readFileSync(rootManifestPath, "utf8"));
  const queue = Object.keys(rootManifest.dependencies || {}).map((name) => manifestFor(rootManifestPath, name));
  const identities = new Set();
  while (queue.length > 0) {
    const manifestPath = queue.shift();
    const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
    const identity = `${manifest.name}@${manifest.version}`;
    if (identities.has(identity)) continue;
    identities.add(identity);
    for (const name of Object.keys(manifest.dependencies || {})) queue.push(manifestFor(manifestPath, name));
    for (const name of Object.keys(manifest.optionalDependencies || {})) {
      const resolved = manifestFor(manifestPath, name, true);
      if (resolved) queue.push(resolved);
    }
  }
  return [...identities].sort();
}

test("third-party notices exactly enumerate the recursive installed production closure", () => {
  const generated = generateThirdPartyNotices(root);
  const expected = independentProductionClosure(root);
  const listed = [...generated.text.matchAll(/^\[production-package\] (\S+@\S+) — /gm)].map((match) => match[1]);

  assert.deepEqual(generated.packageIdentities, expected);
  assert.deepEqual(listed, expected);
  assert.equal(new Set(listed).size, listed.length);
  assert.equal(generateThirdPartyNotices(root).text, generated.text, "notice generation must be deterministic");
  assert.match(generated.text, /Bundled OCR assets/);
  assert.match(generated.text, /Tesseract OCR worker attributions/);
  assert.match(generated.text, /Bundled Electron runtime 43\.4\.0 — MIT/);
  assert.match(generated.text, /LICENSE\.electron\.txt/);
  assert.match(generated.text, /LICENSES\.chromium\.html/);
  assert.doesNotMatch(generated.text, /The Chromium Authors/);
});

test("notice generation rejects unknown licences and missing required licence text", () => {
  for (const fixture of [
    { license: "UNKNOWN", includeLicence: true, message: /no recognised declared licence/ },
    { license: "Apache-2.0", includeLicence: false, message: /no installed licence or notice file/ },
  ]) {
    const fixtureRoot = mkdtempSync(path.join(tmpdir(), "exhibit-builder-notices-"));
    try {
      const packageRoot = path.join(fixtureRoot, "node_modules", "fixture-package");
      mkdirSync(packageRoot, { recursive: true });
      writeFileSync(
        path.join(fixtureRoot, "package.json"),
        JSON.stringify({ name: "fixture-root", version: "1.0.0", dependencies: { "fixture-package": "1.0.0" } }),
      );
      writeFileSync(
        path.join(packageRoot, "package.json"),
        JSON.stringify({ name: "fixture-package", version: "1.0.0", license: fixture.license }),
      );
      writeFileSync(path.join(packageRoot, "index.js"), "module.exports = {};\n");
      if (fixture.includeLicence) writeFileSync(path.join(packageRoot, "LICENSE"), "placeholder\n");

      assert.throws(() => generateThirdPartyNotices(fixtureRoot), fixture.message);
    } finally {
      rmSync(fixtureRoot, { recursive: true, force: true });
    }
  }
});

test("source notice generation does not require a downloaded Electron runtime", () => {
  const fixtureRoot = mkdtempSync(path.join(tmpdir(), "exhibit-builder-clean-install-notices-"));
  try {
    const runtimePackage = path.join(fixtureRoot, "node_modules", "runtime-package");
    const electronPackage = path.join(fixtureRoot, "node_modules", "electron");
    const ocrRoot = path.join(fixtureRoot, "public", "ocr");
    mkdirSync(runtimePackage, { recursive: true });
    mkdirSync(electronPackage, { recursive: true });
    mkdirSync(ocrRoot, { recursive: true });
    writeFileSync(
      path.join(fixtureRoot, "package.json"),
      JSON.stringify({ name: "fixture-root", version: "1.0.0", dependencies: { "runtime-package": "1.0.0" }, devDependencies: { electron: "43.4.0" } }),
    );
    writeFileSync(path.join(runtimePackage, "package.json"), JSON.stringify({ name: "runtime-package", version: "1.0.0", license: "MIT", author: "Fixture Author" }));
    writeFileSync(path.join(runtimePackage, "index.js"), "module.exports = {};\n");
    writeFileSync(path.join(runtimePackage, "LICENSE"), "runtime package licence\n");
    writeFileSync(path.join(electronPackage, "package.json"), JSON.stringify({ name: "electron", version: "43.4.0", license: "MIT" }));
    writeFileSync(path.join(electronPackage, "index.js"), "module.exports = {};\n");
    writeFileSync(path.join(electronPackage, "LICENSE"), "electron licence\n");
    writeFileSync(path.join(ocrRoot, "worker.min.js.LICENSE.txt"), "worker attribution\n");
    writeFileSync(path.join(ocrRoot, "tesseract-core.LICENSE.txt"), "Apache License 2.0\n");

    const generated = generateThirdPartyNotices(fixtureRoot);
    assert.deepEqual(generated.packageIdentities, ["runtime-package@1.0.0"]);
    assert.match(generated.text, /LICENSES\.chromium\.html files beside the executable/);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});

const { resolvePackageLicence, loadLicenceRecords } = require("../scripts/licence-policy.cjs");

function writeNoticeFixture(fixtureRoot, { license, includeLicence = true, elections = null }) {
  const packageRoot = path.join(fixtureRoot, "node_modules", "fixture-package");
  const electronPackage = path.join(fixtureRoot, "node_modules", "electron");
  const ocrRoot = path.join(fixtureRoot, "public", "ocr");
  mkdirSync(packageRoot, { recursive: true });
  mkdirSync(electronPackage, { recursive: true });
  mkdirSync(ocrRoot, { recursive: true });
  writeFileSync(
    path.join(fixtureRoot, "package.json"),
    JSON.stringify({ name: "fixture-root", version: "1.0.0", dependencies: { "fixture-package": "1.0.0" }, devDependencies: { electron: "43.4.0" } }),
  );
  writeFileSync(
    path.join(packageRoot, "package.json"),
    JSON.stringify({ name: "fixture-package", version: "1.0.0", license }),
  );
  writeFileSync(path.join(packageRoot, "index.js"), "module.exports = {};\n");
  if (includeLicence) writeFileSync(path.join(packageRoot, "LICENSE"), "upstream dual licence text including MIT and GPL\n");
  writeFileSync(path.join(electronPackage, "package.json"), JSON.stringify({ name: "electron", version: "43.4.0", license: "MIT" }));
  writeFileSync(path.join(electronPackage, "index.js"), "module.exports = {};\n");
  writeFileSync(path.join(electronPackage, "LICENSE"), "electron licence\n");
  writeFileSync(path.join(ocrRoot, "worker.min.js.LICENSE.txt"), "worker attribution\n");
  writeFileSync(path.join(ocrRoot, "tesseract-core.LICENSE.txt"), "Apache License 2.0\n");
  if (elections) {
    mkdirSync(path.join(fixtureRoot, "scripts"), { recursive: true });
    writeFileSync(path.join(fixtureRoot, "scripts", "licence-elections.json"), JSON.stringify(elections, null, 2));
  }
}

test("every production package has a resolved licence status and JSZip elects MIT", () => {
  const generated = generateThirdPartyNotices(root);
  assert.ok(generated.resolutions.length > 0);
  assert.equal(generated.resolutions.length, generated.packageIdentities.length);
  for (const resolution of generated.resolutions) {
    assert.ok(["allowlisted", "elected", "approved"].includes(resolution.status), `${resolution.identity} status ${resolution.status}`);
  }
  const jszip = generated.resolutions.find((item) => item.identity === "jszip@3.10.1");
  assert.ok(jszip, "jszip@3.10.1 must be in the production graph");
  assert.equal(jszip.status, "elected");
  assert.equal(jszip.elected, "MIT");
  assert.match(generated.text, /EXHIBIT BUILDER LICENCE ELECTION/);
  assert.match(generated.text, /Exhibit Builder uses JSZip 3\.10\.1 under the MIT licence option/);
  assert.match(generated.text, /JSZip is dual licensed/i);
  assert.match(generated.text, /GNU GENERAL PUBLIC LICENSE/);
  const records = loadLicenceRecords(root);
  assert.equal(resolvePackageLicence({ identity: "jszip@3.10.1", declared: "(MIT OR GPL-3.0-or-later)", records }).elected, "MIT");
});

test("notice generation fails closed on GPL-only, AGPL-only and unelected dual licences", () => {
  for (const fixture of [
    { license: "GPL-3.0-only", message: /only under GPL/ },
    { license: "AGPL-3.0-only", message: /only under AGPL/ },
    { license: "(MIT OR GPL-3.0-or-later)", message: /no recorded compatible licence election/ },
  ]) {
    const fixtureRoot = mkdtempSync(path.join(tmpdir(), "exhibit-builder-licence-fail-"));
    try {
      writeNoticeFixture(fixtureRoot, { license: fixture.license });
      assert.throws(() => generateThirdPartyNotices(fixtureRoot), fixture.message);
    } finally {
      rmSync(fixtureRoot, { recursive: true, force: true });
    }
  }
});

test("elected dual licence passes and preserves upstream text after the Exhibit Builder note", () => {
  const fixtureRoot = mkdtempSync(path.join(tmpdir(), "exhibit-builder-licence-elect-"));
  try {
    writeNoticeFixture(fixtureRoot, {
      license: "(MIT OR GPL-3.0-or-later)",
      elections: {
        elections: {
          "fixture-package@1.0.0": { declared: "(MIT OR GPL-3.0-or-later)", elected: "MIT" },
        },
        approvals: {},
      },
    });
    const generated = generateThirdPartyNotices(fixtureRoot);
    const electionIndex = generated.text.indexOf("EXHIBIT BUILDER LICENCE ELECTION");
    const upstreamIndex = generated.text.indexOf("upstream dual licence text including MIT and GPL");
    assert.ok(electionIndex >= 0);
    assert.ok(upstreamIndex > electionIndex);
    assert.match(generated.text, /\[production-package\] fixture-package@1\.0\.0 — \(MIT OR GPL-3\.0-or-later\)/);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});

test("allowlisted conjunctions pass without an election", () => {
  const fixtureRoot = mkdtempSync(path.join(tmpdir(), "exhibit-builder-licence-and-"));
  try {
    writeNoticeFixture(fixtureRoot, { license: "MIT AND Zlib" });
    const generated = generateThirdPartyNotices(fixtureRoot);
    assert.deepEqual(generated.resolutions.map((item) => item.status), ["allowlisted"]);
    assert.doesNotMatch(generated.text, /EXHIBIT BUILDER LICENCE ELECTION/);
  } finally {
    rmSync(fixtureRoot, { recursive: true, force: true });
  }
});
