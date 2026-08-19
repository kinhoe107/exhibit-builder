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
