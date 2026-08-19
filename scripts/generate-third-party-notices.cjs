const { existsSync, readFileSync, realpathSync, readdirSync, statSync, writeFileSync } = require("node:fs");
const { createRequire } = require("node:module");
const path = require("node:path");

const DEFAULT_ROOT = path.resolve(__dirname, "..");
const LICENCE_FILE_PATTERN = /^(?:licen[cs]e|copying|notice)(?:[._-].*)?$/i;

function normaliseText(value) {
  return value.replace(/\r\n?/g, "\n").trim();
}

function compareText(left, right) {
  return left < right ? -1 : left > right ? 1 : 0;
}

function readManifest(manifestPath) {
  const manifest = JSON.parse(readFileSync(manifestPath, "utf8"));
  if (!manifest.name || !manifest.version) {
    throw new Error(`Dependency manifest lacks a name or version: ${manifestPath}`);
  }
  return manifest;
}

function findOwningManifest(entryPath, expectedName) {
  let current = statSync(entryPath).isDirectory() ? entryPath : path.dirname(entryPath);
  while (path.dirname(current) !== current) {
    const candidate = path.join(current, "package.json");
    if (existsSync(candidate)) {
      const manifest = JSON.parse(readFileSync(candidate, "utf8"));
      if (manifest.name === expectedName) return candidate;
    }
    current = path.dirname(current);
  }
  throw new Error(`Could not locate the installed manifest for ${expectedName}.`);
}

function resolveDependencyManifest(importerManifestPath, dependencyName, optional = false) {
  const resolver = createRequire(importerManifestPath);
  try {
    return realpathSync(resolver.resolve(`${dependencyName}/package.json`));
  } catch (manifestError) {
    try {
      return realpathSync(findOwningManifest(resolver.resolve(dependencyName), dependencyName));
    } catch (entryError) {
      if (optional) return null;
      throw new Error(
        `Required production dependency ${dependencyName} could not be resolved from ${importerManifestPath}: ${entryError.message}`,
        { cause: manifestError },
      );
    }
  }
}

function collectProductionPackages(root = DEFAULT_ROOT) {
  const rootManifestPath = path.join(root, "package.json");
  const rootManifest = JSON.parse(readFileSync(rootManifestPath, "utf8"));
  const pending = Object.keys(rootManifest.dependencies || {})
    .sort(compareText)
    .map((name) => ({ manifestPath: resolveDependencyManifest(rootManifestPath, name), name }));
  const packages = new Map();

  while (pending.length > 0) {
    const { manifestPath } = pending.shift();
    const manifest = readManifest(manifestPath);
    const identity = `${manifest.name}@${manifest.version}`;
    if (packages.has(identity)) continue;

    const packageRoot = path.dirname(manifestPath);
    packages.set(identity, { identity, manifest, manifestPath, packageRoot });

    for (const name of Object.keys(manifest.dependencies || {}).sort(compareText)) {
      pending.push({ manifestPath: resolveDependencyManifest(manifestPath, name), name });
    }
    for (const name of Object.keys(manifest.optionalDependencies || {}).sort(compareText)) {
      const optionalManifest = resolveDependencyManifest(manifestPath, name, true);
      if (optionalManifest) pending.push({ manifestPath: optionalManifest, name });
    }
  }

  return [...packages.values()].sort((left, right) => compareText(left.identity, right.identity));
}

function declaredLicence(manifest, identity) {
  const value = typeof manifest.license === "string"
    ? manifest.license
    : typeof manifest.license?.type === "string"
      ? manifest.license.type
      : null;
  if (!value || !value.trim() || /^(?:unknown|unlicensed|none)$/i.test(value.trim())) {
    throw new Error(`Production dependency ${identity} has no recognised declared licence.`);
  }
  return value.trim();
}

function licenceFilesForPackage(packageRecord, required = true) {
  const names = readdirSync(packageRecord.packageRoot)
    .filter((name) => LICENCE_FILE_PATTERN.test(name))
    .filter((name) => statSync(path.join(packageRecord.packageRoot, name)).isFile())
    .sort(compareText);
  if (required && names.length === 0) {
    throw new Error(`Production dependency ${packageRecord.identity} has no installed licence or notice file.`);
  }
  return names;
}

function repositoryKey(manifest) {
  const repository = typeof manifest.repository === "string" ? manifest.repository : manifest.repository?.url;
  return typeof repository === "string" ? repository.trim().toLowerCase().replace(/\.git$/, "") : null;
}

function authorName(manifest) {
  if (typeof manifest.author === "string") return manifest.author.replace(/\s*<[^>]*>\s*$/, "").trim();
  return typeof manifest.author?.name === "string" ? manifest.author.name.trim() : "";
}

function canonicalMitText(manifest, identity) {
  const author = authorName(manifest);
  if (!author) throw new Error(`Production dependency ${identity} ships no MIT notice and has no author attribution.`);
  return [
    `Copyright (c) ${author}`,
    "",
    "Permission is hereby granted, free of charge, to any person obtaining a copy",
    "of this software and associated documentation files (the \"Software\"), to deal",
    "in the Software without restriction, including without limitation the rights",
    "to use, copy, modify, merge, publish, distribute, sublicense, and/or sell",
    "copies of the Software, and to permit persons to whom the Software is",
    "furnished to do so, subject to the following conditions:",
    "",
    "The above copyright notice and this permission notice shall be included in all",
    "copies or substantial portions of the Software.",
    "",
    "THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR",
    "IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,",
    "FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE",
    "AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER",
    "LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,",
    "OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE",
    "SOFTWARE.",
  ].join("\n");
}

function appendFile(sections, heading, filePath) {
  if (!existsSync(filePath)) throw new Error(`Required third-party notice is missing: ${filePath}`);
  const content = normaliseText(readFileSync(filePath, "utf8"));
  if (!content) throw new Error(`Required third-party notice is empty: ${filePath}`);
  sections.push(heading, "-".repeat(78), content, "");
}

function generateThirdPartyNotices(root = DEFAULT_ROOT) {
  const packages = collectProductionPackages(root);
  const sections = [
    "EXHIBIT BUILDER THIRD-PARTY LICENCE NOTICES",
    "",
    "Exhibit Builder incorporates the components listed below. Those components",
    "remain subject to their own terms; this file does not change those terms.",
    "",
    "The production-package inventory is generated recursively from the exact",
    "installed dependency graph. Package entries are ordered by name and version.",
    "",
  ];

  const licenceSources = new Map();
  for (const packageRecord of packages) {
    const licence = declaredLicence(packageRecord.manifest, packageRecord.identity);
    const repository = repositoryKey(packageRecord.manifest);
    const files = licenceFilesForPackage(packageRecord, false);
    if (repository && files.length > 0) {
      const key = `${repository}\0${licence.toLowerCase()}`;
      if (!licenceSources.has(key)) licenceSources.set(key, { packageRecord, files });
    }
  }

  for (const packageRecord of packages) {
    const licence = declaredLicence(packageRecord.manifest, packageRecord.identity);
    let sourceRecord = packageRecord;
    let files = licenceFilesForPackage(packageRecord, false);
    if (files.length === 0) {
      const repository = repositoryKey(packageRecord.manifest);
      const sharedSource = repository
        ? licenceSources.get(`${repository}\0${licence.toLowerCase()}`)
        : null;
      if (sharedSource) {
        sourceRecord = sharedSource.packageRecord;
        files = sharedSource.files;
      } else if (licence.toUpperCase() !== "MIT") {
        throw new Error(`Production dependency ${packageRecord.identity} has no installed licence or notice file.`);
      }
    }
    sections.push(
      "=".repeat(78),
      `[production-package] ${packageRecord.identity} — ${licence}`,
      "=".repeat(78),
      "",
    );
    if (files.length === 0) {
      sections.push(
        "Canonical MIT text with the installed package's author attribution",
        "-".repeat(78),
        canonicalMitText(packageRecord.manifest, packageRecord.identity),
        "",
      );
    } else if (sourceRecord !== packageRecord) {
      sections.push(`Licence text supplied by the same upstream project in ${sourceRecord.identity}.`, "");
    }
    for (const name of files) {
      appendFile(sections, name, path.join(sourceRecord.packageRoot, name));
    }
  }

  sections.push("=".repeat(78), "Bundled OCR assets", "=".repeat(78), "");
  appendFile(sections, "Tesseract OCR worker attributions", path.join(root, "public", "ocr", "worker.min.js.LICENSE.txt"));
  appendFile(
    sections,
    "Apache License 2.0 — Tesseract core WebAssembly and English trained data",
    path.join(root, "public", "ocr", "tesseract-core.LICENSE.txt"),
  );

  const electronManifestPath = resolveDependencyManifest(path.join(root, "package.json"), "electron");
  const electronManifest = readManifest(electronManifestPath);
  const electronRoot = path.dirname(electronManifestPath);
  const electronLicence = declaredLicence(electronManifest, `electron@${electronManifest.version}`);
  sections.push(
    "=".repeat(78),
    `Bundled Electron runtime ${electronManifest.version} — ${electronLicence}`,
    "=".repeat(78),
    "",
  );
  appendFile(sections, "Electron licence", path.join(electronRoot, "LICENSE"));
  sections.push(
    "The packaged application carries Electron's separate LICENSE.electron.txt and",
    "LICENSES.chromium.html files beside the executable. The mandatory installed-release",
    "verification refuses promotion if either runtime notice is absent or empty and",
    "records the SHA-256 digest of both files in the verification report.",
    "",
  );

  return {
    packageIdentities: packages.map(({ identity }) => identity),
    text: `${sections.join("\n")}\n`,
  };
}

function writeThirdPartyNotices(root = DEFAULT_ROOT) {
  const result = generateThirdPartyNotices(root);
  writeFileSync(path.join(root, "THIRD_PARTY_LICENSES.txt"), result.text, "utf8");
  return result;
}

if (require.main === module) writeThirdPartyNotices();

module.exports = {
  collectProductionPackages,
  generateThirdPartyNotices,
  resolveDependencyManifest,
  writeThirdPartyNotices,
};
