const { createHash } = require("node:crypto");
const { readFileSync } = require("node:fs");

function assertExactFile(installedPath, reviewedPath, label) {
  const installed = readFileSync(installedPath);
  const reviewed = readFileSync(reviewedPath);
  if (!installed.equals(reviewed)) throw new Error(`${label} differs from the reviewed source file`);
  return installed.toString("utf8");
}

function assertNonEmptyFile(filePath, label) {
  const bytes = readFileSync(filePath);
  if (bytes.byteLength === 0) throw new Error(`${label} is empty`);
  return {
    sha256: createHash("sha256").update(bytes).digest("hex"),
    size: bytes.byteLength,
  };
}

module.exports = { assertExactFile, assertNonEmptyFile };
