const { realpath, stat } = require("node:fs/promises");
const path = require("node:path");

const MAX_PROTECTED_PATHS = 5_000;
const MAX_PATH_BYTES = 32_768;
const PROJECT_ARCHIVE_SUFFIX = /\.bundle-project$/i;

class SourcePathCollisionError extends Error {
  constructor() {
    super("The chosen location is a selected witness statement, exhibit, email, or template. Choose a different file so the original is not replaced.");
    this.name = "SourcePathCollisionError";
    this.code = "SOURCE_PATH_COLLISION";
  }
}

function boundedAbsolutePath(value, label) {
  if (typeof value !== "string" || !value || value.length > MAX_PATH_BYTES) {
    throw new Error(`${label} is invalid.`);
  }
  if (value.includes("\0")) throw new Error(`${label} is invalid.`);
  if (!path.isAbsolute(value)) throw new Error(`${label} must be an absolute path.`);
  return value;
}

function parseProtectedSourcePaths(value) {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) throw new Error("Protected source paths are invalid.");
  if (value.length > MAX_PROTECTED_PATHS) throw new Error("Too many protected source paths were supplied.");
  const unique = new Set();
  for (const item of value) {
    unique.add(boundedAbsolutePath(item, "Protected source path"));
  }
  return [...unique];
}

function parseOptionalAbsolutePath(value, label) {
  if (value === undefined || value === null || value === "") return undefined;
  return boundedAbsolutePath(value, label);
}

function isProjectArchiveFileName(fileName) {
  return typeof fileName === "string" && PROJECT_ARCHIVE_SUFFIX.test(fileName);
}

function normalizePathKey(absolutePath) {
  const resolved = path.resolve(absolutePath);
  return process.platform === "win32" ? resolved.replace(/\//g, "\\").toLowerCase() : resolved;
}

async function pathIdentity(filePath) {
  try {
    const canonical = await realpath(filePath);
    const info = await stat(canonical, { bigint: true });
    const fileId = info.ino !== 0n ? `${info.dev.toString()}:${info.ino.toString()}` : null;
    return { key: normalizePathKey(canonical), fileId };
  } catch {
    try {
      const parent = await realpath(path.dirname(filePath));
      return { key: normalizePathKey(path.join(parent, path.basename(filePath))), fileId: null };
    } catch {
      return { key: normalizePathKey(filePath), fileId: null };
    }
  }
}

function identitiesMatch(left, right) {
  if (left.fileId && right.fileId && left.fileId === right.fileId) return true;
  return left.key === right.key;
}

async function assertDestinationAllowed(destination, protectedSourcePaths, options = {}) {
  const destinationPath = boundedAbsolutePath(destination, "Save destination");
  const destinationIdentity = await pathIdentity(destinationPath);
  if (isProjectArchiveFileName(options.fileName) && options.allowedOverwritePath) {
    const allowedIdentity = await pathIdentity(options.allowedOverwritePath);
    if (identitiesMatch(destinationIdentity, allowedIdentity)) return;
  }
  for (const protectedPath of protectedSourcePaths) {
    const sourceIdentity = await pathIdentity(protectedPath);
    if (identitiesMatch(destinationIdentity, sourceIdentity)) throw new SourcePathCollisionError();
  }
}

module.exports = {
  MAX_PROTECTED_PATHS,
  SourcePathCollisionError,
  parseProtectedSourcePaths,
  parseOptionalAbsolutePath,
  isProjectArchiveFileName,
  normalizePathKey,
  assertDestinationAllowed,
};
