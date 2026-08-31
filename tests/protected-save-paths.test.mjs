import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createRequire } from "node:module";
import test from "node:test";

const require = createRequire(import.meta.url);
const {
  SourcePathCollisionError,
  parseProtectedSourcePaths,
  parseOptionalAbsolutePath,
  isProjectArchiveFileName,
  normalizePathKey,
  assertDestinationAllowed,
} = require("../electron/protected-save-paths.cjs");

test("renderer-supplied save paths are bounded and must be absolute", () => {
  assert.deepEqual(parseProtectedSourcePaths(undefined), []);
  assert.deepEqual(parseProtectedSourcePaths(null), []);
  assert.equal(parseOptionalAbsolutePath("", "Allowed overwrite path"), undefined);
  assert.throws(() => parseProtectedSourcePaths("C:\\\\evidence.pdf"), /Protected source paths are invalid/);
  assert.throws(() => parseProtectedSourcePaths(["relative.pdf"]), /must be an absolute path/);
  assert.throws(() => parseProtectedSourcePaths([`${path.resolve("/tmp/a")}\0.pdf`]), /invalid/);
  assert.throws(() => parseOptionalAbsolutePath("outputs/bundle.pdf", "Allowed overwrite path"), /must be an absolute path/);
  const tooMany = Array.from({ length: 5001 }, (_, index) => path.resolve(`/tmp/source-${index}.pdf`));
  assert.throws(() => parseProtectedSourcePaths(tooMany), /Too many protected source paths/);
});

test("project archive names are the only allowed overwrite targets", () => {
  assert.equal(isProjectArchiveFileName("Matter.bundle-project"), true);
  assert.equal(isProjectArchiveFileName("bundle.pdf"), false);
  assert.equal(isProjectArchiveFileName("volumes.zip"), false);
});

test("Windows path identity is case-insensitive and separator-stable", () => {
  if (process.platform !== "win32") return;
  const left = normalizePathKey("C:/Evidence/Statement.DOCX");
  const right = normalizePathKey("c:\\evidence\\statement.docx");
  assert.equal(left, right);
});

test("saving generated output onto a selected source is refused before any write", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "exhibit-protected-save-"));
  try {
    const source = path.join(root, "Evidence.pdf");
    const other = path.join(root, "Bundle.pdf");
    await writeFile(source, "source-bytes");
    await assert.rejects(
      () => assertDestinationAllowed(source, [source], { fileName: "Bundle.pdf" }),
      (error) => error instanceof SourcePathCollisionError && error.code === "SOURCE_PATH_COLLISION",
    );
    await assertDestinationAllowed(other, [source], { fileName: "Bundle.pdf" });
    assert.equal(await readFile(source, "utf8"), "source-bytes");
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("project save may replace the opened project archive and still cannot replace evidence", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "exhibit-protected-project-save-"));
  try {
    const project = path.join(root, "Matter.bundle-project");
    const evidence = path.join(root, "Evidence.pdf");
    await writeFile(project, "project-bytes");
    await writeFile(evidence, "source-bytes");
    await assertDestinationAllowed(project, [evidence, project], {
      fileName: "Matter.bundle-project",
      allowedOverwritePath: project,
    });
    await assert.rejects(
      () => assertDestinationAllowed(evidence, [evidence, project], {
        fileName: "Matter.bundle-project",
        allowedOverwritePath: project,
      }),
      SourcePathCollisionError,
    );
    await assert.rejects(
      () => assertDestinationAllowed(project, [evidence, project], { fileName: "Bundle.pdf", allowedOverwritePath: project }),
      SourcePathCollisionError,
    );
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("filesystem identity refuses an alias of a selected source when the platform can create one", async () => {
  const root = await mkdtemp(path.join(tmpdir(), "exhibit-protected-alias-"));
  try {
    const folder = path.join(root, "matter");
    await mkdir(folder);
    const source = path.join(folder, "Evidence.pdf");
    await writeFile(source, "source-bytes");
    const aliasDir = path.join(root, "alias");
    try {
      await symlink(folder, aliasDir, process.platform === "win32" ? "junction" : "dir");
    } catch {
      return;
    }
    const aliased = path.join(aliasDir, "Evidence.pdf");
    await assert.rejects(
      () => assertDestinationAllowed(aliased, [source], { fileName: "Bundle.pdf" }),
      SourcePathCollisionError,
    );
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("the save-file handler refuses before creating the temporary replacement file", async () => {
  const main = await readFile(new URL("../electron/main.cjs", import.meta.url), "utf8");
  assert.match(main, /assertDestinationAllowed/);
  assert.match(main, /parseProtectedSourcePaths/);
  assert.match(main, /SOURCE_PATH_COLLISION/);
  assert.ok(main.indexOf("assertDestinationAllowed") < main.indexOf("writeFile(temporary"));
  const preload = await readFile(new URL("../electron/preload.cjs", import.meta.url), "utf8");
  assert.match(preload, /protectedSourcePaths/);
  assert.match(preload, /allowedOverwritePath/);
  const engine = await readFile(new URL("../app/lib/bundle-engine.ts", import.meta.url), "utf8");
  assert.match(engine, /setSavePathProtectionReader/);
  assert.match(engine, /protectedSourcePaths: protection\.protectedSourcePaths/);
});
