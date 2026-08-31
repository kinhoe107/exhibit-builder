import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

test("documents and enforces the English-only matching and OCR boundary", async () => {
  const [{ parseStatementCitationTokens }, ocrSource, uiSource, readme] = await Promise.all([
    import("../app/lib/bundle-engine.ts"),
    readFile(new URL("../app/lib/ocr.ts", import.meta.url), "utf8"),
    readFile(new URL("../app/BundleBuilder.tsx", import.meta.url), "utf8"),
    readFile(new URL("../README.md", import.meta.url), "utf8"),
  ]);

  const supportedSyntaxInFrenchContext = parseStatementCitationTokens("Je me réfère au contrat et aux courriels [Exhibit; AH-xx].");
  assert.equal(supportedSyntaxInFrenchContext.length, 2, "documented ASCII placeholders remain detectable in surrounding Unicode text");
  assert.equal(parseStatementCitationTokens("Je produis la pièce [Pièce A].").length, 0, "the parser does not pretend that unimplemented translated placeholder vocabulary is supported");
  assert.match(ocrSource, /createWorker\("eng",/, "the shipped OCR worker is explicitly English");
  assert.match(uiSource, /Automatic matching and OCR are English-only in this release/);
  assert.match(uiSource, /multilingual matching, OCR and generated\s+typesetting are not supported/);
  assert.match(readme, /The tested recognition, matching and OCR capability is English/);
});
