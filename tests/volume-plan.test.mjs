import assert from "node:assert/strict";
import test from "node:test";
import { File } from "node:buffer";
import { createHash, webcrypto } from "node:crypto";
import { PDFDocument, PDFName } from "pdf-lib";
import JSZip from "jszip";

globalThis.File = File;
globalThis.crypto ??= webcrypto;

test("plans administrative AH1 volumes without splitting an exhibit", async () => {
  const { planVolumes, volumeReference } = await import("../app/lib/volume-plan.ts");
  const plan = planVolumes([{ id: "one", pages: 4 }, { id: "two", pages: 7 }, { id: "three", pages: 3 }], 10);
  assert.deepEqual(plan.map((volume) => [volume.number, volume.items.map((item) => item.id), volume.pages, volume.oversize]), [[1, ["one"], 4, false], [2, ["two", "three"], 10, false]]);
  assert.equal(volumeReference("AH1", 1, 34, 34), "AH1/34");
  assert.equal(volumeReference("AH1", 1, 34, 34, true), "AH1/Vol. 1/34");
  assert.equal(volumeReference("AH1", 2, 3, 8), "AH1/Vol. 2/3-8");
  const oversize = planVolumes([{ id: "long", pages: 22 }, { id: "later", pages: 2 }], 10);
  assert.deepEqual(oversize.map((volume) => [volume.items.map((item) => item.id), volume.oversize]), [[["long"], true], [["later"], false]]);
});

test("creates separate AH1 volume PDFs and keeps index numbers continuous", async () => {
  const { buildBundle, buildStatementUpdateSuggestions } = await import("../app/lib/bundle-engine.ts");
  const pdf = await PDFDocument.create(); pdf.addPage([595.28, 841.89]).drawText("Evidence");
  const source = new File([await pdf.save()], "evidence.pdf", { type: "application/pdf" });
  const evidence = [{ id: "e", file: source, name: source.name, extension: "pdf", text: "Evidence", marker: null, sha256: "e".repeat(64), pageCount: 1, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" }];
  const candidate = (id, order, paragraph) => ({ id, mark: "AH 1", provisionalNumber: order, description: `Exhibit ${order}`, date: "Date not stated", paragraph, citation: "[AH1/xx]", exhibitInitials: "AH", exhibitSequence: 1, discoverySignals: [], evidenceId: "e", confidence: 100, rationale: "test", included: true, confirmed: true, witnessInitials: "AH", witnessKey: "AH", statementId: "s", statementName: "Statement.docx", sequenceOrder: order, repeatDecision: order === 1 ? undefined : "separate" });
  const candidates = [candidate("one", 1, 3), candidate("two", 2, 5)];
  const statement = new File(["read-only statement"], "Statement.docx");
  const statementHash = createHash("sha256").update(new Uint8Array(await statement.arrayBuffer())).digest("hex");
  const statementInput = { id: "s", file: statement, witnessName: "Witness", witnessInitials: "AH" };
  const analysis = { statementName: statement.name, statementHash, statementId: statementInput.id, statementSources: [{ statementId: statementInput.id, fileName: statement.name, sha256: statementHash }], statementSnapshots: [statementInput], statementHandles: [statementInput], caseTitle: "Test", candidates, evidence, unreferenced: [], statementWarnings: [], generatedAt: new Date().toISOString() };
  const arrangement = { version: 1, nodes: [{ type: "section", id: "agreements", heading: "Agreements and reports", exhibits: [{ type: "exhibit", exhibitId: "candidate-one" }, { type: "exhibit", exhibitId: "candidate-two" }] }] };
  const result = await buildBundle(analysis, candidates, { arrangement, layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 1 } });
  assert.equal(result.volumes?.length, 2);
  assert.deepEqual(result.volumes?.map((volume) => volume.label), ["AH 1 — Volume 1", "AH 1 — Volume 2"]);
  assert.deepEqual(result.volumes?.flatMap((volume) => volume.records.map((record) => record.exhibitNumber)), [1, 2]);
  assert.deepEqual(buildStatementUpdateSuggestions(result.volumes?.flatMap((volume) => volume.records) ?? []).map((item) => item.line), ["Paragraph 3 - [AH1/Vol. 1/3]", "Paragraph 5 - [AH1/Vol. 2/6]"]);
  assert.equal(result.fileName, "Exhibit_Bundle_AH1_Volumes.zip");
  assert.equal(result.manifest.output.volumeCount, 2);
  assert.equal(result.records.length, 2);
  assert.equal(result.pageCount, result.volumes.reduce((total, volume) => total + volume.pageCount, 0));
  assert.ok(result.volumes.every((volume) => volume.pageCount === 3 && volume.checks.some((check) => check.label === "Oversize exhibit volume")));
  assert.ok(result.volumes.every((volume) => volume.manifest.output.fileName === volume.fileName && volume.manifest.output.sha256 === volume.sha256));
  assert.ok(result.volumes.every((volume) => volume.manifest.exhibits.every((record) => record.volumeNumber === volume.number)));
  assert.ok(result.volumes.every((volume) => volume.manifest.exhibits.every((record) => record.statementReferences.every((reference) => reference.reviewProvenance?.candidateId && reference.reviewProvenance.confirmed === true))));
  assert.ok(result.manifest.exhibits.every((record) => record.statementReferences.every((reference) => reference.reviewProvenance?.candidateId)));
  assert.equal(result.manifest.pagination.volumeNumbering, "continuous");
  assert.doesNotMatch(JSON.stringify(result.manifest), /AH\s*2(?:\D|$)/, "administrative volume splitting never creates AH2");
  const zip = await JSZip.loadAsync(result.volumeZipBytes);
  assert.deepEqual(Object.keys(zip.files).sort(), ["Exhibit_Bundle_AH1_Volume_1.pdf", "Exhibit_Bundle_AH1_Volume_1_Manifest.json", "Exhibit_Bundle_AH1_Volume_2.pdf", "Exhibit_Bundle_AH1_Volume_2_Manifest.json"]);
  for (const volume of result.volumes) {
    const reopened = await PDFDocument.load(await zip.file(volume.fileName).async("uint8array"));
    assert.equal(reopened.getPageCount(), volume.pageCount);
    assert.equal(reopened.getPages()[1].node.Annots()?.size(), 1, "only the local row is linked in each repeated complete index");
    const outlineRoot = reopened.context.lookup(reopened.catalog.get(PDFName.of("Outlines")));
    const section = reopened.context.lookup(outlineRoot.get(PDFName.of("First")));
    assert.ok(section.get(PDFName.of("First")), "the local exhibit is a child of its section bookmark");
    assert.equal(section.get(PDFName.of("Next")), undefined, "no remote or empty top-level bookmark is added");
    const localExhibit = reopened.context.lookup(section.get(PDFName.of("First")));
    assert.equal(localExhibit.get(PDFName.of("Title")).decodeText(), `${volume.records[0].exhibitNumber}. ${volume.records[0].description}`);
  }

  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const visibleStarts = [];
  for (const [index, volume] of result.volumes.entries()) {
    const rendered = await pdfjs.getDocument({ data: new Uint8Array(volume.bytes) }).promise;
    const coverText = (await (await rendered.getPage(1)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
    const indexText = (await (await rendered.getPage(2)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ");
    assert.match(coverText, new RegExp(`Volume ${index + 1} of 2`));
    assert.doesNotMatch(coverText, /AH 1/, "internal bundle identity is not printed on the cover");
    assert.match(indexText, /Agreements and reports/);
    assert.match(indexText, /Exhibit 1/);
    assert.match(indexText, /Exhibit 2/);
    assert.match(indexText, /Vol\. 1/);
    assert.match(indexText, /Vol\. 2/);
    assert.match(indexText, /\b3\b/);
    assert.match(indexText, /\b6\b/, "every volume repeats the full index with final cross-volume page references");
    visibleStarts.push(await rendered.getPageLabels());
    await rendered.destroy();
  }
  assert.deepEqual(visibleStarts.map((labels) => labels[0]), ["1", "4"]);

  const restarted = await buildBundle(analysis, candidates, {
    pagination: { volumeNumbering: "restart" },
    layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 1 },
  });
  assert.equal(restarted.manifest.pagination.volumeNumbering, "restart");
  for (const volume of restarted.volumes) {
    const rendered = await pdfjs.getDocument({ data: new Uint8Array(volume.bytes) }).promise;
    assert.equal((await rendered.getPageLabels())[0], "1", "an explicitly restarted volume begins again at page 1");
    await rendered.destroy();
  }

  const separateContinuous = await buildBundle(analysis, candidates, {
    pagination: { matchPdfPageOrder: false, volumeNumbering: "continuous", prefix: "LV-", countTemplates: true, preliminary: "arabic" },
    layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 1 },
  });
  const separateLabels = await Promise.all(separateContinuous.volumes.map(async (volume) => {
    const rendered = await pdfjs.getDocument({ data: new Uint8Array(volume.bytes) }).promise;
    const labels = await rendered.getPageLabels();
    await rendered.destroy();
    return labels;
  }));
  assert.deepEqual(separateLabels[0], ["LV-1", "LV-2", "LV-3"]);
  assert.deepEqual(separateLabels[1], ["LV-4", "LV-5", "LV-6"], "a saved exhibit-local numbering request still continues the finished PDF sequence across volumes");
});
