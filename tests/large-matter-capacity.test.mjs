import assert from "node:assert/strict";
import { webcrypto } from "node:crypto";
import { File } from "node:buffer";
import test from "node:test";
import JSZip from "jszip";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

globalThis.File = File;
globalThis.crypto ??= webcrypto;

const COUNT = 200;

function xmlEscape(value) {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

async function statementFile() {
  const paragraphs = Array.from({ length: COUNT }, (_, index) => {
    const number = String(index + 1).padStart(3, "0");
    const text = `${index + 1}. I refer to capacity exhibit ${number}, supplied for the deterministic large-matter test [Exhibit].`;
    return `<w:p><w:r><w:t>${xmlEscape(text)}</w:t></w:r></w:p>`;
  }).join("");
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`);
  return new File([await zip.generateAsync({ type: "uint8array" })], "Capacity_Witness_Statement.docx", { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function evidenceFiles() {
  const files = [];
  for (let index = 0; index < COUNT; index += 1) {
    const number = String(index + 1).padStart(3, "0");
    const pdf = await PDFDocument.create();
    const font = await pdf.embedFont(StandardFonts.Helvetica);
    const page = pdf.addPage([595.28, 841.89]);
    page.drawText(`CAPACITY EXHIBIT ${number}`, { x: 54, y: 770, size: 18, font, color: rgb(0.05, 0.18, 0.32) });
    page.drawText(`Deterministic text-native source ${number}`, { x: 54, y: 735, size: 11, font });
    files.push(new File([await pdf.save()], `Capacity_Exhibit_${number}.pdf`, { type: "application/pdf" }));
  }
  return files;
}

test("analyses, builds, saves and reopens a 200-exhibit long-statement matter", { timeout: 120_000 }, async () => {
  const [{ analyseBundleStatements, buildBundle }, { createProjectArchive, openProjectArchive }, { bundleArrangementFromLegacyOrder }] = await Promise.all([
    import("../app/lib/bundle-engine.ts"),
    import("../app/lib/project-archive.ts"),
    import("../app/lib/bundle-arrangement.ts"),
  ]);
  const statement = await statementFile();
  const evidence = await evidenceFiles();
  const progress = [];
  const analysis = await analyseBundleStatements([{ id: "capacity-statement", file: statement, witnessName: "Capacity Test", witnessInitials: "CT" }], evidence, (stage, detail) => progress.push(`${stage}:${detail ?? ""}`));
  assert.equal(analysis.candidates.length, COUNT);
  assert.equal(analysis.evidence.length, COUNT);
  assert.ok(progress.some((entry) => entry.includes(`Reading evidence files:${COUNT} of ${COUNT}`)));
  assert.ok(progress.some((entry) => entry.includes(`Comparing statement references:${COUNT} of ${COUNT}`)));

  const confirmed = analysis.candidates.map((candidate, index) => ({
    ...candidate,
    evidenceId: analysis.evidence[index].id,
    description: `Capacity exhibit ${String(index + 1).padStart(3, "0")}`,
    confirmed: true,
    included: true,
    confidence: 100,
    rationale: "Source selected by deterministic capacity fixture",
  }));
  const arrangement = bundleArrangementFromLegacyOrder(confirmed.map((candidate) => candidate.id));
  const build = await buildBundle({ ...analysis, candidates: confirmed }, confirmed, {
    arrangement,
    layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 85 },
  });
  assert.equal(build.records.length, COUNT);
  assert.ok((build.volumes?.length ?? 0) >= 3, "the capacity fixture is physically split into several volumes");
  assert.equal(new Set(build.records.map((record) => record.sourceHash)).size, COUNT);
  assert.equal(build.records.at(-1)?.description, "Capacity exhibit 200");

  const sources = [
    { id: "capacity-statement", role: "statement", name: statement.name, sha256: analysis.statementHash, file: statement },
    ...analysis.evidence.map((record) => ({ id: record.id, role: "evidence", name: record.name, sha256: record.sha256, file: record.file })),
  ];
  const archive = await createProjectArchive({
    schemaVersion: 8,
    name: "200 exhibit capacity fixture",
    createdAt: analysis.generatedAt,
    updatedAt: analysis.generatedAt,
    profileId: "exhibit-neutral",
    pagination: { matchPdfPageOrder: true, volumeNumbering: "continuous", scheme: "bundle", prefix: "", suffix: "", startAt: 1, padding: 0, preliminary: "arabic", countTemplates: true, position: "bottom-centre", fontSize: 8, includePrefixInIndex: false },
    layout: { includeDividerPages: false, includeExhibitCoverPages: false, countOptionalPagesInReferences: false, volumePageLimit: 85 },
    witnessSettings: { "capacity-statement": { initials: "CT", nextNumber: 1 } },
    candidates: confirmed,
    analysis: { statements: [{ id: "capacity-statement", witnessName: "Capacity Test", witnessInitials: "CT", name: statement.name }] },
    arrangement,
  }, sources);
  const reopened = await openProjectArchive(new File([archive], "Capacity.bundle-project", { type: "application/zip" }));
  assert.equal(reopened.sources.length, COUNT + 1);
  assert.equal(reopened.snapshot.arrangement?.nodes.length, COUNT);
});
