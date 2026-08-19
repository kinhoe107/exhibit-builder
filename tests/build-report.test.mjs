import assert from "node:assert/strict";
import { File } from "node:buffer";
import test from "node:test";
import { createBuildReportPayload, formatBuildReportText } from "../app/lib/build-report.ts";

globalThis.File = File;

const sourceHash = "a".repeat(64);
const outputHash = "b".repeat(64);

function samplePayload(multiVolume = false) {
  return createBuildReportPayload({
    generatedAt: "2026-08-15T12:00:00.000Z",
    projectName: "Ridgeway exhibits",
    build: {
      bytes: new Uint8Array(),
      fileName: multiVolume ? "Ridgeway_Volumes.zip" : "Exhibit_Bundle.pdf",
      sha256: outputHash,
      pageCount: multiVolume ? 40 : 12,
      records: [
        {
          mark: "AH1",
          description: "Supply agreement",
          fileName: "agreement.pdf",
          startPage: 3,
          endPage: 6,
          statementParagraph: 2,
          statementReferences: [{ paragraph: 2, citation: "I refer to the agreement.", exhibitPageLabelStart: "AH-3", exhibitPageLabelEnd: "AH-6", volumeNumber: 1 }],
          exhibitNumber: 1,
          exhibitPageLabelStart: "AH-3",
          exhibitPageLabelEnd: "AH-6",
          volumeNumber: 1,
          sourceHash,
          citationStatus: "cited",
          documentDate: "8 August 2026",
        },
        {
          mark: "AH2",
          description: "Uncited checklist",
          fileName: "checklist.pdf",
          startPage: 7,
          endPage: 8,
          statementParagraph: null,
          statementReferences: [],
          exhibitNumber: 2,
          exhibitPageLabelStart: "AH-7",
          exhibitPageLabelEnd: "AH-8",
          volumeNumber: multiVolume ? 2 : 1,
          sourceHash: "c".repeat(64),
          manualAddition: true,
          citationStatus: "not-cited-manual-addition",
          documentDate: "Date not stated",
          manualAddedAt: "2026-08-15T11:00:00.000Z",
        },
      ],
      manifest: {
        statement: { fileName: "Statement.docx", sha256: "d".repeat(64), modified: false },
        exhibits: [
          { description: "Supply agreement", fileName: "agreement.pdf", sourceHash },
          { description: "Uncited checklist", fileName: "checklist.pdf", sourceHash: "c".repeat(64) },
        ],
        omittedCitations: [{ paragraph: 9, description: "Excluded letter", candidateId: "omitted-1" }],
        excludedFiles: [{ fileName: "draft.pdf", sha256: "e".repeat(64), reason: "No confirmed citation match" }],
        ...(multiVolume ? { volumes: [{ number: 1, label: "Volume 1", fileName: "Volume_1.pdf", sha256: "f".repeat(64), pageCount: 20 }, { number: 2, label: "Volume 2", fileName: "Volume_2.pdf", sha256: "g".repeat(64), pageCount: 20 }] } : {}),
      },
      checks: [
        { label: "Witness statement unchanged", status: "pass", detail: "The source statement was read only." },
        { label: "Cited material omitted by reviewer", status: "warning", detail: "1 cited reference intentionally excluded." },
      ],
    },
    candidates: [
      { id: "c1", mark: "AH 1", provisionalNumber: 1, description: "Supply agreement", date: "8 August 2026", paragraph: 2, citation: "I refer to the agreement.", citationResolution: "none", discoverySignals: [], evidenceId: "e1", confidence: 90, rationale: "Matched", included: true, confirmed: true },
      { id: "manual", mark: "EX 1", provisionalNumber: 2, description: "Uncited checklist", date: "Date not stated", paragraph: 0, citation: "", citationResolution: "none", discoverySignals: [], evidenceId: "e2", confidence: 100, rationale: "Added", included: true, confirmed: true, manualAddition: true, sequenceOrder: 2000 },
      { id: "omitted-1", mark: "AH 3", provisionalNumber: 3, description: "Excluded letter", date: "Date not stated", paragraph: 9, citation: "I refer to the letter.", citationResolution: "none", discoverySignals: [], evidenceId: null, confidence: 0, rationale: "Unmatched", included: false, confirmed: false },
    ],
    analysis: {
      statementName: "Statement.docx",
      statementHash: "d".repeat(64),
      caseTitle: "Ridgeway",
      candidates: [],
      evidence: [
        { id: "e1", file: new File(["a"], "agreement.pdf"), name: "agreement.pdf", extension: "pdf", text: "", marker: null, sha256: sourceHash, pageCount: 4, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" },
        { id: "e2", file: new File(["b"], "checklist.pdf"), name: "checklist.pdf", extension: "pdf", text: "", marker: null, sha256: "c".repeat(64), pageCount: 2, readableText: true, encrypted: false, rotationPages: [], ocrPages: [], ocrStatus: "not-needed" },
      ],
      unreferenced: [],
      statementWarnings: [],
      generatedAt: "2026-08-15T11:00:00.000Z",
    },
    preflight: [
      { id: "p1", severity: "pass", label: "Readable sources", detail: "All included sources could be read." },
    ],
    resolutions: [
      { blockerId: "ocr-1", action: "proceed-without-ocr", fileName: "scan.pdf", approvedAt: "2026-08-15T11:30:00.000Z", note: "Visual review completed." },
    ],
  });
}

test("technical JSON payload keeps the existing report shape", () => {
  const payload = samplePayload();
  assert.equal(payload.product, "Exhibit Builder");
  assert.deepEqual(Object.keys(payload.exhibits[0]).sort(), ["citationStatus", "description", "documentDate", "number", "physicalPdfPages", "sourceFile", "statementReferenceMark", "statementReferencePages", "statementReferences", "volumeNumber"]);
  assert.equal(payload.exhibits[1].citationStatus, "not-cited-manual-addition");
  assert.equal(payload.review[1].paragraph, null);
  assert.equal(payload.review[1].citationStatus, "not-cited-manual-addition");
});

test("readable UTF-8 report covers exhibits, hashes, manuals, exclusions, warnings and volumes", () => {
  const text = formatBuildReportText(samplePayload(true));
  assert.match(text, /Exhibit Builder readable build report/);
  assert.match(text, /Output SHA-256: b{64}/);
  assert.match(text, /Source SHA-256: a{64}/);
  assert.match(text, /Printed PDF pages: 3-6/);
  assert.match(text, /Citation status: not-cited-manual-addition/);
  assert.match(text, /Added exhibits \(not cited in the statement\)/);
  assert.match(text, /Omitted citations/);
  assert.match(text, /Excluded letter \(paragraph 9\)/);
  assert.match(text, /Excluded files/);
  assert.match(text, /draft\.pdf/);
  assert.match(text, /Cited material omitted by reviewer/);
  assert.match(text, /proceed-without-ocr/);
  assert.match(text, /Volume 1: Volume_1\.pdf/);
  assert.match(text, /Volume 2: Volume_2\.pdf/);
  assert.match(text, /\[pass\] Readable sources/);
  assert.equal(Buffer.from(text, "utf8").toString("utf8"), text);
});
