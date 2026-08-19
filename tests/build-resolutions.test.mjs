import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { File } from "node:buffer";
import { webcrypto } from "node:crypto";
import test from "node:test";

globalThis.File = File;
globalThis.crypto ??= webcrypto;

const root = new URL("./fixtures/core/", import.meta.url);

test("OCR exception is hash-bound and changes the build to an auditable warning", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const { applyBuildResolutions, templateFallbackSlots } = await import("../app/lib/build-resolutions.ts");
  const statement = new File([await readFile(new URL("01_Witness_Statement_Amelia_Hart.docx", root))], "01_Witness_Statement_Amelia_Hart.docx");
  const source = new File([await readFile(new URL("Evidence_Inbox/Executed_Supply_Agreement_2026-02-01.pdf", root))], "Executed_Supply_Agreement_2026-02-01.pdf");
  const analysis = await analyseFiles(statement, [source]);
  const evidence = analysis.evidence[0];
  const altered = {
    ...analysis,
    evidence: analysis.evidence.map((record) => record.id === evidence.id ? { ...record, ocrStatus: "unavailable", ocrPages: [] } : record),
  };
  const candidate = { ...altered.candidates[0], evidenceId: evidence.id, confirmed: true, included: true };

  await assert.rejects(
    buildBundle(altered, [candidate]),
    /OCR unavailable/i,
  );

  const resolution = {
    blockerId: `ocr-${evidence.id}`,
    checkCode: "ocr.unavailable",
    profileId: "exhibit-neutral",
    action: "proceed-without-ocr",
    sourceId: evidence.id,
    sourceSha256: evidence.sha256,
    fileName: evidence.name,
    approvedAt: new Date().toISOString(),
    note: "Visually reviewed for this regression test.",
    visualReviewConfirmed: true,
  };
  const built = await buildBundle(altered, [candidate], { resolutions: [resolution] });
  assert.equal(built.manifest.technicalExceptions.length, 1);
  assert.equal(built.manifest.technicalExceptions[0].sourceSha256, evidence.sha256);
  assert.ok(built.checks.some((check) => check.label === "Approved technical exceptions"));

  const stale = { ...resolution, sourceSha256: "0".repeat(64) };
  await assert.rejects(buildBundle(altered, [candidate], { resolutions: [stale] }), /OCR unavailable/i);

  const effective = applyBuildResolutions(
    [{ id: `ocr-${evidence.id}`, code: "ocr.unavailable", policy: "exception-eligible", severity: "blocking", label: "OCR unavailable", detail: "Needs OCR.", sourceId: evidence.id, sourceSha256: evidence.sha256, fileName: evidence.name }],
    [resolution],
  );
  assert.equal(effective[0].severity, "warning");
  assert.match(effective[0].detail, /without a tool-generated OCR text layer/i);
  const template = { slot: "cover", sha256: "template-hash" };
  assert.equal(templateFallbackSlots([{ ...resolution, action: "use-built-in-template", templateSlots: ["cover"], templateHashes: { cover: template.sha256 } }], [template]).has("cover"), true);
  assert.equal(templateFallbackSlots([{ ...resolution, action: "use-built-in-template", templateSlots: ["cover"] }], [template]).has("cover"), false);
  assert.equal(templateFallbackSlots([{ ...resolution, action: "use-built-in-template", templateSlots: ["cover"], templateHashes: { cover: "stale" } }], [template]).has("cover"), false);
  assert.equal(templateFallbackSlots([{ ...resolution, action: "use-built-in-template", templateSlots: ["cover"], templateHashes: { cover: template.sha256 } }], []).has("cover"), false);
});

test("technical resolutions cannot waive a missing source", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const statement = new File([await readFile(new URL("01_Witness_Statement_Amelia_Hart.docx", root))], "01_Witness_Statement_Amelia_Hart.docx");
  const source = new File([await readFile(new URL("Evidence_Inbox/Executed_Supply_Agreement_2026-02-01.pdf", root))], "Executed_Supply_Agreement_2026-02-01.pdf");
  const analysis = await analyseFiles(statement, [source]);
  const candidate = { ...analysis.candidates[0], evidenceId: null, confirmed: true, included: true };
  await assert.rejects(
    buildBundle(analysis, [candidate], { resolutions: [{ blockerId: candidate.id, action: "proceed-without-ocr", approvedAt: new Date().toISOString() }] }),
    /confirmed source file/i,
  );
});

test("source exclusions remain explicit in the build audit manifest", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const statement = new File([await readFile(new URL("01_Witness_Statement_Amelia_Hart.docx", root))], "01_Witness_Statement_Amelia_Hart.docx");
  const firstSource = new File([await readFile(new URL("Evidence_Inbox/Executed_Supply_Agreement_2026-02-01.pdf", root))], "Executed_Supply_Agreement_2026-02-01.pdf");
  const secondSource = new File([await readFile(new URL("Evidence_Inbox/Apex_Invoice_AC-7782.pdf", root))], "Apex_Invoice_AC-7782.pdf");
  const analysis = await analyseFiles(statement, [firstSource, secondSource]);
  const firstCandidate = { ...analysis.candidates[0], evidenceId: analysis.evidence[0].id, included: true, confirmed: true };
  const secondCandidate = { ...analysis.candidates[9], evidenceId: analysis.evidence[1].id, included: true, confirmed: true };
  const resolution = {
    action: "exclude-source",
    sourceId: analysis.evidence[0].id,
    sourceSha256: analysis.evidence[0].sha256,
    approvedAt: new Date().toISOString(),
    note: "Reviewer excluded this source for the audit test.",
  };
  const build = await buildBundle(analysis, [firstCandidate, secondCandidate], { resolutions: [resolution] });
  assert.equal(build.records.length, 1);
  assert.equal(build.manifest.omittedCitations[0].sourceSha256, analysis.evidence[0].sha256);
  assert.equal(build.manifest.omittedCitations[0].decisionAction, "exclude-source");
  assert.equal(build.manifest.excludedFiles.find((item) => item.sha256 === analysis.evidence[0].sha256).reason, resolution.note);
  assert.equal(build.manifest.output.pageSize, "A4");
  assert.equal(build.manifest.output.orientation.nonA4, 0);
});

test("candidate exclusion reasons do not leak onto unrelated excluded files", async () => {
  const { analyseFiles, buildBundle } = await import("../app/lib/bundle-engine.ts");
  const statement = new File([await readFile(new URL("01_Witness_Statement_Amelia_Hart.docx", root))], "01_Witness_Statement_Amelia_Hart.docx");
  const selectedSource = new File([await readFile(new URL("Evidence_Inbox/Executed_Supply_Agreement_2026-02-01.pdf", root))], "Executed_Supply_Agreement_2026-02-01.pdf");
  const excludedSource = new File([await readFile(new URL("Evidence_Inbox/Apex_Invoice_AC-7782.pdf", root))], "Apex_Invoice_AC-7782.pdf");
  const unrelatedSource = new File([await readFile(new URL("Evidence_Inbox/PO_NRL-1047.pdf", root))], "PO_NRL-1047.pdf");
  const analysis = await analyseFiles(statement, [selectedSource, excludedSource, unrelatedSource]);
  const selectedCandidate = { ...analysis.candidates[0], evidenceId: analysis.evidence[0].id, included: true, confirmed: true };
  const excludedCandidate = { ...analysis.candidates[9], evidenceId: analysis.evidence[1].id, included: true, confirmed: true };
  const resolution = {
    action: "exclude-candidate",
    candidateId: excludedCandidate.id,
    sourceId: analysis.evidence[1].id,
    sourceSha256: analysis.evidence[1].sha256,
    approvedAt: new Date().toISOString(),
    note: "Reviewer excluded this cited item for the audit test.",
  };
  const build = await buildBundle(analysis, [selectedCandidate, excludedCandidate], { resolutions: [resolution] });
  assert.equal(build.manifest.excludedFiles.find((item) => item.sha256 === analysis.evidence[1].sha256).reason, resolution.note);
  assert.equal(build.manifest.excludedFiles.find((item) => item.sha256 === analysis.evidence[2].sha256).reason, "No confirmed citation match");
});
