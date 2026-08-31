import assert from "node:assert/strict";
import { createHash, webcrypto } from "node:crypto";
import { File } from "node:buffer";
import test from "node:test";

globalThis.File = File;
globalThis.crypto ??= webcrypto;

test("build reports real local assembly stages in order", async () => {
  const { buildBundle } = await import("../app/lib/bundle-engine.ts");
  const source = new File(["A local exhibit"], "Exhibit.txt", { type: "text/plain" });
  const evidence = {
    id: "evidence-1", name: source.name, file: source, extension: "txt", text: "A local exhibit", sha256: "a".repeat(64), pageCount: 0, rotationPages: [], marker: null,
  };
  const candidate = {
    id: "candidate-1", evidenceId: evidence.id, included: true, confirmed: true, description: "Local exhibit", mark: "AH 1", provisionalNumber: 1, sequenceOrder: 1,
    paragraph: 1, citation: "I refer to the local exhibit.", statementName: "Statement.docx", statementId: "statement-1", witnessInitials: "AH", exhibitInitials: "AH", exhibitSequence: 1,
  };
  const statement = new File(["read-only statement"], "Statement.docx");
  const statementHash = createHash("sha256").update(new Uint8Array(await statement.arrayBuffer())).digest("hex");
  const statementInput = { id: "statement-1", file: statement, witnessName: "Witness", witnessInitials: "AH" };
  const analysis = { caseTitle: "Local matter", statementName: statement.name, statementHash, statementId: statementInput.id, statementSources: [{ statementId: statementInput.id, fileName: statement.name, sha256: statementHash }], statementSnapshots: [statementInput], statementHandles: [statementInput], evidence: [evidence], candidates: [candidate], unreferenced: [], generatedAt: new Date().toISOString() };
  const stages = [];
  const result = await buildBundle(analysis, [candidate], { onProgress: (stage) => stages.push(stage) });
  assert.equal(result.records.length, 1);
  assert.deepEqual(stages, [
    "Checking the confirmed exhibits",
    "Preparing the exhibit pages",
    "Rendering exhibits",
    "Creating the authoritative build plan",
    "Applying page numbers, bookmarks and links",
    "Reopening and validating the finished PDF",
    "Bundle complete",
  ]);
});
