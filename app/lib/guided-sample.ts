import {
  SAMPLE_STATEMENT,
  type AnalysisResult,
} from "./bundle-engine.ts";

export const GUIDED_SAMPLE_MAPPED_FILES = [
  "01_SAMPLE_Agreement.pdf",
  "02_SAMPLE_Invoice.pdf",
  "03_SAMPLE_Project_Report.docx",
  "04_SAMPLE_Claimant_Email.eml",
  "05_SAMPLE_Cost_Workbook.xlsx",
  "01_SAMPLE_Agreement.pdf",
] as const;

export const GUIDED_SAMPLE_WORKBOOK = "05_SAMPLE_Cost_Workbook.xlsx";

export const GUIDED_SAMPLE_MAPPING_RATIONALE =
  "Guided sample mapping: this obvious source is supplied so the user can practise the approval step.";

export function hasGuidedSampleEvidence(evidenceNames: string[]) {
  const names = new Set(evidenceNames);
  return GUIDED_SAMPLE_MAPPED_FILES.every((name) => names.has(name));
}

export function isGuidedSampleSelection(statementName: string, evidenceNames: string[]) {
  return statementName === SAMPLE_STATEMENT && hasGuidedSampleEvidence(evidenceNames);
}

export function applyGuidedSampleMapping(result: AnalysisResult): AnalysisResult {
  const evidence = result.evidence.map((record) => {
    if (record.name !== GUIDED_SAMPLE_WORKBOOK) return record;
    return {
      ...record,
      sheetSelections: record.sheetSelections?.map((selection, index) => ({
        ...selection,
        included: index === 0,
      })),
    };
  });
  const candidates = result.candidates.map((candidate, index) => {
    const mapped = evidence.find((record) => record.name === GUIDED_SAMPLE_MAPPED_FILES[index]);
    if (!mapped) return candidate;
    return {
      ...candidate,
      evidenceId: mapped.id,
      confidence: 100,
      rationale: GUIDED_SAMPLE_MAPPING_RATIONALE,
      repeatDecision: index === 5 ? "same" as const : candidate.repeatDecision,
    };
  });
  return { ...result, evidence, candidates };
}

export function analysisWithGuidedMapping(
  result: AnalysisResult,
  statementName: string,
  evidenceNames: string[],
) {
  if (!isGuidedSampleSelection(statementName, evidenceNames)) return result;
  return applyGuidedSampleMapping(result);
}
