import type { PreflightCheck } from "./bundle-types.ts";

export type BuildReadinessInput = {
  includedCount: number;
  confirmedCount: number;
  pendingApprovalCount: number;
  templateReviewPending: boolean;
  unapprovedTemplateNames?: string[];
  preflight: PreflightCheck[];
};

export type BuildBlocker = {
  id: string;
  label: string;
  detail: string;
  kind: "approval" | "template" | "preflight" | "state";
  code?: string;
  fileName?: string;
  sourceId?: string;
  sourceSha256?: string;
  candidateId?: string;
  target: "review" | "sheets" | "templates" | "finalise";
  actionLabel: string;
};

/**
 * Turns the same conditions used by the disabled build button into concise,
 * user-facing reasons.  Approval state and technical preflight state are kept
 * separate so “0 approvals pending” cannot be mistaken for “ready to build”.
 */
export function buildBlockers(input: BuildReadinessInput): BuildBlocker[] {
  const blockers: BuildBlocker[] = [];
  const unconfirmedCount = Math.max(0, input.includedCount - input.confirmedCount);

  if (input.includedCount === 0) {
    blockers.push({
      id: "no-included-exhibits",
      label: "No exhibits included",
      detail: "Include at least one exhibit before building the exhibit bundle.",
      kind: "state",
      target: "review",
      actionLabel: "Return to exhibit review",
    });
  }

  if (unconfirmedCount > 0 || input.pendingApprovalCount > 0) {
    const count = Math.max(unconfirmedCount, input.pendingApprovalCount);
    blockers.push({
      id: "exhibit-approvals",
      label: "Exhibit approvals required",
      detail: `${count} included exhibit${count === 1 ? " is" : "s are"} still awaiting confirmation. Use the outstanding-approvals filter to locate ${count === 1 ? "it" : "them"}.`,
      kind: "approval",
      target: "review",
      actionLabel: `Review ${count} exhibit match${count === 1 ? "" : "es"}`,
    });
  }

  if (input.templateReviewPending) {
    const names = input.unapprovedTemplateNames?.filter(Boolean) ?? [];
    blockers.push({
      id: "template-approval",
      label: "Template review required",
      detail: names.length
        ? `Complete the outstanding template review${names.length === 1 ? "" : "s"}: ${names.join("; ")}.`
        : "Preview and confirm every selected custom template before building.",
      kind: "template",
      target: "templates",
      actionLabel: names.length === 1 ? "Review template" : "Review templates",
    });
  }

  for (const check of input.preflight.filter((item) => item.severity === "blocking")) {
    // The aggregate approval item above is clearer than repeating one line for
    // every ordinary unconfirmed card. Other blocking checks remain specific.
    if (check.label === "Unconfirmed exhibit") continue;
    const isWorkbook = check.code === "workbook.no_sheet" || check.code === "workbook.sheet_unreadable";
    const isTemplate = check.code?.startsWith("template.") ?? false;
    const isOcr = check.code?.startsWith("ocr.") ?? false;
    blockers.push({
      id: `preflight-${check.id}`,
      label: check.label,
      detail: check.detail,
      kind: "preflight",
      code: check.code,
      fileName: check.fileName,
      sourceId: check.sourceId,
      sourceSha256: check.sourceSha256,
      candidateId: check.candidateId,
      target: isWorkbook ? "sheets" : isTemplate ? "templates" : "review",
      actionLabel: isWorkbook
        ? `Choose sheets${check.fileName ? ` for ${check.fileName}` : ""}`
        : isTemplate
          ? "Review template"
          : isOcr
            ? `Review OCR issue${check.fileName ? ` for ${check.fileName}` : ""}`
            : check.candidateId
              ? "Open the exhibit card"
            : check.fileName
              ? `Review ${check.fileName}`
              : "Review this issue",
    });
  }

  return blockers;
}
