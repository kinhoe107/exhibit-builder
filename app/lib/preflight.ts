import type { AnalysisResult, ExhibitCandidate } from "./bundle-engine.ts";
import type { BundleProfile, PreflightCheck } from "./bundle-types.ts";
import { deriveExhibitGroups } from "./exhibit-groups.ts";
import { emailChildrenForDisposition, unresolvedEmailAttachments } from "./email-attachments.ts";
import { formatUnsafePdfAction } from "./pdf-action-safety.ts";
import { workbookPlanCheckCopy } from "./workbook-print-copy.ts";

function countWords(value: string) {
  return value.trim().split(/\s+/).filter(Boolean).length;
}

function fingerprint(value: string) {
  const words = value.toLowerCase().replace(/[^a-z0-9 ]/g, " ").split(/\s+/).filter((word) => word.length > 3);
  return new Set(words.slice(0, 2_000));
}

function overlap(left: Set<string>, right: Set<string>) {
  if (!left.size || !right.size) return 0;
  let common = 0;
  for (const word of left) if (right.has(word)) common += 1;
  return common / Math.min(left.size, right.size);
}

function pushWorkbookPlanChecks(
  checks: PreflightCheck[],
  options: {
    idPrefix: string;
    sheetName: string;
    warnings: string[];
    fileName: string;
    sourceId: string;
    sourceSha256?: string;
    candidateId?: string;
  },
) {
  for (const copy of workbookPlanCheckCopy(options.sheetName, options.warnings)) {
    checks.push({
      id: `${options.idPrefix}-${copy.idSuffix}`,
      severity: copy.blocking ? "blocking" : "warning",
      label: copy.label,
      detail: copy.detail,
      fileName: options.fileName,
      sourceId: options.sourceId,
      sourceSha256: options.sourceSha256,
      candidateId: options.candidateId,
    });
  }
}

function unmatchedExhibitSubject(candidate: ExhibitCandidate) {
  if (candidate.manualAddition || !Number.isInteger(candidate.paragraph) || candidate.paragraph <= 0) {
    const description = candidate.description?.trim();
    return description || "This added exhibit";
  }
  const count = candidate.citationCount ?? 1;
  if (count > 1 && candidate.citationOrdinal) {
    return `Paragraph ${candidate.paragraph}, reference ${candidate.citationOrdinal} of ${count},`;
  }
  return `Paragraph ${candidate.paragraph}`;
}

export function runPreflight(
  analysis: AnalysisResult,
  candidates: ExhibitCandidate[],
  profile: BundleProfile,
): PreflightCheck[] {
  const checks: PreflightCheck[] = [];
  const included = candidates.filter((candidate) => candidate.included);
  const groups = deriveExhibitGroups(analysis, candidates);
  const groupByCandidateId = new Map(
    groups.flatMap((group) => group.collisionMembers.map((member) => [member.id, group] as const)),
  );
  for (const candidate of included) {
    if (!candidate.evidenceId) {
      checks.push({ id: `unmatched-${candidate.id}`, severity: "blocking", label: "Unmatched exhibit", detail: `${unmatchedExhibitSubject(candidate)} has no selected source file.`, candidateId: candidate.id });
    }
    if (profile.requireConfirmedExhibits && !candidate.confirmed) {
      checks.push({ id: `unconfirmed-${candidate.id}`, severity: "blocking", label: "Unconfirmed exhibit", detail: `${candidate.mark} has not been confirmed by a reviewer.` });
    }
    const record = candidate.evidenceId ? analysis.evidence.find((item) => item.id === candidate.evidenceId) : undefined;
    const group = groupByCandidateId.get(candidate.id);
    const hiddenSameEmailRepeat = candidate.repeatDecision === "same" && group?.canonical.id !== candidate.id;
    if (record?.emailAttachments?.length && !hiddenSameEmailRepeat) {
      for (const child of unresolvedEmailAttachments(record.emailAttachments, candidate.emailAttachmentDispositions)) {
        checks.push({
          id: `email-attachment-${candidate.id}-${child.identity}`,
          code: "email.attachment_unresolved",
          severity: "blocking",
          label: "Email attachment needs a decision",
          detail: `${record.name} includes ${child.name}. Choose Print with this email, Add as its own exhibit, or Leave out.`,
          fileName: record.name,
          sourceId: record.id,
          sourceSha256: record.sha256,
          candidateId: candidate.id,
        });
      }
      for (const child of record.emailAttachments) {
        const disposition = candidate.emailAttachmentDispositions?.[child.identity];
        if (!child.supported && disposition && disposition !== "leave-out") {
          checks.push({
            id: `email-attachment-unsupported-${candidate.id}-${child.identity}`,
            code: "email.attachment_unsupported",
            severity: "blocking",
            label: "Unsupported email attachment",
            detail: `${child.name} cannot be converted. Leave it out.`,
            fileName: record.name,
            sourceId: record.id,
            sourceSha256: record.sha256,
            candidateId: candidate.id,
          });
        }
        if (disposition === "print-with-email" && child.extension === "xlsx") {
          const workbook = child.workbook;
          if (!workbook) {
            checks.push({
              id: `email-xlsx-unreadable-${candidate.id}-${child.identity}`,
              severity: "blocking",
              label: "Unreadable workbook",
              detail: `${child.name} could not be analysed for printing with ${record.name}.`,
              fileName: child.name,
              sourceId: child.identity,
              sourceSha256: child.sha256,
              candidateId: candidate.id,
            });
          } else {
            const picks = child.sheetSelections ?? [];
            if (!picks.some((pick) => pick.included)) {
              checks.push({
                id: `email-xlsx-none-${candidate.id}-${child.identity}`,
                severity: "blocking",
                label: "No worksheet selected",
                detail: `Select at least one worksheet for ${child.name} before printing it with ${record.name}.`,
                fileName: child.name,
                sourceId: child.identity,
                sourceSha256: child.sha256,
                candidateId: candidate.id,
              });
            }
            for (const sheet of workbook.sheets.filter((item) => picks.find((pick) => pick.name === item.name)?.included)) {
              if (!sheet.cells.length) {
                checks.push({
                  id: `email-xlsx-empty-${candidate.id}-${child.identity}-${sheet.name}`,
                  severity: "blocking",
                  label: "Unreadable worksheet",
                  detail: `${child.name} / ${sheet.name} has no readable cells in its selected range.`,
                  fileName: child.name,
                  sourceId: child.identity,
                  sourceSha256: child.sha256,
                  candidateId: candidate.id,
                });
              }
              pushWorkbookPlanChecks(checks, {
                idPrefix: `email-xlsx-plan-${candidate.id}-${child.identity}-${sheet.name}`,
                sheetName: sheet.name,
                warnings: sheet.renderPlan.warnings,
                fileName: child.name,
                sourceId: child.identity,
                sourceSha256: child.sha256,
                candidateId: candidate.id,
              });
            }
          }
        }
      }
    }
  }
  const duplicateMarks = new Set<string>();
  const outputMarks = new Set<string>();
  for (const group of groups) {
    // AH1/RC1 is a bundle identifier and correctly repeats for each distinct
    // exhibit.  Only the numeric index item must be unique within that bundle.
    const mark = `${group.outputMark.replace(/\s+/g, "").toUpperCase()}:${group.exhibitNumber}`;
    if (outputMarks.has(mark)) duplicateMarks.add(mark);
    outputMarks.add(mark);
  }
  for (const mark of duplicateMarks) {
    checks.push({ id: `duplicate-${mark}`, severity: "blocking", label: "Duplicate index number", detail: `${mark} is assigned more than once.` });
  }
  for (const group of groups) {
    if (group.decisionPending) checks.push({ id: `repeat-pending-${group.sourceHash}`, severity: "blocking", label: "Repeat source decision required", detail: `${group.members.map((member) => `paragraph ${member.paragraph}`).join(" and ")} select identical file content. Confirm whether they are the same exhibit or separate exhibits.` });
    if (group.selectionConflict) checks.push({ id: `repeat-selection-${group.sourceHash}`, severity: "blocking", label: "Conflicting repeat selection", detail: `${group.canonical.mark} has conflicting page ranges or worksheet selections across citations. Align them before confirming one exhibit.` });
    const segmentExtensions = new Set(group.members.map((member) => analysis.evidence.find((record) => record.id === member.evidenceId)?.extension).filter(Boolean));
    const segmentHashes = new Set(group.members.map((member) => analysis.evidence.find((record) => record.id === member.evidenceId)?.sha256).filter(Boolean));
    if (segmentHashes.size > 1 && (segmentExtensions.size !== 1 || !segmentExtensions.has("pdf"))) checks.push({ id: `segment-unsupported-${group.id}`, severity: "blocking", label: "Mixed exhibit segments need review", detail: `${group.outputMark} has multiple source documents but only all-PDF cumulative segments are currently safe to build.` });
  }
  // Document-level checks cover every source segment entering the bundle, then
  // deduplicate identical content by SHA-256 so repeated citations get one
  // source-wide decision/control rather than duplicate UI rows.
  const selectedRecords = new Map<string, (typeof analysis.evidence)[number]>();
  for (const group of groups) {
    for (const member of group.members) {
      const record = analysis.evidence.find((item) => item.id === member.evidenceId);
      if (record) selectedRecords.set(record.sha256, record);
    }
  }
  for (const record of selectedRecords.values()) {
    const documentCheckStart = checks.length;
    // This is advisory only.  A large source is common in legal work, but
    // signalling it up front prevents a legitimate local OCR/build from
    // being mistaken for a frozen application.
    if (record.file.size >= 100 * 1024 * 1024) {
      checks.push({
        id: `large-source-${record.id}`,
        severity: "warning",
        label: "Large source file",
        detail: `${Math.ceil(record.file.size / (1024 * 1024))} MB. Local processing may take longer; keep Exhibit Builder open until final validation completes.`,
        fileName: record.name,
      });
    }
    if (record.extension === "xlsx") {
      if (!record.workbook) checks.push({ id: `xlsx-unreadable-${record.id}`, severity: "blocking", label: "Unreadable workbook", detail: "Workbook analysis was not available.", fileName: record.name });
      else {
        const picks = record.sheetSelections ?? [];
        if (!picks.some((pick) => pick.included)) checks.push({ id: `xlsx-none-${record.id}`, severity: "blocking", label: "No worksheet selected", detail: "Select at least one visible worksheet to include.", fileName: record.name });
        for (const sheet of record.workbook.sheets.filter((item) => picks.find((pick) => pick.name === item.name)?.included)) if (!sheet.cells.length) checks.push({ id: `xlsx-empty-${record.id}-${sheet.name}`, severity: "blocking", label: "Unreadable worksheet", detail: `${sheet.name} has no readable cells in its selected range.`, fileName: record.name });
        for (const sheet of record.workbook.sheets) {
          const selected = picks.find((pick) => pick.name === sheet.name)?.included;
          if (sheet.state !== "visible" && !selected) {
            checks.push({ id: `xlsx-hidden-${record.id}-${sheet.name}`, severity: "warning", label: "Hidden worksheet not selected", detail: `${sheet.name} is ${sheet.state} and is not selected.`, fileName: record.name });
          }
          if (selected) {
            pushWorkbookPlanChecks(checks, {
              idPrefix: `xlsx-plan-${record.id}-${sheet.name}`,
              sheetName: sheet.name,
              warnings: sheet.renderPlan.warnings,
              fileName: record.name,
              sourceId: record.id,
              sourceSha256: record.sha256,
            });
          }
        }
      }
    }
    if (record.encrypted) {
      checks.push({ id: `encrypted-${record.id}`, severity: "blocking", label: "Encrypted PDF", detail: "Remove password protection before building.", fileName: record.name });
    }
    if (record.extension === "pdf" && record.ocrStatus === "completed") {
      checks.push({ id: `ocr-complete-${record.id}`, severity: "pass", label: "OCR completed locally", detail: "A searchable text layer was generated at 300dpi.", fileName: record.name });
    }
    if (record.extension === "pdf" && (record.ocrStatus === "failed" || record.ocrStatus === "unavailable")) {
      checks.push({
        id: `ocr-${record.id}`,
        severity: profile.requireOcr ? "blocking" : "warning",
        label: record.ocrStatus === "failed" ? "OCR failed" : "OCR unavailable",
        detail: record.ocrStatus === "failed"
          ? `Local OCR could not create usable searchable text for this PDF.${record.ocrFailureReason ? ` ${record.ocrFailureReason}` : ""}`
          : "This PDF needs local OCR, which is unavailable in this environment.",
        fileName: record.name,
      });
    }
    const printedPdfChildren = new Map<string, NonNullable<(typeof record)["emailAttachments"]>[number]>();
    const printedDocxChildren = new Map<string, NonNullable<(typeof record)["emailAttachments"]>[number]>();
    for (const candidate of included) {
      if (candidate.evidenceId !== record.id) continue;
      for (const child of emailChildrenForDisposition(record.emailAttachments, candidate.emailAttachmentDispositions, "print-with-email")) {
        if (child.extension === "pdf") printedPdfChildren.set(child.identity, child);
        if (child.extension === "docx") printedDocxChildren.set(child.identity, child);
      }
    }
    for (const child of printedPdfChildren.values()) {
      if (child.encrypted) {
        checks.push({
          id: `encrypted-${record.id}-${child.identity}`,
          severity: "blocking",
          label: "Encrypted PDF",
          detail: `${child.name} (printed with ${record.name}) is password-protected. Remove the password, then replace the parent email.`,
          fileName: child.name,
          sourceId: child.identity,
          sourceSha256: child.sha256,
        });
      }
      if (child.unsafePdfActions?.length) {
        const actionSummary = child.unsafePdfActions
          .slice(0, 8)
          .map(formatUnsafePdfAction)
          .join(", ");
        checks.push({
          id: `active-pdf-actions-${record.id}-${child.identity}`,
          severity: "blocking",
          label: "Active PDF actions are not permitted",
          detail: `${child.name} (printed with ${record.name}): ${actionSummary}${child.unsafePdfActions.length > 8 ? ` and ${child.unsafePdfActions.length - 8} more` : ""}. Flatten or remove the actions in a trusted PDF application, then replace the parent email. Ordinary visual annotations remain supported; source-document internal navigation links are removed during assembly.`,
          fileName: child.name,
          sourceId: child.identity,
          sourceSha256: child.sha256,
        });
      }
      if (child.pdfAnalysisFailure) {
        checks.push({
          id: `email-child-unreadable-${record.id}-${child.identity}`,
          severity: "blocking",
          label: "PDF attachment could not be analysed",
          detail: `${child.name} (printed with ${record.name}) could not be analysed. Leave the attachment out, or replace the parent email with a readable PDF.`,
          fileName: child.name,
          sourceId: child.identity,
          sourceSha256: child.sha256,
        });
      }
      if (child.pageSizeMeasurementFailed) {
        checks.push({
          id: `email-child-page-size-${record.id}-${child.identity}`,
          severity: "warning",
          label: "Email attachment page size could not be measured",
          detail: `${child.name} (printed with ${record.name}) could not be measured, so non-A4 review is unavailable. The build still converts it unless you leave it out.`,
          fileName: child.name,
          sourceId: child.identity,
          sourceSha256: child.sha256,
        });
      }
      if (child.ocrStatus === "completed") {
        checks.push({
          id: `ocr-complete-${record.id}-${child.identity}`,
          severity: "pass",
          label: "OCR completed locally",
          detail: "A searchable text layer was generated at 300dpi.",
          fileName: child.name,
          sourceId: child.identity,
          sourceSha256: child.sha256,
        });
      }
      if (!child.encrypted && (child.ocrStatus === "failed" || child.ocrStatus === "unavailable")) {
        checks.push({
          id: `ocr-${record.id}-${child.identity}`,
          severity: profile.requireOcr ? "blocking" : "warning",
          label: child.ocrStatus === "failed" ? "OCR failed" : "OCR unavailable",
          detail: child.ocrStatus === "failed"
            ? `Local OCR could not create usable searchable text for ${child.name} (printed with ${record.name}).${child.ocrFailureReason ? ` ${child.ocrFailureReason}` : ""}`
            : `${child.name} (printed with ${record.name}) needs local OCR, which is unavailable in this environment.`,
          fileName: child.name,
          sourceId: child.identity,
          sourceSha256: child.sha256,
        });
      }
    }
    for (const child of printedDocxChildren.values()) {
      if (child.docxAnalysisFailure) {
        checks.push({
          id: `email-child-unreadable-docx-${record.id}-${child.identity}`,
          severity: "blocking",
          label: "Word attachment could not be analysed",
          detail: `${child.name} (printed with ${record.name}) could not be analysed. Leave the attachment out, or replace the parent email with a readable Word file.`,
          fileName: child.name,
          sourceId: child.identity,
          sourceSha256: child.sha256,
        });
      }
    }
    if (record.extension === "pdf" && record.pageCount > 0 && record.rotationPages.length) {
      checks.push({
        id: `rotation-${record.id}`,
        severity: "warning",
        label: "Rotated pages",
        detail: `${record.rotationPages.length} page(s) have a non-zero rotation and should be visually reviewed.`,
        fileName: record.name,
      });
    }
    const rotatedAnnotations = record.rotationPages.filter((page) => record.annotationPages?.includes(page));
    if (record.extension === "pdf" && rotatedAnnotations.length) {
      checks.push({
        id: `rotated-annotations-${record.id}`,
        severity: "blocking",
        label: "Rotated PDF annotations cannot be preserved safely",
        detail: `${rotatedAnnotations.length === 1 ? "Page" : "Pages"} ${rotatedAnnotations.join(", ")} ${rotatedAnnotations.length === 1 ? "combines" : "combine"} page rotation with PDF annotations. Flatten the rotation and annotations in a trusted PDF application, then replace this source file.`,
        fileName: record.name,
      });
    }
    if (record.extension === "pdf" && record.unsafePdfActions?.length) {
      const actionSummary = record.unsafePdfActions
        .slice(0, 8)
        .map(formatUnsafePdfAction)
        .join(", ");
      checks.push({
        id: `active-pdf-actions-${record.id}`,
        severity: "blocking",
        label: "Active PDF actions are not permitted",
        detail: `${actionSummary}${record.unsafePdfActions.length > 8 ? ` and ${record.unsafePdfActions.length - 8} more` : ""}. Flatten or remove the actions in a trusted PDF application, then replace this source file. Ordinary visual annotations remain supported; source-document internal navigation links are removed during assembly.`,
        fileName: record.name,
      });
    }
    if (record.extension === "pdf" && record.pageCount > 0 && countWords(record.text) < 3) {
      checks.push({ id: `sparse-${record.id}`, severity: "warning", label: "Very little extracted text", detail: "Check scan quality and legibility.", fileName: record.name });
    }
    if (record.text.trim().length === 0) {
      checks.push({ id: `blank-${record.id}`, severity: "warning", label: "Blank or unreadable document", detail: "No usable text was found; visually confirm this file is not blank.", fileName: record.name });
    }
    // Every document-level check receives its identity at construction time.
    // Do not recover identity from a filename or an encoded display/check ID.
    for (let index = documentCheckStart; index < checks.length; index += 1) {
      const check = checks[index];
      checks[index] = {
        ...check,
        sourceId: check.sourceId ?? record.id,
        sourceSha256: check.sourceSha256 ?? record.sha256,
      };
    }
  }
  const byHash = new Map<string, string[]>();
  for (const record of analysis.evidence) byHash.set(record.sha256, [...(byHash.get(record.sha256) ?? []), record.name]);
  for (const names of byHash.values()) {
    if (names.length > 1) checks.push({ id: `duplicate-${names.join("-")}`, severity: "warning", label: "Duplicate source files", detail: `Identical source hashes: ${names.join(", ")}.` });
  }
  // A deterministic, bounded local signal for documents that look alike but
  // are not byte-identical. It never combines or excludes anything itself.
  const comparable = analysis.evidence.filter((record) => record.text.trim().length > 80).slice(0, 60);
  const comparableFingerprints = new Map(comparable.map((record) => [record.id, fingerprint(record.text)]));
  for (let leftIndex = 0; leftIndex < comparable.length; leftIndex += 1) {
    const left = comparable[leftIndex];
    const leftFingerprint = comparableFingerprints.get(left.id)!;
    for (let rightIndex = leftIndex + 1; rightIndex < comparable.length; rightIndex += 1) {
      const right = comparable[rightIndex];
      if (left.sha256 === right.sha256) continue;
      const score = overlap(leftFingerprint, comparableFingerprints.get(right.id)!);
      if (score < 0.82) continue;
      checks.push({
        id: `near-duplicate-${left.id}-${right.id}`,
        severity: "warning",
        label: "Possible near-duplicate sources",
        detail: `${left.name} and ${right.name} have unusually similar extracted text (${Math.round(score * 100)}%). Review them before selecting both.`,
        relatedSourceIds: [left.id, right.id],
      });
    }
  }
  const selectedEvidenceIds = new Set(included.map((candidate) => candidate.evidenceId).filter(Boolean));
  const selectedHashes = new Set(analysis.evidence.filter((record) => selectedEvidenceIds.has(record.id)).map((record) => record.sha256));
  const excluded = analysis.evidence.filter((record) => !selectedEvidenceIds.has(record.id));
  const duplicateCopies = excluded.filter((record) => selectedHashes.has(record.sha256));
  if (duplicateCopies.length) {
    checks.push({ id: "duplicate-physical-copies", severity: "warning", label: "Unselected duplicate physical copies", detail: `${duplicateCopies.map((record) => record.name).join(", ")} contain identical content to a selected source and will remain outside the bundle.` });
  }
  if (excluded.length) {
    checks.push({ id: "unreferenced", severity: "warning", label: "Unreferenced evidence", detail: `${excluded.length} supplied file(s) are outside the confirmed bundle.` });
  }
  if (!checks.some((item) => item.severity === "blocking")) {
    checks.push({ id: "ready", severity: "pass", label: "Ready to build", detail: "No blocking preflight issue was found." });
  }
  // Source-specific resolutions are bound to stable local IDs and content
  // hashes, never a display filename. This prevents equal filenames from
  // inheriting another file's OCR or technical-exception approval.
  const codeForLabel = (check: PreflightCheck) => {
    if (check.label === "Email attachment needs a decision") return "email.attachment_unresolved";
    if (check.label === "Unsupported email attachment") return "email.attachment_unsupported";
    if (check.label === "PDF attachment could not be analysed") return "email.attachment_unreadable";
    if (check.label === "Word attachment could not be analysed") return "email.word_unreadable";
    if (check.label === "Unmatched exhibit") return "exhibit.unmatched";
    if (check.label === "Unconfirmed exhibit") return "exhibit.unconfirmed";
    if (check.label === "Duplicate index number") return "numbering.duplicate";
    if (check.label === "Repeat source decision required") return "repeat.decision_pending";
    if (check.label === "Conflicting repeat selection") return "repeat.selection_conflict";
    if (check.label === "Mixed exhibit segments need review") return "segment.unsupported";
    if (check.label === "Unreadable workbook") return "workbook.unreadable";
    if (check.label === "No worksheet selected") return "workbook.no_sheet";
    if (check.label === "Unreadable worksheet") return "workbook.sheet_unreadable";
    if (check.label === "Encrypted PDF") return "source.encrypted";
    if (check.label === "OCR failed") return "ocr.failed";
    if (check.label === "OCR unavailable") return "ocr.unavailable";
    if (check.label === "Rotated pages") return "source.rotated";
    if (check.label === "Large source file") return "source.large";
    if (check.label === "Very little extracted text") return "source.sparse";
    if (check.label === "Blank or unreadable document") return "source.blank";
    if (check.label === "Hidden worksheet not selected") return "workbook.hidden_sheet";
    if (check.label === "Workbook print fidelity failed") return "workbook.fidelity_failed";
    if (check.label === "This Excel sheet is too wide to print on A4") return "workbook.fidelity_failed";
    if (check.label === "This Excel sheet is too tall to print on A4") return "workbook.fidelity_failed";
    if (check.label === "This Excel sheet cannot be printed as it stands") return "workbook.fidelity_failed";
    if (check.label === "Worksheet print note") return "workbook.render_warning";
    if (check.label === "Duplicate source files") return "source.duplicate";
    if (check.label === "Possible near-duplicate sources") return "source.near_duplicate";
    if (check.label === "Unselected duplicate physical copies") return "source.duplicate_copy";
    if (check.label === "Unreferenced evidence") return "source.unreferenced";
    return check.id === "ready" ? "build.ready" : `build.${check.id}`;
  };
  const policyFor = (check: PreflightCheck) => {
    if (check.label === "OCR failed" || check.label === "OCR unavailable") return "exception-eligible" as const;
    if (check.label === "Unmatched exhibit" || check.label === "Unconfirmed exhibit" || check.label === "Duplicate index number" || check.label === "Repeat source decision required" || check.label === "Conflicting repeat selection" || check.label === "Email attachment needs a decision") return "hard-legal" as const;
    if (check.severity === "blocking") return "hard-technical" as const;
    return "warning" as const;
  };
  return checks.map((check) => {
    if (check.sourceSha256 || !check.fileName) return { ...check, code: check.code ?? codeForLabel(check), policy: policyFor(check) };
    const source = check.sourceId
      ? analysis.evidence.find((record) => record.id === check.sourceId)
      : undefined;
    return source
      ? { ...check, code: check.code ?? codeForLabel(check), policy: policyFor(check), sourceId: check.sourceId ?? source.id, sourceSha256: check.sourceSha256 ?? source.sha256 }
      : { ...check, code: check.code ?? codeForLabel(check), policy: policyFor(check) };
  });
}

export function hasBlockingPreflight(checks: PreflightCheck[]) {
  return checks.some((check) => check.severity === "blocking");
}
