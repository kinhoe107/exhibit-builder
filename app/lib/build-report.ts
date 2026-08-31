import type { AnalysisResult, BuildResult, ExhibitCandidate } from "./bundle-engine.ts";
import type { BuildResolution, PreflightCheck } from "./bundle-types.ts";

export type BuildReportExhibit = {
  number?: number;
  description: string;
  documentDate: string | null;
  sourceFile: string;
  volumeNumber: number;
  physicalPdfPages: string;
  statementReferencePages: string;
  statementReferenceMark: string;
  citationStatus: string;
  statementReferences: Array<{ paragraph: number; volumeNumber: number; pageRange: string }>;
  emailAttachments?: Array<{ name: string; identity: string; sha256: string; parentSha256: string; disposition: string }>;
};

export type BuildReportReviewRow = {
  paragraph: number | null;
  included: boolean;
  confirmed: boolean;
  confirmationMethod: string | null;
  confirmedAt: string | null;
  confidence: number;
  rationale: string;
  description: string;
  documentDate: string;
  citationStatus: string;
  manualAddedAt: string | null;
  warningAcknowledgedAt: string | null;
  aliases: string[];
  note: string;
  sourceFile: string | null;
  sourceSha256: string | null;
  order: number;
  repeatDecision: string | null;
};

export type BuildReportPayload = {
  product: string;
  generatedAt: string;
  project: string;
  buildManifest: Record<string, unknown>;
  output: { fileName: string; sha256: string; pageCount: number };
  exhibits: BuildReportExhibit[];
  review: BuildReportReviewRow[];
  preflight: PreflightCheck[];
  validation: BuildResult["checks"];
  resolutions: BuildResolution[];
};

export function createBuildReportPayload(input: {
  generatedAt?: string;
  projectName: string;
  build: BuildResult;
  candidates: ExhibitCandidate[];
  analysis?: AnalysisResult | null;
  preflight: PreflightCheck[];
  resolutions: BuildResolution[];
}): BuildReportPayload {
  const { build, candidates, analysis, preflight, resolutions, projectName } = input;
  return {
    product: "Exhibit Builder",
    generatedAt: input.generatedAt ?? new Date().toISOString(),
    project: projectName,
    buildManifest: build.manifest,
    output: { fileName: build.fileName, sha256: build.sha256, pageCount: build.pageCount },
    exhibits: build.records.map((record) => ({
      number: record.exhibitNumber,
      description: record.description,
      documentDate: record.documentDate ?? null,
      sourceFile: record.fileName,
      volumeNumber: record.volumeNumber ?? 1,
      physicalPdfPages: `${record.startPage}-${record.endPage}`,
      statementReferencePages: record.exhibitPageLabelStart === record.exhibitPageLabelEnd ? record.exhibitPageLabelStart ?? "" : `${record.exhibitPageLabelStart}-${record.exhibitPageLabelEnd}`,
      statementReferenceMark: record.mark,
      citationStatus: record.citationStatus ?? "cited",
      statementReferences: record.statementReferences.map((reference) => ({
        paragraph: reference.paragraph,
        volumeNumber: reference.volumeNumber ?? record.volumeNumber ?? 1,
        pageRange: reference.exhibitPageLabelStart === reference.exhibitPageLabelEnd ? reference.exhibitPageLabelStart ?? "" : `${reference.exhibitPageLabelStart}-${reference.exhibitPageLabelEnd}`,
      })),
      ...(record.emailAttachments?.length ? { emailAttachments: record.emailAttachments } : {}),
    })),
    review: candidates.map((candidate) => ({
      paragraph: candidate.manualAddition ? null : candidate.paragraph,
      included: candidate.included,
      confirmed: candidate.confirmed,
      confirmationMethod: candidate.confirmationMethod ?? null,
      confirmedAt: candidate.confirmedAt ?? null,
      confidence: candidate.confidence,
      rationale: candidate.rationale,
      description: candidate.description,
      documentDate: candidate.date,
      citationStatus: candidate.manualAddition ? "not-cited-manual-addition" : "cited",
      manualAddedAt: candidate.manualAddedAt ?? null,
      warningAcknowledgedAt: candidate.manualWarningAcknowledgedAt ?? null,
      aliases: candidate.aliases ?? [],
      note: candidate.reviewNote ?? "",
      sourceFile: analysis?.evidence.find((record) => record.id === candidate.evidenceId)?.name ?? null,
      sourceSha256: analysis?.evidence.find((record) => record.id === candidate.evidenceId)?.sha256 ?? null,
      order: candidate.sequenceOrder ?? candidate.provisionalNumber,
      repeatDecision: candidate.repeatDecision ?? null,
    })),
    preflight,
    validation: build.checks,
    resolutions,
  };
}

function line(label: string, value: string | number | null | undefined) {
  return `${label}: ${value === null || value === undefined || value === "" ? "—" : value}`;
}

function asList(manifest: Record<string, unknown>, key: string) {
  return Array.isArray(manifest[key]) ? manifest[key] as Array<Record<string, unknown>> : [];
}

function textField(value: unknown, fallback = "—") {
  if (typeof value === "string" && value) return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return fallback;
}

export function formatBuildReportText(payload: BuildReportPayload) {
  const manifest = payload.buildManifest ?? {};
  const statement = (manifest.statement ?? {}) as { fileName?: string; sha256?: string };
  const volumes = payload.output.fileName.toLowerCase().endsWith(".zip") || asList(manifest, "volumes").length
    ? asList(manifest, "volumes")
    : [];
  const omitted = asList(manifest, "omittedCitations");
  const excluded = asList(manifest, "excludedFiles");
  const manuals = payload.exhibits.filter((exhibit) => exhibit.citationStatus === "not-cited-manual-addition");
  const manifestExhibits = asList(manifest, "exhibits");
  const sourceHashFor = (exhibit: BuildReportExhibit) => {
    const fromManifest = manifestExhibits.find((item) => item.fileName === exhibit.sourceFile && item.description === exhibit.description);
    return (fromManifest?.sourceHash as string | undefined) ?? payload.review.find((row) => row.sourceFile === exhibit.sourceFile)?.sourceSha256;
  };
  const warnings = payload.validation.filter((check) => check.status !== "pass");
  const lines = [
    "Exhibit Builder readable build report",
    line("Generated", payload.generatedAt),
    line("Project", payload.project),
    line("Output file", payload.output.fileName),
    line("Output SHA-256", payload.output.sha256),
    line("Output pages", payload.output.pageCount),
    line("Statement", statement.fileName ?? "—"),
    line("Statement SHA-256", statement.sha256 ?? "—"),
    "",
  ];
  if (volumes.length) {
    lines.push("Volumes");
    for (const volume of volumes) {
      lines.push(`- ${textField(volume.label, `Volume ${textField(volume.number, "?")}`)}: ${textField(volume.fileName, "")} (${textField(volume.pageCount, "?")} pages); SHA-256 ${textField(volume.sha256)}`);
    }
    lines.push("");
  }
  lines.push("Exhibits");
  for (const exhibit of payload.exhibits) {
    lines.push(`- Exhibit ${exhibit.number ?? exhibit.statementReferenceMark}: ${exhibit.description}`);
    lines.push(`  ${line("Source", exhibit.sourceFile)}`);
    lines.push(`  ${line("Source SHA-256", sourceHashFor(exhibit))}`);
    lines.push(`  ${line("Document date", exhibit.documentDate)}`);
    lines.push(`  ${line("Volume", exhibit.volumeNumber)}`);
    lines.push(`  ${line("Printed PDF pages", exhibit.physicalPdfPages)}`);
    lines.push(`  ${line("Statement reference pages", exhibit.statementReferencePages)}`);
    lines.push(`  ${line("Citation status", exhibit.citationStatus)}`);
    if (exhibit.citationStatus === "not-cited-manual-addition") lines.push(`  ${line("Provenance", "Added by reviewer; not cited in the statement")}`);
    if (exhibit.emailAttachments?.length) {
      lines.push("  Email attachments:");
      for (const child of exhibit.emailAttachments) {
        lines.push(`  - ${child.name}; SHA-256 ${child.sha256}; ${child.disposition}`);
      }
    }
    if (exhibit.statementReferences.length) {
      lines.push(`  Statement references: ${exhibit.statementReferences.map((reference) => `paragraph ${reference.paragraph} (vol ${reference.volumeNumber}, ${reference.pageRange})`).join("; ")}`);
    }
  }
  lines.push("");
  lines.push("Match review provenance");
  const reviewed = payload.review.filter((row) => row.included && row.confirmed);
  if (!reviewed.length) {
    lines.push("- None recorded");
  } else {
    for (const row of reviewed) {
      const method = row.confirmationMethod === "bulk" ? "bulk" : row.confirmationMethod === "individual" ? "individually" : "recorded";
      const when = row.confirmedAt ? ` at ${row.confirmedAt}` : "";
      lines.push(`- ${row.paragraph ? `Paragraph ${row.paragraph}` : row.description}: ${row.description}. Source ${row.sourceFile ?? "—"}. Source SHA-256 ${row.sourceSha256 ?? "—"}. Comparison score ${row.confidence}/100. Rationale: ${row.rationale || "—"}. Confirmed ${method}${when}.`);
    }
  }
  lines.push("");
  if (manuals.length) {
    lines.push("Added exhibits (not cited in the statement)");
    for (const exhibit of manuals) {
      lines.push(`- ${exhibit.description} (${exhibit.sourceFile}); SHA-256 ${sourceHashFor(exhibit) ?? "—"}`);
    }
    lines.push("");
  }
  if (omitted.length) {
    lines.push("Omitted citations");
    for (const item of omitted) {
      lines.push(`- ${textField(item.description, textField(item.citation, textField(item.candidateId, "Omitted citation")))}${item.paragraph ? ` (paragraph ${textField(item.paragraph)})` : ""}`);
    }
    lines.push("");
  }
  if (excluded.length) {
    lines.push("Excluded files");
    for (const file of excluded) {
      lines.push(`- ${textField(file.fileName, "Unnamed file")}; SHA-256 ${textField(file.sha256)}; ${textField(file.reason, "Excluded")}`);
    }
    lines.push("");
  }
  lines.push("Warnings and exceptions");
  if (!warnings.length && !payload.resolutions.length) {
    lines.push("- None");
  } else {
    for (const check of warnings) lines.push(`- ${check.label}: ${check.detail}`);
    for (const resolution of payload.resolutions) {
      lines.push(`- ${resolution.fileName ?? "Project setting"}: ${resolution.action}${resolution.note ? ` (${resolution.note})` : ""}`);
    }
  }
  lines.push("");
  lines.push("Final preflight");
  if (!payload.preflight.length) lines.push("- None recorded");
  else {
    for (const check of payload.preflight) {
      lines.push(`- [${check.severity}] ${check.label}${check.fileName ? ` (${check.fileName})` : ""}: ${check.detail}`);
    }
  }
  return `${lines.join("\r\n")}\r\n`;
}
