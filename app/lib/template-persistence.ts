import type { ProjectSnapshot, ProjectSource, TemplateFile, TemplateSlot } from "./bundle-types.ts";

/**
 * Restores a custom template only with the exact rendered artifact that was
 * reviewed. Confirmation state is deliberately discarded when either source
 * hash differs, the rendered PDF is missing, or the review names another PDF.
 */
export function restoreProjectTemplates(snapshot: ProjectSnapshot, sources: ProjectSource[]): TemplateFile[] {
  return sources.filter((source) => source.role === "template").map((source) => {
    const slot = source.id.replace("template-", "") as TemplateSlot;
    const metadata = snapshot.templateReviews?.find((review) => review.slot === slot && review.sourceId === source.id && review.sourceSha256 === source.sha256);
    const sourceFormat = metadata?.sourceFormat ?? (source.file.name.split(".").pop()?.toLowerCase() ?? "pdf") as "pdf" | "docx" | "doc";
    const rendered = sourceFormat === "pdf"
      ? source
      : metadata?.renderedSourceId
        ? sources.find((item) => item.id === metadata.renderedSourceId && item.role === "template-rendered")
        : undefined;
    const exactReview = metadata
      && rendered?.sha256 === metadata.pdfSha256
      && metadata.reviewState.matterReview?.pdfSha256 === metadata.pdfSha256
      ? metadata.reviewState
      : undefined;
    return {
      slot,
      file: source.file,
      sha256: source.sha256,
      sourceFormat,
      pdfFile: rendered?.file,
      pdfSha256: rendered?.sha256,
      reviewState: exactReview,
    };
  });
}
