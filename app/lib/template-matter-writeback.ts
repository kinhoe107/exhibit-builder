import { rgb, type PDFFont, type PDFPage, type RGB } from "pdf-lib";
import type { TemplateMatterFinding, TemplateMatterPatch } from "./template-matter-review.ts";
import { writableMatterFindings, type TemplateMatterReview } from "./template-matter-review.ts";

export type TemplatePagePlacement = {
  pageIndex: number;
  sourcePageNumber: number;
  scale: number;
  offsetX: number;
  offsetY: number;
};

export type CoverTextAlignment = "left" | "center" | "right";

const COVER = rgb(1, 1, 1);
const INK = rgb(0.1, 0.14, 0.2);

function changedPatches(review: TemplateMatterReview, patches: TemplateMatterPatch[] | undefined) {
  if (!patches?.length) return [] as Array<{ finding: TemplateMatterFinding; value: string }>;
  const byId = new Map(writableMatterFindings(review).map((finding) => [finding.id, finding]));
  return patches.flatMap((patch) => {
    const finding = byId.get(patch.findingId);
    if (!finding?.geometry) return [];
    if (finding.value === patch.value) return [];
    return [{ finding, value: patch.value }];
  });
}

function mappedBox(finding: TemplateMatterFinding, placement: TemplatePagePlacement) {
  const geometry = finding.geometry;
  if (!geometry) return null;
  const pad = 1.2;
  const x = placement.offsetX + geometry.x * placement.scale - pad;
  const y = placement.offsetY + geometry.y * placement.scale - pad;
  const width = geometry.width * placement.scale + pad * 2;
  const height = Math.max(geometry.height, geometry.fontSize) * placement.scale + pad * 2;
  const size = geometry.fontSize * placement.scale;
  return { x, y, width, height, size };
}

function colourFor(finding: TemplateMatterFinding): RGB {
  const colour = finding.geometry?.color;
  if (!colour) return INK;
  return rgb(colour.r, colour.g, colour.b);
}

/** Infer the original run's alignment from leftover space on each side of the glyph box. */
export function inferCoverTextAlignment(box: { x: number; width: number }, pageWidth: number): CoverTextAlignment {
  const leftGap = box.x;
  const rightGap = pageWidth - (box.x + box.width);
  const slack = Math.max(24, pageWidth * 0.08);
  if (Math.abs(leftGap - rightGap) <= slack) return "center";
  if (rightGap <= 54 && leftGap > pageWidth * 0.35) return "right";
  return "left";
}

export function alignedCoverTextX(
  alignment: CoverTextAlignment,
  box: { x: number; width: number },
  textWidth: number,
) {
  if (alignment === "center") return box.x + box.width / 2 - textWidth / 2;
  if (alignment === "right") return box.x + box.width - textWidth;
  return box.x + 1;
}

export function applyTemplateMatterPatches(
  pages: PDFPage[],
  placements: TemplatePagePlacement[],
  review: TemplateMatterReview | undefined,
  patches: TemplateMatterPatch[] | undefined,
  fonts: { regular: PDFFont; bold: PDFFont },
) {
  if (!review) return;
  const pending = changedPatches(review, patches);
  if (!pending.length) return;
  const placementBySourcePage = new Map(placements.map((placement) => [placement.sourcePageNumber, placement]));
  for (const { finding, value } of pending) {
    const pageNumber = finding.geometry?.pageNumber ?? finding.pageNumbers[0];
    const placement = pageNumber ? placementBySourcePage.get(pageNumber) : undefined;
    if (!placement) continue;
    const page = pages[placement.pageIndex];
    const box = mappedBox(finding, placement);
    if (!page || !box) continue;
    const font = finding.kind === "party-name" || finding.kind === "matter-title" ? fonts.bold : fonts.regular;
    const alignment = inferCoverTextAlignment(box, page.getWidth());
    let size = box.size;
    while (size > 6 && font.widthOfTextAtSize(value, size) > box.width) size -= 0.4;
    const textWidth = value ? font.widthOfTextAtSize(value, size) : 0;
    page.drawRectangle({
      x: box.x,
      y: box.y,
      width: box.width,
      height: box.height,
      color: COVER,
    });
    if (!value) continue;
    page.drawText(value, {
      x: alignedCoverTextX(alignment, box, textWidth),
      y: box.y + 1,
      size,
      font,
      color: colourFor(finding),
    });
  }
}
