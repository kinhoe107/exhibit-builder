import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFRef,
  type PDFObject,
} from "pdf-lib";

export type UnsafePdfAction = {
  page: number;
  location: "page" | "annotation";
  action: string;
};

const ACTION_TYPE = PDFName.of("S");
const NEXT_ACTION = PDFName.of("Next");

function resolvedObject(document: PDFDocument, value: PDFObject | undefined) {
  if (!value) return undefined;
  return value instanceof PDFRef ? document.context.lookup(value) : value;
}

function actionName(value: PDFObject | undefined) {
  return value instanceof PDFName ? value.toString().replace(/^\//, "") : "Unknown";
}

/**
 * Inspect an action or additional-actions object through direct and indirect
 * dictionaries and arrays. The only action retained by Exhibit Builder is an
 * internal GoTo destination; every other action can cause external activity,
 * execute viewer JavaScript, submit data or otherwise change document state.
 */
function inspectActionObject(
  document: PDFDocument,
  value: PDFObject | undefined,
  page: number,
  location: UnsafePdfAction["location"],
  findings: UnsafePdfAction[],
  visited: Set<PDFObject>,
) {
  const resolved = resolvedObject(document, value);
  if (!resolved || visited.has(resolved)) return;
  visited.add(resolved);

  if (resolved instanceof PDFArray) {
    for (let index = 0; index < resolved.size(); index += 1) {
      inspectActionObject(document, resolved.get(index), page, location, findings, visited);
    }
    return;
  }
  if (!(resolved instanceof PDFDict)) return;

  const subtype = resolvedObject(document, resolved.get(ACTION_TYPE));
  if (subtype) {
    const name = actionName(subtype);
    if (name !== "GoTo") findings.push({ page, location, action: name });
    inspectActionObject(document, resolved.get(NEXT_ACTION), page, location, findings, visited);
    return;
  }

  // /AA is an event-name dictionary whose values are actions. Walking every
  // value also handles indirect action arrays without relying on viewer output.
  for (const [, child] of resolved.entries()) {
    inspectActionObject(document, child, page, location, findings, visited);
  }
}

export function unsafePdfActions(document: PDFDocument): UnsafePdfAction[] {
  const findings: UnsafePdfAction[] = [];
  document.getPages().forEach((page, pageIndex) => {
    const pageNumber = pageIndex + 1;
    inspectActionObject(
      document,
      page.node.get(PDFName.of("AA")),
      pageNumber,
      "page",
      findings,
      new Set(),
    );
    const annotations = resolvedObject(document, page.node.get(PDFName.of("Annots")));
    if (!(annotations instanceof PDFArray)) return;
    for (let index = 0; index < annotations.size(); index += 1) {
      const annotation = resolvedObject(document, annotations.get(index));
      if (!(annotation instanceof PDFDict)) continue;
      inspectActionObject(
        document,
        annotation.get(PDFName.of("A")),
        pageNumber,
        "annotation",
        findings,
        new Set(),
      );
      inspectActionObject(
        document,
        annotation.get(PDFName.of("AA")),
        pageNumber,
        "annotation",
        findings,
        new Set(),
      );
    }
  });
  return findings;
}

export function assertPdfActionsSafe(document: PDFDocument, fileName: string) {
  const findings = unsafePdfActions(document);
  if (!findings.length) return;
  const details = findings
    .slice(0, 8)
    .map((finding) => `page ${finding.page} ${finding.location} /${finding.action}`)
    .join(", ");
  const remainder = findings.length > 8 ? ` and ${findings.length - 8} more` : "";
  throw new Error(`${fileName} contains active PDF actions (${details}${remainder}). Flatten or remove the actions in a trusted PDF application, then replace this source file.`);
}
