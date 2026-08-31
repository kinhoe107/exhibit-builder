import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFRef,
  PDFString,
  type PDFObject,
} from "pdf-lib";

export type UnsafePdfActionLocation =
  | "page"
  | "annotation"
  | "catalog"
  | "outline"
  | "form"
  | "javascript"
  | "document";

export type UnsafePdfAction = {
  page?: number;
  location: UnsafePdfActionLocation;
  action: string;
};

export const PDF_ACTION_INSPECTION_LIMIT = 8_000;

const ACTION_TYPE = PDFName.of("S");
const NEXT_ACTION = PDFName.of("Next");
const TYPE_NAME = PDFName.of("Type");
const OPEN_ACTION = PDFName.of("OpenAction");
const ADDITIONAL_ACTIONS = PDFName.of("AA");
const ACTION = PDFName.of("A");
const ANNOTS = PDFName.of("Annots");
const NAMES = PDFName.of("Names");
const JAVASCRIPT = PDFName.of("JavaScript");
const JS = PDFName.of("JS");
const KIDS = PDFName.of("Kids");
const OUTLINES = PDFName.of("Outlines");
const FIRST = PDFName.of("First");
const NEXT = PDFName.of("Next");
const ACRO_FORM = PDFName.of("AcroForm");
const FIELDS = PDFName.of("Fields");
const DESTINATION_OPERANDS: Record<string, { count: number; allowNull: boolean }> = {
  XYZ: { count: 3, allowNull: true },
  Fit: { count: 0, allowNull: false },
  FitH: { count: 1, allowNull: true },
  FitV: { count: 1, allowNull: true },
  FitR: { count: 4, allowNull: false },
  FitB: { count: 0, allowNull: false },
  FitBH: { count: 1, allowNull: true },
  FitBV: { count: 1, allowNull: true },
};
const STRUCTURAL_TYPES = new Set([
  "Page",
  "Pages",
  "Catalog",
  "Font",
  "XObject",
  "ExtGState",
  "Pattern",
  "Annot",
  "Outlines",
  "Metadata",
  "Group",
  "OCG",
  "Filespec",
  "EmbeddedFile",
  "ObjStm",
  "XRef",
  "Sig",
]);

type WalkMode = "action" | "actionOrDest" | "aaMap" | "nameTree" | "outline" | "field" | "jsEntry";
type VisitMap = Map<PDFObject, Set<WalkMode>>;

type WorkItem = {
  value: PDFObject | undefined;
  page?: number;
  location: UnsafePdfActionLocation;
  mode: WalkMode;
  visited: VisitMap;
};

type WalkContext = {
  document: PDFDocument;
  pageIdentities: Set<PDFObject>;
  findings: UnsafePdfAction[];
  inspected: number;
  limitReached: boolean;
};

function resolvedObject(document: PDFDocument, value: PDFObject | undefined) {
  if (!value) return undefined;
  return value instanceof PDFRef ? document.context.lookup(value) : value;
}

function actionName(value: PDFObject | undefined) {
  return value instanceof PDFName ? value.toString().replace(/^\//, "") : "Unknown";
}

function isNamedDestination(value: PDFObject) {
  return value instanceof PDFName || value instanceof PDFString;
}

function collectPageIdentities(document: PDFDocument) {
  const identities = new Set<PDFObject>();
  for (const page of document.getPages()) {
    identities.add(page.ref);
    identities.add(page.node);
  }
  return identities;
}

function isDocumentPage(context: WalkContext, value: PDFObject | undefined) {
  if (!value) return false;
  if (context.pageIdentities.has(value)) return true;
  const resolved = resolvedObject(context.document, value);
  return resolved !== undefined && context.pageIdentities.has(resolved);
}

function isNullOperand(value: PDFObject | undefined) {
  return value !== undefined && value.toString() === "null" && !(value instanceof PDFName) && !(value instanceof PDFString);
}

function isDestinationOperand(document: PDFDocument, value: PDFObject | undefined, allowNull: boolean) {
  const resolved = resolvedObject(document, value);
  if (resolved instanceof PDFNumber) return true;
  return allowNull && isNullOperand(resolved);
}

function isDestinationArray(context: WalkContext, value: PDFArray) {
  if (value.size() < 2) return false;
  if (!isDocumentPage(context, value.get(0))) return false;
  const fit = resolvedObject(context.document, value.get(1));
  if (!(fit instanceof PDFName)) return false;
  const spec = DESTINATION_OPERANDS[actionName(fit)];
  if (!spec || value.size() !== 2 + spec.count) return false;
  for (let index = 0; index < spec.count; index += 1) {
    if (!isDestinationOperand(context.document, value.get(2 + index), spec.allowNull)) return false;
  }
  return true;
}

function isStructuralDictionary(document: PDFDocument, value: PDFDict) {
  const type = resolvedObject(document, value.get(TYPE_NAME));
  if (!(type instanceof PDFName)) return false;
  const name = actionName(type);
  return name !== "Action" && STRUCTURAL_TYPES.has(name);
}

function recordFinding(
  context: WalkContext,
  page: number | undefined,
  location: UnsafePdfActionLocation,
  action: string,
) {
  const finding: UnsafePdfAction = page === undefined ? { location, action } : { page, location, action };
  context.findings.push(finding);
}

function markInspectionLimit(context: WalkContext) {
  context.limitReached = true;
  if (!context.findings.some((finding) => finding.location === "document" && finding.action === "InspectionLimit")) {
    recordFinding(context, undefined, "document", "InspectionLimit");
  }
}

function takeObject(context: WalkContext, resolved: PDFObject, visited: VisitMap, mode: WalkMode) {
  if (context.limitReached) return false;
  let modes = visited.get(resolved);
  if (!modes) {
    modes = new Set();
    visited.set(resolved, modes);
  }
  if (modes.has(mode)) return false;
  modes.add(mode);
  context.inspected += 1;
  if (context.inspected <= PDF_ACTION_INSPECTION_LIMIT) return true;
  markInspectionLimit(context);
  return false;
}

function enqueue(context: WalkContext, work: WorkItem[], item: WorkItem) {
  if (context.limitReached) return;
  if (work.length >= PDF_ACTION_INSPECTION_LIMIT) {
    markInspectionLimit(context);
    return;
  }
  work.push(item);
}

function enqueueAction(
  context: WalkContext,
  work: WorkItem[],
  value: PDFObject | undefined,
  page: number | undefined,
  location: UnsafePdfActionLocation,
  visited: VisitMap,
  mode: WalkMode = "action",
) {
  enqueue(context, work, { value, page, location, mode, visited });
}

/**
 * Inspect actions through an explicit work list. Recursion is avoided so a
 * deep Kids/Next chain hits the object bound instead of the JavaScript stack.
 * The only permitted action is internal GoTo. Destination arrays are recognised
 * only in action-or-destination context, and only when element 1 is a page from
 * this document's page tree, the length matches the fit operator, and every
 * remaining operand is a number or null as that operator allows. Cycle checks
 * are per mode so a structural visit cannot hide a later action inspection.
 * Generic action arrays are always walked as actions. `/AA` dictionaries are
 * event maps and are walked even when they carry an unrelated `/Type`. Known
 * structural dictionaries are left to dedicated walk modes.
 */
function drain(context: WalkContext, seed: WorkItem[]) {
  const work = seed;
  while (work.length && !context.limitReached) {
    const item = work.pop();
    if (!item) break;
    processWorkItem(context, work, item);
  }
}

function processWorkItem(context: WalkContext, work: WorkItem[], item: WorkItem) {
  if (context.limitReached) return;
  const resolved = resolvedObject(context.document, item.value);
  if (!resolved) return;

  if (item.mode === "actionOrDest") {
    if (isNamedDestination(resolved)) return;
    if (resolved instanceof PDFArray && isDestinationArray(context, resolved)) return;
    processWorkItem(context, work, { ...item, mode: "action" });
    return;
  }

  if (item.mode === "aaMap") {
    if (!(resolved instanceof PDFDict) || !takeObject(context, resolved, item.visited, item.mode)) return;
    const subtype = resolvedObject(context.document, resolved.get(ACTION_TYPE));
    if (subtype) {
      const name = actionName(subtype);
      if (name !== "GoTo") recordFinding(context, item.page, item.location, name);
    }
    for (const [, child] of resolved.entries()) {
      if (context.limitReached) return;
      enqueueAction(context, work, child, item.page, item.location, item.visited);
    }
    return;
  }

  if (item.mode === "jsEntry") {
    if (resolved instanceof PDFDict) {
      if (resolvedObject(context.document, resolved.get(ACTION_TYPE))) {
        processWorkItem(context, work, { ...item, mode: "action", location: "javascript" });
        return;
      }
      if (resolved.get(JS)) {
        if (!takeObject(context, resolved, item.visited, item.mode)) return;
        recordFinding(context, undefined, "javascript", "JavaScript");
        return;
      }
    }
    processWorkItem(context, work, { ...item, mode: "action", location: "javascript" });
    return;
  }

  if (item.mode === "nameTree") {
    if (!(resolved instanceof PDFDict) || !takeObject(context, resolved, item.visited, item.mode)) return;
    const names = resolvedObject(context.document, resolved.get(NAMES));
    if (names instanceof PDFArray) {
      for (let index = 1; index < names.size() && !context.limitReached; index += 2) {
        enqueueAction(context, work, names.get(index), undefined, "javascript", item.visited, "jsEntry");
      }
    }
    const kids = resolvedObject(context.document, resolved.get(KIDS));
    if (!(kids instanceof PDFArray)) return;
    for (let index = 0; index < kids.size() && !context.limitReached; index += 1) {
      enqueueAction(context, work, kids.get(index), undefined, "javascript", item.visited, "nameTree");
    }
    return;
  }

  if (item.mode === "outline") {
    if (!(resolved instanceof PDFDict) || !takeObject(context, resolved, item.visited, item.mode)) return;
    enqueueAction(context, work, resolved.get(ACTION), undefined, "outline", item.visited, "actionOrDest");
    enqueueAction(context, work, resolved.get(ADDITIONAL_ACTIONS), undefined, "outline", item.visited, "aaMap");
    enqueueAction(context, work, resolved.get(FIRST), undefined, "outline", item.visited, "outline");
    enqueueAction(context, work, resolved.get(NEXT), undefined, "outline", item.visited, "outline");
    return;
  }

  if (item.mode === "field") {
    if (!(resolved instanceof PDFDict) || !takeObject(context, resolved, item.visited, item.mode)) return;
    enqueueAction(context, work, resolved.get(ACTION), undefined, "form", item.visited);
    enqueueAction(context, work, resolved.get(ADDITIONAL_ACTIONS), undefined, "form", item.visited, "aaMap");
    const kids = resolvedObject(context.document, resolved.get(KIDS));
    if (!(kids instanceof PDFArray)) return;
    for (let index = 0; index < kids.size() && !context.limitReached; index += 1) {
      enqueueAction(context, work, kids.get(index), undefined, "form", item.visited, "field");
    }
    return;
  }

  if (!takeObject(context, resolved, item.visited, item.mode)) return;

  if (resolved instanceof PDFArray) {
    for (let index = 0; index < resolved.size() && !context.limitReached; index += 1) {
      enqueueAction(context, work, resolved.get(index), item.page, item.location, item.visited);
    }
    return;
  }
  if (!(resolved instanceof PDFDict)) return;

  const subtype = resolvedObject(context.document, resolved.get(ACTION_TYPE));
  if (subtype) {
    const name = actionName(subtype);
    if (name !== "GoTo") recordFinding(context, item.page, item.location, name);
    enqueueAction(context, work, resolved.get(NEXT_ACTION), item.page, item.location, item.visited);
    return;
  }

  if (isStructuralDictionary(context.document, resolved)) return;

  for (const [, child] of resolved.entries()) {
    if (context.limitReached) return;
    enqueueAction(context, work, child, item.page, item.location, item.visited);
  }
}

function inspectCatalog(context: WalkContext) {
  const catalog = context.document.catalog;
  drain(context, [{ value: catalog.get(OPEN_ACTION), location: "catalog", mode: "actionOrDest", visited: new Map() }]);
  drain(context, [{ value: catalog.get(ADDITIONAL_ACTIONS), location: "catalog", mode: "aaMap", visited: new Map() }]);

  const names = resolvedObject(context.document, catalog.get(NAMES));
  if (names instanceof PDFDict) {
    drain(context, [{ value: names.get(JAVASCRIPT), location: "javascript", mode: "nameTree", visited: new Map() }]);
  }

  const outlines = resolvedObject(context.document, catalog.get(OUTLINES));
  if (outlines instanceof PDFDict) {
    drain(context, [{ value: outlines.get(FIRST), location: "outline", mode: "outline", visited: new Map() }]);
  }

  const acroForm = resolvedObject(context.document, catalog.get(ACRO_FORM));
  if (!(acroForm instanceof PDFDict)) return;
  const fields = resolvedObject(context.document, acroForm.get(FIELDS));
  if (!(fields instanceof PDFArray)) return;
  const visited: VisitMap = new Map();
  const seed: WorkItem[] = [];
  for (let index = 0; index < fields.size() && !context.limitReached; index += 1) {
    enqueueAction(context, seed, fields.get(index), undefined, "form", visited, "field");
  }
  drain(context, seed);
}

function inspectPages(context: WalkContext) {
  context.document.getPages().forEach((page, pageIndex) => {
    const pageNumber = pageIndex + 1;
    drain(context, [{ value: page.node.get(ADDITIONAL_ACTIONS), page: pageNumber, location: "page", mode: "aaMap", visited: new Map() }]);
    const annotations = resolvedObject(context.document, page.node.get(ANNOTS));
    if (!(annotations instanceof PDFArray)) return;
    for (let index = 0; index < annotations.size(); index += 1) {
      const annotation = resolvedObject(context.document, annotations.get(index));
      if (!(annotation instanceof PDFDict)) continue;
      drain(context, [{ value: annotation.get(ACTION), page: pageNumber, location: "annotation", mode: "action", visited: new Map() }]);
      drain(context, [{ value: annotation.get(ADDITIONAL_ACTIONS), page: pageNumber, location: "annotation", mode: "aaMap", visited: new Map() }]);
    }
  });
}

export function formatUnsafePdfAction(finding: UnsafePdfAction) {
  const action = `/${finding.action}`;
  if (typeof finding.page === "number" && finding.page > 0) {
    return `page ${finding.page} ${finding.location} ${action}`;
  }
  return `${finding.location} ${action}`;
}

export function unsafePdfActions(document: PDFDocument): UnsafePdfAction[] {
  const context: WalkContext = {
    document,
    pageIdentities: collectPageIdentities(document),
    findings: [],
    inspected: 0,
    limitReached: false,
  };
  inspectCatalog(context);
  inspectPages(context);
  return context.findings;
}

export function assertPdfActionsSafe(document: PDFDocument, fileName: string) {
  const findings = unsafePdfActions(document);
  if (!findings.length) return;
  const details = findings.slice(0, 8).map(formatUnsafePdfAction).join(", ");
  const remainder = findings.length > 8 ? ` and ${findings.length - 8} more` : "";
  throw new Error(`${fileName} contains active PDF actions (${details}${remainder}). Flatten or remove the actions in a trusted PDF application, then replace this source file.`);
}
