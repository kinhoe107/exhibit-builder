export type BuildPlanReference = {
  id: string;
  relativeStart: number | null;
  relativeEnd: number | null;
};

export type BuildPlanItemInput = {
  id: string;
  recordIndex: number;
  indexNumber: number;
  witnessKey: string;
  initials: string;
  sequence: number;
  sourceHashes: string[];
  bodyStartPage: number;
  bodyContentStartPage: number;
  bodyEndPage: number;
  physicalPages: number;
  optionalPages: number;
  contentPages: number;
  references: BuildPlanReference[];
};

export type PlannedBuildItem = BuildPlanItemInput & {
  volumeNumber: number;
  physicalStartPage: number;
  physicalContentStartPage: number;
  physicalEndPage: number;
  legalStartPage: number;
  legalEndPage: number;
  references: Array<BuildPlanReference & {
    legalStartPage: number | null;
    legalEndPage: number | null;
  }>;
};

export type PlannedBuildVolume = {
  number: number;
  label: string;
  fileName: string;
  coverPages: number;
  indexPages: number;
  exhibitPages: number;
  referencePages: number;
  totalPages: number;
  oversize: boolean;
  items: PlannedBuildItem[];
};

export type BuildPlanIndexNode =
  | {
      kind: "section";
      id: string;
      title: string;
      itemIds: string[];
    }
  | {
      kind: "exhibit";
      itemId: string;
    };

export type BundleBuildPlan = {
  schemaVersion: "1.0";
  bundleIdentity: string;
  canonicalOrder: string[];
  pageLimit: number;
  includeDividerPages: boolean;
  includeExhibitCoverPages: boolean;
  countOptionalPagesInReferences: boolean;
  /** Complete reader-facing index structure repeated in every physical volume. */
  indexNodes: BuildPlanIndexNode[];
  multiVolume: boolean;
  volumes: PlannedBuildVolume[];
};

export type BuildPlanOptions = {
  pageLimit: number;
  coverPages: number;
  includeDividerPages: boolean;
  includeExhibitCoverPages: boolean;
  countOptionalPagesInReferences: boolean;
  matchPdfPageOrder: boolean;
  volumeNumbering: "continuous" | "restart";
  startAt: number;
  /** Exact page count returned by the authoritative IndexLayoutPlan. */
  completeIndexPages: number;
  /** Canonical section/exhibit structure. Every item must occur exactly once. */
  indexNodes?: BuildPlanIndexNode[];
};

function physicalPageCount(items: BuildPlanItemInput[]) {
  return items.reduce((total, item) => total + item.physicalPages, 0);
}

function totalPageCount(items: BuildPlanItemInput[], coverPages: number, completeIndexPages: number) {
  return coverPages + completeIndexPages + physicalPageCount(items);
}

function deepFreeze<T>(value: T): T {
  if (!value || typeof value !== "object" || Object.isFrozen(value)) return value;
  Object.freeze(value);
  for (const nested of Object.values(value as Record<string, unknown>)) deepFreeze(nested);
  return value;
}

/**
 * Creates the single authoritative assembly plan used by the PDF compositor,
 * manifest, validation, downloads and statement-reference suggestions.
 * Administrative volumes never alter the witness exhibit-bundle identity.
 */
export function createBuildPlan(
  sourceItems: BuildPlanItemInput[],
  options: BuildPlanOptions,
): BundleBuildPlan {
  if (!sourceItems.length) throw new Error("A build plan needs at least one exhibit.");
  // Finished-PDF page order is the only numbering source. Optional divider and
  // cover pages always advance that sequence. A saved exhibit-local scheme
  // cannot restart stamps, index rows or statement suggestions.
  const matchPdfPageOrder = true;
  const countOptionalPagesInReferences = true;
  const items = sourceItems.map((item) => ({
    ...item,
    sourceHashes: [...item.sourceHashes],
    references: item.references.map((reference) => ({ ...reference })),
  }));
  const coverPages = Math.max(1, Math.floor(options.coverPages));
  const completeIndexPages = Math.floor(options.completeIndexPages);
  if (!Number.isSafeInteger(completeIndexPages) || completeIndexPages < 1) {
    throw new Error("The complete index must contain at least one page.");
  }
  const indexNodes: BuildPlanIndexNode[] = options.indexNodes
    ? options.indexNodes.map((node) => node.kind === "section"
      ? { kind: "section", id: node.id, title: node.title, itemIds: [...node.itemIds] }
      : { kind: "exhibit", itemId: node.itemId })
    : items.map((item) => ({ kind: "exhibit", itemId: item.id }));
  const itemIds = items.map((item) => item.id);
  if (new Set(itemIds).size !== itemIds.length) throw new Error("Build-plan exhibit IDs must be unique.");
  const indexItemIds = indexNodes.flatMap((node) => node.kind === "section" ? node.itemIds : [node.itemId]);
  if (
    indexNodes.some((node) => node.kind === "section" && (!node.id.trim() || !node.title.trim())) ||
    indexItemIds.length !== itemIds.length ||
    new Set(indexItemIds).size !== indexItemIds.length ||
    indexItemIds.some((id, index) => id !== itemIds[index])
  ) {
    throw new Error("The complete index structure must contain every exhibit exactly once in canonical order.");
  }
  const pageLimit = Number.isFinite(options.pageLimit) ? Math.max(0, Math.floor(options.pageLimit)) : 0;
  const packed: BuildPlanItemInput[][] = [];
  let current: BuildPlanItemInput[] = [];
  const flush = () => {
    if (current.length) packed.push(current);
    current = [];
  };
  for (const item of items) {
    const next = [...current, item];
    if (pageLimit > 0 && current.length && totalPageCount(next, coverPages, completeIndexPages) > pageLimit) flush();
    current.push(item);
  }
  flush();

  const identities = new Set(items.map((item) => `${item.initials.replace(/\s+/g, "").toUpperCase()} ${item.sequence}`));
  const bundleIdentity = identities.size === 1 ? [...identities][0] : "Exhibit Bundle";
  const multiVolume = packed.length > 1;
  const packedPageCounts = packed.map((volumeItems) => ({
    // Every physical PDF carries the same complete bundle index. The page
    // target therefore includes this full repeated preliminary section.
    indexPages: completeIndexPages,
    exhibitPages: physicalPageCount(volumeItems),
    referencePages: volumeItems.reduce((total, item) => total + item.contentPages + (countOptionalPagesInReferences ? item.optionalPages : 0), 0),
  }));
  const volumes = packed.map((volumeItems, volumeIndex): PlannedBuildVolume => {
    const number = volumeIndex + 1;
    const { indexPages, exhibitPages, referencePages } = packedPageCounts[volumeIndex];
    const totalPages = coverPages + indexPages + exhibitPages;
    const earlier = packedPageCounts.slice(0, volumeIndex);
    const continuousPdfOffset = options.volumeNumbering === "continuous"
      ? earlier.reduce((total, counts) => total + coverPages + counts.indexPages + counts.exhibitPages, 0)
      : 0;
    const continuousReferenceOffset = options.volumeNumbering === "continuous"
      ? earlier.reduce((total, counts) => total + counts.referencePages, 0)
      : 0;
    let referenceCursor = Math.max(1, Math.floor(options.startAt)) + continuousReferenceOffset;
    let physicalCursor = coverPages + indexPages + 1;
    const plannedItems = volumeItems.map((item): PlannedBuildItem => {
      const countedOptionalPages = countOptionalPagesInReferences ? item.optionalPages : 0;
      const physicalStartPage = physicalCursor;
      const physicalContentStartPage = physicalStartPage + item.optionalPages;
      const physicalEndPage = physicalStartPage + item.physicalPages - 1;
      const legalStartPage = matchPdfPageOrder
        ? physicalContentStartPage + continuousPdfOffset
        : referenceCursor + countedOptionalPages;
      const legalEndPage = legalStartPage + item.contentPages - 1;
      if (!matchPdfPageOrder) referenceCursor = legalEndPage + 1;
      physicalCursor = physicalEndPage + 1;
      return {
        ...item,
        volumeNumber: number,
        physicalStartPage,
        physicalContentStartPage,
        physicalEndPage,
        legalStartPage,
        legalEndPage,
        references: item.references.map((reference) => ({
          ...reference,
          legalStartPage: reference.relativeStart === null ? null : legalStartPage + reference.relativeStart,
          legalEndPage: reference.relativeEnd === null ? null : legalStartPage + reference.relativeEnd,
        })),
      };
    });
    const label = multiVolume ? `${bundleIdentity} — Volume ${number}` : bundleIdentity;
    const safeIdentity = bundleIdentity.replace(/\s+/g, "");
    return {
      number,
      label,
      fileName: multiVolume ? `Exhibit_Bundle_${safeIdentity}_Volume_${number}.pdf` : "Exhibit_Bundle.pdf",
      coverPages,
      indexPages,
      exhibitPages,
      referencePages,
      totalPages,
      oversize: pageLimit > 0 && totalPages > pageLimit,
      items: plannedItems,
    };
  });

  return deepFreeze({
    schemaVersion: "1.0",
    bundleIdentity,
    canonicalOrder: items.map((item) => item.id),
    pageLimit,
    includeDividerPages: options.includeDividerPages,
    includeExhibitCoverPages: options.includeExhibitCoverPages,
    countOptionalPagesInReferences,
    indexNodes,
    multiVolume,
    volumes,
  });
}
