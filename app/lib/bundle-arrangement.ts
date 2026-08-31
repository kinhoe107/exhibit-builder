/**
 * The persisted, authoritative arrangement of exhibits and index sections.
 *
 * A top-level exhibit is intentionally different from an exhibit inside a
 * section: its position among section nodes is meaningful and is preserved.
 * All operations in this module are pure and return a new arrangement.
 */
export type ArrangementExhibitNode = {
  type: "exhibit";
  exhibitId: string;
};

export type ArrangementSectionNode = {
  type: "section";
  id: string;
  heading: string;
  exhibits: ArrangementExhibitNode[];
};

export type ArrangementNode = ArrangementExhibitNode | ArrangementSectionNode;

export type BundleArrangement = {
  version: 1;
  nodes: ArrangementNode[];
};

export type ExhibitMoveTarget =
  | { sectionId: string; index: number }
  | { sectionId: null; index: number };

export type NewArrangementSection = {
  id: string;
  heading: string;
  /** Top-level insertion index after selected exhibits have been removed. */
  index: number;
  exhibitIds?: readonly string[];
};

export const ARRANGEMENT_LIMITS = {
  nodes: 20_000,
  idLength: 1_024,
  headingLength: 512,
} as const;

function cloneExhibit(node: ArrangementExhibitNode): ArrangementExhibitNode {
  return { type: "exhibit", exhibitId: node.exhibitId };
}

export function cloneBundleArrangement(arrangement: BundleArrangement): BundleArrangement {
  return {
    version: 1,
    nodes: arrangement.nodes.map((node) => node.type === "exhibit"
      ? cloneExhibit(node)
      : { type: "section", id: node.id, heading: node.heading, exhibits: node.exhibits.map(cloneExhibit) }),
  };
}

function validIdentifier(value: unknown): value is string {
  return typeof value === "string" && value.length > 0 && value.length <= ARRANGEMENT_LIMITS.idLength && !value.includes("\0");
}

function validHeading(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0 && value.length <= ARRANGEMENT_LIMITS.headingLength && !value.includes("\0");
}

/** Return all structural validation issues without mutating the supplied value. */
export function bundleArrangementIssues(value: unknown): string[] {
  if (!value || typeof value !== "object" || Array.isArray(value)) return ["Arrangement must be an object."];
  const candidate = value as Partial<BundleArrangement>;
  if (candidate.version !== 1) return ["Arrangement version is unsupported."];
  if (!Array.isArray(candidate.nodes)) return ["Arrangement nodes must be an array."];
  const issues: string[] = [];
  const exhibitIds = new Set<string>();
  const sectionIds = new Set<string>();
  let nodeCount = candidate.nodes.length;
  if (candidate.nodes.length > ARRANGEMENT_LIMITS.nodes) issues.push("Arrangement contains too many top-level nodes.");

  const recordExhibit = (node: unknown, location: string) => {
    if (!node || typeof node !== "object" || Array.isArray(node) || (node as { type?: unknown }).type !== "exhibit" || !validIdentifier((node as { exhibitId?: unknown }).exhibitId)) {
      issues.push(`${location} is not a valid exhibit node.`);
      return;
    }
    const exhibitId = (node as ArrangementExhibitNode).exhibitId;
    if (exhibitIds.has(exhibitId)) issues.push(`Exhibit ${exhibitId} occurs more than once in the arrangement.`);
    exhibitIds.add(exhibitId);
  };

  candidate.nodes.forEach((node, index) => {
    if (!node || typeof node !== "object" || Array.isArray(node)) {
      issues.push(`Arrangement node ${index + 1} is invalid.`);
      return;
    }
    if ((node as { type?: unknown }).type === "exhibit") {
      recordExhibit(node, `Arrangement node ${index + 1}`);
      return;
    }
    if ((node as { type?: unknown }).type !== "section") {
      issues.push(`Arrangement node ${index + 1} has an unsupported type.`);
      return;
    }
    const section = node as Partial<ArrangementSectionNode>;
    if (!validIdentifier(section.id)) issues.push(`Arrangement section ${index + 1} has an invalid ID.`);
    else if (sectionIds.has(section.id)) issues.push(`Section ${section.id} occurs more than once in the arrangement.`);
    else sectionIds.add(section.id);
    if (!validHeading(section.heading)) issues.push(`Arrangement section ${index + 1} has an invalid heading.`);
    if (!Array.isArray(section.exhibits)) {
      issues.push(`Arrangement section ${index + 1} must contain an exhibit array.`);
      return;
    }
    nodeCount += section.exhibits.length;
    section.exhibits.forEach((exhibit, exhibitIndex) => recordExhibit(exhibit, `Exhibit ${exhibitIndex + 1} in section ${section.id ?? index + 1}`));
  });
  if (nodeCount > ARRANGEMENT_LIMITS.nodes) issues.push("Arrangement contains too many total nodes.");
  return issues;
}

export function isBundleArrangement(value: unknown): value is BundleArrangement {
  return bundleArrangementIssues(value).length === 0;
}

export function validateBundleArrangement(value: unknown): BundleArrangement {
  const issues = bundleArrangementIssues(value);
  if (issues.length) throw new Error(`Bundle arrangement is invalid: ${issues[0]}`);
  return cloneBundleArrangement(value as BundleArrangement);
}

/** Safely migrate the old flat finalOrder value into the new model. */
export function bundleArrangementFromLegacyOrder(finalOrder: unknown): BundleArrangement {
  if (finalOrder === undefined || finalOrder === null) return { version: 1, nodes: [] };
  if (!Array.isArray(finalOrder) || finalOrder.length > ARRANGEMENT_LIMITS.nodes || !finalOrder.every((item): item is string => typeof item === "string")) throw new Error("Legacy final order is invalid or too large.");
  const nodes = finalOrder.map((exhibitId) => ({ type: "exhibit" as const, exhibitId }));
  return validateBundleArrangement({ version: 1, nodes });
}

export function flattenBundleArrangement(arrangement: BundleArrangement): string[] {
  const valid = validateBundleArrangement(arrangement);
  return valid.nodes.flatMap((node) => node.type === "exhibit" ? [node.exhibitId] : node.exhibits.map((exhibit) => exhibit.exhibitId));
}

export type ExhibitContainerLocation = {
  sectionId: string | null;
  /** Index within the section or the contiguous unheaded run. */
  index: number;
  length: number;
  /** Top-level index of the unheaded run start; unused for headed exhibits. */
  topLevelStart: number;
};

/**
 * Locate an exhibit for Earlier/Later/Top/Bottom. Unheaded exhibits are
 * confined to their contiguous run so a heading block is never one “step”.
 */
export function exhibitContainerLocation(arrangement: BundleArrangement, exhibitId: string): ExhibitContainerLocation | null {
  const valid = validateBundleArrangement(arrangement);
  for (let nodeIndex = 0; nodeIndex < valid.nodes.length; nodeIndex += 1) {
    const node = valid.nodes[nodeIndex];
    if (node.type === "exhibit" && node.exhibitId === exhibitId) {
      let start = nodeIndex;
      let end = nodeIndex;
      while (start > 0 && valid.nodes[start - 1].type === "exhibit") start -= 1;
      while (end < valid.nodes.length - 1 && valid.nodes[end + 1].type === "exhibit") end += 1;
      return { sectionId: null, index: nodeIndex - start, length: end - start + 1, topLevelStart: start };
    }
    if (node.type === "section") {
      const index = node.exhibits.findIndex((exhibit) => exhibit.exhibitId === exhibitId);
      if (index >= 0) return { sectionId: node.id, index, length: node.exhibits.length, topLevelStart: 0 };
    }
  }
  return null;
}

/** Move an exhibit inside its section or unheaded run. Index is container-relative. */
export function moveArrangementExhibitInContainer(arrangement: BundleArrangement, exhibitId: string, targetIndex: number): BundleArrangement {
  const valid = validateBundleArrangement(arrangement);
  const location = exhibitContainerLocation(valid, exhibitId);
  if (!location) throw new Error(`Exhibit ${exhibitId} is not present in the arrangement.`);
  if (!Number.isSafeInteger(targetIndex) || targetIndex < 0 || targetIndex >= location.length) {
    throw new Error("Arrangement move index is outside the current container.");
  }
  if (location.sectionId) return moveArrangementExhibit(valid, exhibitId, { sectionId: location.sectionId, index: targetIndex });
  return moveArrangementExhibit(valid, exhibitId, { sectionId: null, index: location.topLevelStart + targetIndex });
}

function siblingInsertIndex(
  exhibits: Array<{ exhibitId: string }>,
  start: number,
  isSibling: (id: string) => boolean,
) {
  let insertAt = start;
  while (insertAt < exhibits.length && isSibling(exhibits[insertAt].exhibitId)) insertAt += 1;
  return insertAt;
}

function lookupAfter(insertAfter: Readonly<Record<string, string>>, exhibitId: string) {
  const afterId = insertAfter[exhibitId];
  return typeof afterId === "string" && afterId.length > 0 ? afterId : undefined;
}

function topLevelSiblingRunEnd(nodes: ArrangementNode[], start: number, isSibling: (id: string) => boolean) {
  let insertAt = start;
  while (insertAt < nodes.length) {
    const following = nodes[insertAt];
    if (following.type !== "exhibit" || !isSibling(following.exhibitId)) break;
    insertAt += 1;
  }
  return insertAt;
}

function placeNewExhibitAfter(
  nodes: ArrangementNode[],
  exhibitId: string,
  afterId: string,
  insertAfter: Readonly<Record<string, string>>,
) {
  const isSibling = (id: string) => lookupAfter(insertAfter, id) === afterId;
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (node.type === "exhibit" && node.exhibitId === afterId) {
      nodes.splice(topLevelSiblingRunEnd(nodes, index + 1, isSibling), 0, { type: "exhibit", exhibitId });
      return true;
    }
    if (node.type !== "section") continue;
    const parentIndex = node.exhibits.findIndex((exhibit) => exhibit.exhibitId === afterId);
    if (parentIndex < 0) continue;
    const exhibits = [...node.exhibits];
    exhibits.splice(siblingInsertIndex(exhibits, parentIndex + 1, isSibling), 0, { type: "exhibit", exhibitId });
    nodes[index] = { ...node, exhibits };
    return true;
  }
  return false;
}

/**
 * Remove missing exhibits, preserve every surviving position and section, and
 * append newly discovered exhibits as unsectioned top-level nodes. Optional
 * insertAfter reseats those exhibits immediately after their parent (and after
 * any siblings already placed there), including exhibits that already had a slot.
 */
export function reconcileBundleArrangement(
  arrangement: BundleArrangement,
  exhibitIds: readonly string[],
  insertAfter: Readonly<Record<string, string>> = {},
): BundleArrangement {
  const valid = validateBundleArrangement(arrangement);
  if (exhibitIds.length > ARRANGEMENT_LIMITS.nodes || exhibitIds.some((id) => !validIdentifier(id)) || new Set(exhibitIds).size !== exhibitIds.length) {
    throw new Error("The current exhibit IDs are invalid, duplicated, or exceed the arrangement limit.");
  }
  const current = new Set(exhibitIds);
  const gluedIds = new Set(exhibitIds.filter((id) => Boolean(lookupAfter(insertAfter, id))));
  const retained = new Set<string>();
  const nodes: ArrangementNode[] = [];
  for (const node of valid.nodes) {
    if (node.type === "exhibit") {
      if (current.has(node.exhibitId) && !gluedIds.has(node.exhibitId)) {
        nodes.push(node);
        retained.add(node.exhibitId);
      }
      continue;
    }
    const exhibits = node.exhibits.filter((exhibit) => current.has(exhibit.exhibitId) && !gluedIds.has(exhibit.exhibitId));
    exhibits.forEach((exhibit) => retained.add(exhibit.exhibitId));
    // Empty user-created sections remain meaningful and are preserved.
    nodes.push({ ...node, exhibits });
  }
  for (const exhibitId of exhibitIds) {
    if (retained.has(exhibitId)) continue;
    const afterId = lookupAfter(insertAfter, exhibitId);
    if (afterId && placeNewExhibitAfter(nodes, exhibitId, afterId, insertAfter)) {
      retained.add(exhibitId);
      continue;
    }
    nodes.push({ type: "exhibit", exhibitId });
    retained.add(exhibitId);
  }
  return { version: 1, nodes };
}

function withoutExhibits(arrangement: BundleArrangement, exhibitIds: ReadonlySet<string>): ArrangementNode[] {
  return arrangement.nodes.flatMap<ArrangementNode>((node) => {
    if (node.type === "exhibit") return exhibitIds.has(node.exhibitId) ? [] : [cloneExhibit(node)];
    return [{ ...node, exhibits: node.exhibits.filter((exhibit) => !exhibitIds.has(exhibit.exhibitId)).map(cloneExhibit) }];
  });
}

export function moveArrangementExhibit(arrangement: BundleArrangement, exhibitId: string, target: ExhibitMoveTarget): BundleArrangement {
  const valid = validateBundleArrangement(arrangement);
  if (!flattenBundleArrangement(valid).includes(exhibitId)) throw new Error(`Exhibit ${exhibitId} is not present in the arrangement.`);
  if (!Number.isSafeInteger(target.index) || target.index < 0) throw new Error("Arrangement move index is invalid.");
  const nodes = withoutExhibits(valid, new Set([exhibitId]));
  const exhibit: ArrangementExhibitNode = { type: "exhibit", exhibitId };
  if (target.sectionId === null) {
    if (target.index > nodes.length) throw new Error("Arrangement move index is outside the top-level arrangement.");
    nodes.splice(target.index, 0, exhibit);
    return { version: 1, nodes };
  }
  const sectionIndex = nodes.findIndex((node) => node.type === "section" && node.id === target.sectionId);
  if (sectionIndex < 0) throw new Error(`Arrangement section ${target.sectionId} does not exist.`);
  const section = nodes[sectionIndex] as ArrangementSectionNode;
  if (target.index > section.exhibits.length) throw new Error("Arrangement move index is outside the target section.");
  const exhibits = [...section.exhibits];
  exhibits.splice(target.index, 0, exhibit);
  nodes[sectionIndex] = { ...section, exhibits };
  return { version: 1, nodes };
}

export function addArrangementSection(arrangement: BundleArrangement, section: NewArrangementSection): BundleArrangement {
  const valid = validateBundleArrangement(arrangement);
  if (!validIdentifier(section.id) || !validHeading(section.heading)) throw new Error("Arrangement section ID or heading is invalid.");
  if (!Number.isSafeInteger(section.index) || section.index < 0) throw new Error("Arrangement section index is invalid.");
  if (valid.nodes.some((node) => node.type === "section" && node.id === section.id)) throw new Error(`Arrangement section ${section.id} already exists.`);
  const selected = section.exhibitIds ?? [];
  if (new Set(selected).size !== selected.length) throw new Error("An exhibit cannot be selected for a section more than once.");
  const currentOrder = flattenBundleArrangement(valid);
  const present = new Set(currentOrder);
  if (selected.some((id) => !present.has(id))) throw new Error("A selected exhibit is not present in the arrangement.");
  const selectedSet = new Set(selected);
  const nodes = withoutExhibits(valid, selectedSet);
  if (section.index > nodes.length) throw new Error("Arrangement section index is outside the top-level arrangement.");
  nodes.splice(section.index, 0, {
    type: "section",
    id: section.id,
    heading: section.heading,
    // Creating a heading must not silently reorder the selected exhibits.
    exhibits: currentOrder.filter((exhibitId) => selectedSet.has(exhibitId)).map((exhibitId) => ({ type: "exhibit", exhibitId })),
  });
  return { version: 1, nodes };
}

export function renameArrangementSection(arrangement: BundleArrangement, sectionId: string, heading: string): BundleArrangement {
  const valid = validateBundleArrangement(arrangement);
  if (!validHeading(heading)) throw new Error("Arrangement section heading is invalid.");
  let found = false;
  const nodes = valid.nodes.map((node) => {
    if (node.type !== "section" || node.id !== sectionId) return node;
    found = true;
    return { ...node, heading };
  });
  if (!found) throw new Error(`Arrangement section ${sectionId} does not exist.`);
  return { version: 1, nodes };
}

/** Delete a section while keeping its exhibits at the section's former position. */
export function deleteArrangementSectionKeepItems(arrangement: BundleArrangement, sectionId: string): BundleArrangement {
  const valid = validateBundleArrangement(arrangement);
  const index = valid.nodes.findIndex((node) => node.type === "section" && node.id === sectionId);
  if (index < 0) throw new Error(`Arrangement section ${sectionId} does not exist.`);
  const section = valid.nodes[index] as ArrangementSectionNode;
  const nodes = [...valid.nodes];
  nodes.splice(index, 1, ...section.exhibits.map(cloneExhibit));
  return { version: 1, nodes };
}

/** Move a whole section without changing the exhibits it owns. */
export function moveArrangementSection(arrangement: BundleArrangement, sectionId: string, topLevelIndex: number): BundleArrangement {
  const valid = validateBundleArrangement(arrangement);
  if (!Number.isSafeInteger(topLevelIndex) || topLevelIndex < 0) throw new Error("Arrangement section move index is invalid.");
  const from = valid.nodes.findIndex((node) => node.type === "section" && node.id === sectionId);
  if (from < 0) throw new Error(`Arrangement section ${sectionId} does not exist.`);
  const nodes = [...valid.nodes];
  const [section] = nodes.splice(from, 1);
  if (topLevelIndex > nodes.length) throw new Error("Arrangement section move index is outside the top-level arrangement.");
  nodes.splice(topLevelIndex, 0, section);
  return { version: 1, nodes };
}

/** Insert a heading (and its exhibits) before the given top-level node index. */
export function moveArrangementSectionBefore(
  arrangement: BundleArrangement,
  sectionId: string,
  beforeNodeIndex: number,
): BundleArrangement {
  const valid = validateBundleArrangement(arrangement);
  if (!Number.isSafeInteger(beforeNodeIndex) || beforeNodeIndex < 0 || beforeNodeIndex > valid.nodes.length) {
    throw new Error("Arrangement section move index is invalid.");
  }
  const from = valid.nodes.findIndex((node) => node.type === "section" && node.id === sectionId);
  if (from < 0) throw new Error(`Arrangement section ${sectionId} does not exist.`);
  if (from === beforeNodeIndex || from + 1 === beforeNodeIndex) return valid;
  return moveArrangementSection(valid, sectionId, from < beforeNodeIndex ? beforeNodeIndex - 1 : beforeNodeIndex);
}

/**
 * Sort each section independently. Contiguous unsectioned runs are sorted
 * independently so exhibits never cross a section boundary as a side effect.
 */
export function sortBundleArrangementWithinSections(
  arrangement: BundleArrangement,
  compare: (leftExhibitId: string, rightExhibitId: string) => number,
): BundleArrangement {
  const valid = validateBundleArrangement(arrangement);
  const nodes: ArrangementNode[] = [];
  let run: ArrangementExhibitNode[] = [];
  const flush = () => {
    nodes.push(...run.sort((left, right) => compare(left.exhibitId, right.exhibitId)));
    run = [];
  };
  for (const node of valid.nodes) {
    if (node.type === "exhibit") {
      run.push(cloneExhibit(node));
      continue;
    }
    flush();
    nodes.push({ ...node, exhibits: [...node.exhibits].sort((left, right) => compare(left.exhibitId, right.exhibitId)) });
  }
  flush();
  return { version: 1, nodes };
}
