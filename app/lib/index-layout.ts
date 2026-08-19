/**
 * Pure, deterministic index layout.
 *
 * Coordinates use a top-left origin. A PDF renderer can convert a planned
 * baseline with `pdfY = pageHeight - baseline`. Keeping pagination and drawing
 * geometry in one immutable value prevents the planner and renderer from
 * making different wrapping or page-break decisions.
 */

export type IndexLayoutRectangle = Readonly<{
  x: number;
  top: number;
  width: number;
  height: number;
}>;

export type IndexLayoutColumn = Readonly<{
  x: number;
  width: number;
}>;

export type IndexLayoutGeometryProfile = Readonly<{
  id: string;
  pageWidth: number;
  pageHeight: number;
  contentTop: number;
  contentBottom: number;
  tableX: number;
  tableWidth: number;
  rowGap: number;
  columns: Readonly<{
    exhibit: IndexLayoutColumn;
    date?: IndexLayoutColumn;
    description: IndexLayoutColumn;
    pageReference: IndexLayoutColumn;
  }>;
}>;

export type IndexLayoutTypography = Readonly<{
  exhibitFontSize: number;
  descriptionFontSize: number;
  pageReferenceFontSize: number;
  minimumPageReferenceFontSize: number;
  sectionFontSize: number;
  lineHeight: number;
  sectionLineHeight: number;
  horizontalPadding: number;
  verticalPadding: number;
  minimumExhibitRowHeight: number;
  minimumSectionRowHeight: number;
  maximumPageReferenceLines: number;
}>;

export type IndexSectionInput = Readonly<{
  kind: "section";
  id: string;
  title: string;
}>;

export type IndexExhibitInput = Readonly<{
  kind: "exhibit";
  id: string;
  exhibitLabel: string;
  description: string;
  pageLabel: string;
  /** Printed only when the geometry profile already has a Date column. */
  date?: string;
  /** Omit for a row that must remain visible but cannot link within this PDF. */
  linkTargetId?: string;
  /** Extra space before the first unheaded exhibit after a named section. */
  precedingGroupBreak?: boolean;
}>;

export type IndexLayoutRowInput = IndexSectionInput | IndexExhibitInput;

export type IndexTextRole = "section" | "exhibit" | "description" | "date" | "page-reference";

export type IndexLayoutTextLine = Readonly<{
  text: string;
  x: number;
  top: number;
  baseline: number;
  width: number;
  fontSize: number;
  lineHeight: number;
  role: IndexTextRole;
}>;

type PlannedRowBase = Readonly<{
  id: string;
  pageNumber: number;
  bounds: IndexLayoutRectangle;
  height: number;
}>;

export type PlannedIndexSection = PlannedRowBase & Readonly<{
  kind: "section";
  title: string;
  lines: readonly IndexLayoutTextLine[];
}>;

export type PlannedIndexExhibit = PlannedRowBase & Readonly<{
  kind: "exhibit";
  exhibitLabel: string;
  description: string;
  pageLabel: string;
  pageReferenceFontSize: number;
  cells: Readonly<{
    exhibit: IndexLayoutRectangle;
    date?: IndexLayoutRectangle;
    description: IndexLayoutRectangle;
    pageReference: IndexLayoutRectangle;
  }>;
  exhibitLines: readonly IndexLayoutTextLine[];
  dateLines?: readonly IndexLayoutTextLine[];
  descriptionLines: readonly IndexLayoutTextLine[];
  pageReferenceLines: readonly IndexLayoutTextLine[];
  linkTargetId: string | null;
  linkRectangle: IndexLayoutRectangle | null;
}>;

export type PlannedIndexRow = PlannedIndexSection | PlannedIndexExhibit;

export type PlannedIndexPage = Readonly<{
  pageNumber: number;
  rowIds: readonly string[];
  usedHeight: number;
  remainingHeight: number;
}>;

export type IndexLayoutPlan = Readonly<{
  schemaVersion: "1.0";
  coordinateSystem: "top-left";
  geometry: IndexLayoutGeometryProfile;
  typography: IndexLayoutTypography;
  pageCount: number;
  pages: readonly PlannedIndexPage[];
  rows: readonly PlannedIndexRow[];
}>;

export type IndexLayoutErrorCode =
  | "INVALID_GEOMETRY"
  | "INVALID_TYPOGRAPHY"
  | "DUPLICATE_ROW_ID"
  | "EMPTY_ROW_TEXT"
  | "UNRENDERABLE_PAGE_LABEL"
  | "ROW_TOO_TALL"
  | "ORPHAN_HEADING_UNAVOIDABLE";

export type IndexLayoutFailure = Readonly<{
  code: IndexLayoutErrorCode;
  message: string;
  rowId?: string;
  pageLabel?: string;
}>;

export type IndexLayoutResult =
  | Readonly<{ ok: true; plan: IndexLayoutPlan }>
  | Readonly<{ ok: false; error: IndexLayoutFailure }>;

export type IndexTextMeasurer = (text: string, fontSize: number, role: IndexTextRole) => number;

export type CreateIndexLayoutOptions = Readonly<{
  rows: readonly IndexLayoutRowInput[];
  geometry?: "built-in" | "custom-template" | IndexLayoutGeometryProfile;
  typography?: Partial<IndexLayoutTypography>;
  measureText?: IndexTextMeasurer;
}>;

const A4_WIDTH = 595.28;
const A4_HEIGHT = 841.89;

export const BUILT_IN_INDEX_GEOMETRY: IndexLayoutGeometryProfile = deepFreeze({
  id: "built-in",
  pageWidth: A4_WIDTH,
  pageHeight: A4_HEIGHT,
  contentTop: 204,
  contentBottom: 739,
  tableX: 44,
  tableWidth: 507,
  rowGap: 0,
  columns: {
    exhibit: { x: 44, width: 40 },
    date: { x: 84, width: 110 },
    description: { x: 194, width: 262 },
    pageReference: { x: 456, width: 95 },
  },
});

export const CUSTOM_TEMPLATE_INDEX_GEOMETRY: IndexLayoutGeometryProfile = deepFreeze({
  id: "custom-template",
  pageWidth: A4_WIDTH,
  pageHeight: A4_HEIGHT,
  contentTop: 164,
  contentBottom: 699,
  tableX: 60,
  tableWidth: 483,
  rowGap: 0,
  columns: {
    exhibit: { x: 60, width: 39 },
    description: { x: 99, width: 357 },
    pageReference: { x: 456, width: 87 },
  },
});

export const DEFAULT_INDEX_TYPOGRAPHY: IndexLayoutTypography = deepFreeze({
  exhibitFontSize: 10,
  descriptionFontSize: 10,
  pageReferenceFontSize: 10,
  minimumPageReferenceFontSize: 6,
  sectionFontSize: 10,
  lineHeight: 12,
  sectionLineHeight: 13,
  horizontalPadding: 4,
  verticalPadding: 6,
  minimumExhibitRowHeight: 46,
  minimumSectionRowHeight: 30,
  maximumPageReferenceLines: 3,
});

type MeasuredLine = { text: string; width: number };

type PreparedSection = {
  kind: "section";
  input: IndexSectionInput;
  height: number;
  lines: MeasuredLine[];
};

type PreparedExhibit = {
  kind: "exhibit";
  input: IndexExhibitInput;
  height: number;
  exhibitLines: MeasuredLine[];
  dateLines: MeasuredLine[];
  descriptionLines: MeasuredLine[];
  pageReferenceLines: MeasuredLine[];
  pageReferenceFontSize: number;
};

export type IndexTemplateTextItem = Readonly<{
  str: string;
  x: number;
  y: number;
  width: number;
}>;

const DATE_HEADER = /^(date|dated|document\s*date)$/i;
const COMPANION_HEADER = /^(no\.?|#|exhibit|description|page|pages|pg\.?)$/i;

function normalizeHeader(text: string) {
  return text.replace(/\s+/gu, " ").trim();
}

/**
 * Detect a Date column from a custom index template's visible header text.
 * Three-column templates (No. / Description / Page) return null so a fourth
 * column is never invented.
 */
export function detectIndexTemplateDateColumn(
  items: readonly IndexTemplateTextItem[],
  pageHeight = A4_HEIGHT,
): IndexLayoutColumn | null {
  const band = pageHeight * 0.55;
  const rowTolerance = 8;
  const headers = items
    .filter((item) => item.y >= band)
    .map((item) => ({ ...item, label: normalizeHeader(item.str) }))
    .filter((item) => item.label);
  const companions = headers.filter((item) => COMPANION_HEADER.test(item.label));
  if (!companions.length) return null;
  const dated = headers.filter((item) => (
    DATE_HEADER.test(item.label)
    && companions.some((companion) => Math.abs(companion.y - item.y) <= rowTolerance)
  ));
  if (!dated.length) return null;
  const date = [...dated].sort((left, right) => left.x - right.x)[0];
  const sameRow = [...companions, ...dated].filter((item) => Math.abs(item.y - date.y) <= rowTolerance);
  const toTheRight = sameRow.filter((item) => item.x > date.x + 8).sort((left, right) => left.x - right.x)[0];
  const width = toTheRight ? Math.max(36, toTheRight.x - date.x) : Math.max(36, date.width + 16);
  return { x: round(date.x), width: round(Math.min(width, 160)) };
}

/** Map a detected Date header into custom-template geometry without moving the page column. */
export function applyDetectedDateColumn(
  base: IndexLayoutGeometryProfile,
  date: IndexLayoutColumn,
): IndexLayoutGeometryProfile | null {
  const exhibit = base.columns.exhibit;
  const description = base.columns.description;
  const pageReference = base.columns.pageReference;
  const tableRight = base.tableX + base.tableWidth;
  let x = Math.max(base.tableX, date.x);
  let width = Math.max(36, date.width);
  if (x + width > tableRight) width = tableRight - x;
  if (x < exhibit.x + exhibit.width) {
    const nextX = exhibit.x + exhibit.width;
    width -= nextX - x;
    x = nextX;
  }
  if (x + width > pageReference.x) width = pageReference.x - x;
  if (width < 36) return null;
  let descriptionX = description.x;
  let descriptionWidth = description.width;
  const dateRight = x + width;
  if (descriptionX < dateRight && descriptionX + descriptionWidth > x) {
    if (descriptionX < x) descriptionWidth = x - descriptionX;
    else {
      descriptionWidth = descriptionX + descriptionWidth - dateRight;
      descriptionX = dateRight;
    }
  }
  if (descriptionWidth < 80) return null;
  return deepFreeze({
    ...base,
    id: `${base.id}-with-date`,
    columns: {
      exhibit: { ...exhibit },
      date: { x: round(x), width: round(width) },
      description: { x: round(descriptionX), width: round(descriptionWidth) },
      pageReference: { ...pageReference },
    },
  });
}

type PreparedRow = PreparedSection | PreparedExhibit;

function deepFreeze<T>(value: T): T {
  if (!value || typeof value !== "object" || Object.isFrozen(value)) return value;
  Object.freeze(value);
  for (const nested of Object.values(value as Record<string, unknown>)) deepFreeze(nested);
  return value;
}

function round(value: number) {
  return Math.round(value * 1000) / 1000;
}

/** A stable fallback for planning before a renderer-specific font is available. */
export const defaultIndexTextMeasure: IndexTextMeasurer = (text, fontSize) => {
  let units = 0;
  for (const character of Array.from(text)) {
    if (/\s/u.test(character)) units += 0.28;
    else if (/[ilI1|.,:;'`]/u.test(character)) units += 0.3;
    else if (/[MW@%&#]/u.test(character)) units += 0.9;
    else if (/[A-Z]/u.test(character)) units += 0.65;
    else if (/[0-9]/u.test(character)) units += 0.56;
    else if (/[^\u0000-\u00ff]/u.test(character)) units += 1;
    else units += 0.52;
  }
  return round(units * fontSize);
};

function failure(code: IndexLayoutErrorCode, message: string, details: Pick<IndexLayoutFailure, "rowId" | "pageLabel"> = {}): IndexLayoutResult {
  return deepFreeze({ ok: false, error: { code, message, ...details } });
}

function isPositive(value: number) {
  return Number.isFinite(value) && value > 0;
}

function validateGeometry(geometry: IndexLayoutGeometryProfile): string | null {
  if (![geometry.pageWidth, geometry.pageHeight, geometry.tableWidth].every(isPositive)) return "Page and table dimensions must be positive.";
  if (!Number.isFinite(geometry.contentTop) || !Number.isFinite(geometry.contentBottom) || geometry.contentBottom <= geometry.contentTop) return "The content bounds must define a positive vertical area.";
  if (geometry.contentTop < 0 || geometry.contentBottom > geometry.pageHeight) return "The content bounds must stay inside the page.";
  if (geometry.tableX < 0 || geometry.tableX + geometry.tableWidth > geometry.pageWidth) return "The table must stay inside the page.";
  if (!Number.isFinite(geometry.rowGap) || geometry.rowGap < 0) return "The row gap cannot be negative.";
  const columns = [
    geometry.columns.exhibit,
    ...(geometry.columns.date ? [geometry.columns.date] : []),
    geometry.columns.description,
    geometry.columns.pageReference,
  ];
  if (columns.some((column) => !isPositive(column.width) || !Number.isFinite(column.x))) return "Every column must have a finite position and positive width.";
  const tableRight = geometry.tableX + geometry.tableWidth;
  if (columns.some((column) => column.x < geometry.tableX || column.x + column.width > tableRight)) return "Every column must stay inside the table.";
  const sorted = [...columns].sort((left, right) => left.x - right.x);
  if (sorted.some((column, index) => index > 0 && sorted[index - 1].x + sorted[index - 1].width > column.x)) return "Index columns cannot overlap.";
  return null;
}

function validateTypography(typography: IndexLayoutTypography): string | null {
  const positive = [
    typography.exhibitFontSize,
    typography.descriptionFontSize,
    typography.pageReferenceFontSize,
    typography.minimumPageReferenceFontSize,
    typography.sectionFontSize,
    typography.lineHeight,
    typography.sectionLineHeight,
    typography.minimumExhibitRowHeight,
    typography.minimumSectionRowHeight,
    typography.maximumPageReferenceLines,
  ];
  if (!positive.every(isPositive)) return "Font sizes, line heights, row heights and line limits must be positive.";
  if (typography.minimumPageReferenceFontSize > typography.pageReferenceFontSize) return "The minimum page-reference font cannot exceed its preferred font.";
  if (!Number.isFinite(typography.horizontalPadding) || typography.horizontalPadding < 0 || !Number.isFinite(typography.verticalPadding) || typography.verticalPadding < 0) return "Cell padding cannot be negative.";
  if (!Number.isInteger(typography.maximumPageReferenceLines)) return "The maximum page-reference line count must be an integer.";
  return null;
}

function hardBreakToken(token: string, maxWidth: number, fontSize: number, role: IndexTextRole, measure: IndexTextMeasurer): MeasuredLine[] {
  const result: MeasuredLine[] = [];
  let line = "";
  for (const character of Array.from(token)) {
    const candidate = `${line}${character}`;
    if (line && measure(candidate, fontSize, role) > maxWidth) {
      result.push({ text: line, width: round(measure(line, fontSize, role)) });
      line = character;
    } else line = candidate;
  }
  if (line) result.push({ text: line, width: round(measure(line, fontSize, role)) });
  return result;
}

function wrapText(text: string, maxWidth: number, fontSize: number, role: IndexTextRole, measure: IndexTextMeasurer): MeasuredLine[] {
  const words = text.trim().split(/\s+/u).filter(Boolean);
  if (!words.length) return [];
  const lines: MeasuredLine[] = [];
  let current = "";
  for (const word of words) {
    const pieces = measure(word, fontSize, role) <= maxWidth ? [{ text: word, width: measure(word, fontSize, role) }] : hardBreakToken(word, maxWidth, fontSize, role, measure);
    for (const piece of pieces) {
      const candidate = current ? `${current} ${piece.text}` : piece.text;
      if (current && measure(candidate, fontSize, role) > maxWidth) {
        lines.push({ text: current, width: round(measure(current, fontSize, role)) });
        current = piece.text;
      } else current = candidate;
    }
  }
  if (current) lines.push({ text: current, width: round(measure(current, fontSize, role)) });
  return lines;
}

function pageLabelFragments(label: string) {
  const fragments: string[] = [];
  let current = "";
  for (const character of Array.from(label.trim())) {
    current += character;
    if (/\s/u.test(character) || /[,;/\-\u2013\u2014]/u.test(character)) {
      if (current.trim()) fragments.push(current.trim());
      current = "";
    }
  }
  if (current.trim()) fragments.push(current.trim());
  return fragments;
}

function wrapPageLabel(label: string, maxWidth: number, fontSize: number, measure: IndexTextMeasurer): MeasuredLine[] | null {
  const fragments = pageLabelFragments(label);
  if (!fragments.length || fragments.some((fragment) => measure(fragment, fontSize, "page-reference") > maxWidth)) return null;
  const lines: MeasuredLine[] = [];
  let current = "";
  for (const fragment of fragments) {
    const separator = current && !current.endsWith("-") && !current.endsWith("/") && !current.endsWith("\u2013") && !current.endsWith("\u2014") ? " " : "";
    const candidate = `${current}${separator}${fragment}`;
    if (current && measure(candidate, fontSize, "page-reference") > maxWidth) {
      lines.push({ text: current, width: round(measure(current, fontSize, "page-reference")) });
      current = fragment;
    } else current = candidate;
  }
  if (current) lines.push({ text: current, width: round(measure(current, fontSize, "page-reference")) });
  return lines;
}

function prepareRows(
  rows: readonly IndexLayoutRowInput[],
  geometry: IndexLayoutGeometryProfile,
  typography: IndexLayoutTypography,
  measure: IndexTextMeasurer,
): IndexLayoutResult | PreparedRow[] {
  const available = {
    exhibit: geometry.columns.exhibit.width - typography.horizontalPadding * 2,
    date: geometry.columns.date ? geometry.columns.date.width - typography.horizontalPadding * 2 : 0,
    description: geometry.columns.description.width - typography.horizontalPadding * 2,
    pageReference: geometry.columns.pageReference.width - typography.horizontalPadding * 2,
    section: geometry.tableWidth - typography.horizontalPadding * 2,
  };
  const requiredWidths = geometry.columns.date
    ? Object.values(available)
    : [available.exhibit, available.description, available.pageReference, available.section];
  if (requiredWidths.some((width) => width <= 0)) return failure("INVALID_GEOMETRY", "Column padding leaves no usable text width.");
  const prepared: PreparedRow[] = [];
  for (const row of rows) {
    if (!row.id.trim()) return failure("EMPTY_ROW_TEXT", "Every index row needs an identifier.", { rowId: row.id });
    if (row.kind === "section") {
      if (!row.title.trim()) return failure("EMPTY_ROW_TEXT", "A section heading cannot be empty.", { rowId: row.id });
      const lines = wrapText(row.title, available.section, typography.sectionFontSize, "section", measure);
      const height = Math.max(typography.minimumSectionRowHeight, lines.length * typography.sectionLineHeight + typography.verticalPadding * 2);
      prepared.push({ kind: "section", input: { ...row }, lines, height: round(height) });
      continue;
    }
    if (![row.exhibitLabel, row.description, row.pageLabel].every((value) => value.trim())) return failure("EMPTY_ROW_TEXT", "Exhibit labels, descriptions and page references cannot be empty.", { rowId: row.id });
    const exhibitLines = wrapText(row.exhibitLabel, available.exhibit, typography.exhibitFontSize, "exhibit", measure);
    const dateText = geometry.columns.date ? (row.date?.trim() || "Date not stated") : "";
    const dateLines = geometry.columns.date
      ? wrapText(dateText, available.date, typography.descriptionFontSize, "date", measure)
      : [];
    const descriptionLines = wrapText(row.description, available.description, typography.descriptionFontSize, "description", measure);
    let fitted: { fontSize: number; lines: MeasuredLine[] } | null = null;
    const candidateSizes: number[] = [];
    for (let size = typography.pageReferenceFontSize; size > typography.minimumPageReferenceFontSize; size -= 0.25) candidateSizes.push(round(size));
    if (candidateSizes.at(-1) !== typography.minimumPageReferenceFontSize) candidateSizes.push(typography.minimumPageReferenceFontSize);
    for (const fontSize of candidateSizes) {
      const lines = wrapPageLabel(row.pageLabel, available.pageReference, fontSize, measure);
      if (lines && lines.length <= typography.maximumPageReferenceLines) {
        fitted = { fontSize, lines };
        break;
      }
    }
    if (!fitted) return failure(
      "UNRENDERABLE_PAGE_LABEL",
      `Page reference for row ${row.id} cannot fit within ${typography.maximumPageReferenceLines} lines at the minimum font size.`,
      { rowId: row.id, pageLabel: row.pageLabel },
    );
    const textHeight = Math.max(
      exhibitLines.length * typography.lineHeight,
      dateLines.length * typography.lineHeight,
      descriptionLines.length * typography.lineHeight,
      fitted.lines.length * typography.lineHeight,
    );
    const height = Math.max(typography.minimumExhibitRowHeight, textHeight + typography.verticalPadding * 2);
    prepared.push({
      kind: "exhibit",
      input: { ...row },
      exhibitLines,
      dateLines,
      descriptionLines,
      pageReferenceLines: fitted.lines,
      pageReferenceFontSize: fitted.fontSize,
      height: round(height),
    });
  }
  return prepared;
}

function textLines(
  lines: MeasuredLine[],
  column: IndexLayoutColumn,
  rowTop: number,
  fontSize: number,
  lineHeight: number,
  padding: Pick<IndexLayoutTypography, "horizontalPadding" | "verticalPadding">,
  role: IndexTextRole,
  align: "left" | "right" = "left",
): IndexLayoutTextLine[] {
  return lines.map((line, index) => {
    const x = align === "right"
      ? column.x + column.width - padding.horizontalPadding - line.width
      : column.x + padding.horizontalPadding;
    const top = rowTop + padding.verticalPadding + index * lineHeight;
    return {
      text: line.text,
      x: round(x),
      top: round(top),
      baseline: round(top + fontSize),
      width: line.width,
      fontSize,
      lineHeight,
      role,
    };
  });
}

function rectangle(x: number, top: number, width: number, height: number): IndexLayoutRectangle {
  return { x: round(x), top: round(top), width: round(width), height: round(height) };
}

function groupBreakGap(row: PreparedRow, used: number, typography: IndexLayoutTypography): number {
  if (used <= 0 || row.kind !== "exhibit" || !row.input.precedingGroupBreak) return 0;
  return typography.minimumSectionRowHeight;
}

/**
 * Creates an immutable index plan, or a typed failure when safe layout is not
 * possible. The function never mutates its row, geometry or typography input.
 */
export function createIndexLayoutPlan(options: CreateIndexLayoutOptions): IndexLayoutResult {
  const selected = options.geometry === "custom-template"
    ? CUSTOM_TEMPLATE_INDEX_GEOMETRY
    : !options.geometry || options.geometry === "built-in"
      ? BUILT_IN_INDEX_GEOMETRY
      : options.geometry;
  const geometry: IndexLayoutGeometryProfile = {
    ...selected,
    columns: {
      exhibit: { ...selected.columns.exhibit },
      ...(selected.columns.date ? { date: { ...selected.columns.date } } : {}),
      description: { ...selected.columns.description },
      pageReference: { ...selected.columns.pageReference },
    },
  };
  const typography: IndexLayoutTypography = { ...DEFAULT_INDEX_TYPOGRAPHY, ...options.typography };
  const geometryProblem = validateGeometry(geometry);
  if (geometryProblem) return failure("INVALID_GEOMETRY", geometryProblem);
  const typographyProblem = validateTypography(typography);
  if (typographyProblem) return failure("INVALID_TYPOGRAPHY", typographyProblem);
  const identifiers = new Set<string>();
  for (const row of options.rows) {
    if (identifiers.has(row.id)) return failure("DUPLICATE_ROW_ID", `Index row identifier ${row.id} is duplicated.`, { rowId: row.id });
    identifiers.add(row.id);
  }
  const prepared = prepareRows(options.rows, geometry, typography, options.measureText ?? defaultIndexTextMeasure);
  if (!Array.isArray(prepared)) return prepared;

  const pageCapacity = geometry.contentBottom - geometry.contentTop;
  for (const row of prepared) {
    if (row.height > pageCapacity) return failure("ROW_TOO_TALL", `Index row ${row.input.id} is taller than the usable page area.`, { rowId: row.input.id });
  }

  const placements: Array<{ prepared: PreparedRow; pageNumber: number; top: number }> = [];
  const pageRows: string[][] = [[]];
  const pageUsed: number[] = [0];
  let pageNumber = 1;
  let used = 0;
  const newPage = () => {
    pageNumber += 1;
    used = 0;
    pageRows.push([]);
    pageUsed.push(0);
  };
  for (let index = 0; index < prepared.length; index += 1) {
    const row = prepared[index];
    const precedingGap = (used > 0 ? geometry.rowGap : 0) + groupBreakGap(row, used, typography);
    if (row.kind === "section" && prepared[index + 1]?.kind === "exhibit") {
      const companion = prepared[index + 1];
      const pairHeight = row.height + geometry.rowGap + groupBreakGap(companion, 1, typography) + companion.height;
      if (pairHeight > pageCapacity) return failure(
        "ORPHAN_HEADING_UNAVOIDABLE",
        `Section ${row.input.id} and its first exhibit cannot fit together in the usable page area.`,
        { rowId: row.input.id },
      );
      if (used > 0 && used + precedingGap + pairHeight > pageCapacity) newPage();
    } else if (used > 0 && used + precedingGap + row.height > pageCapacity) newPage();
    const gap = (used > 0 ? geometry.rowGap : 0) + groupBreakGap(row, used, typography);
    const top = geometry.contentTop + used + gap;
    placements.push({ prepared: row, pageNumber, top: round(top) });
    pageRows[pageNumber - 1].push(row.input.id);
    used += gap + row.height;
    pageUsed[pageNumber - 1] = round(used);
  }

  const rows: PlannedIndexRow[] = placements.map(({ prepared: row, pageNumber: assignedPage, top }) => {
    const bounds = rectangle(geometry.tableX, top, geometry.tableWidth, row.height);
    if (row.kind === "section") {
      return {
        kind: "section",
        id: row.input.id,
        title: row.input.title,
        pageNumber: assignedPage,
        bounds,
        height: row.height,
        lines: textLines(
          row.lines,
          { x: geometry.tableX, width: geometry.tableWidth },
          top,
          typography.sectionFontSize,
          typography.sectionLineHeight,
          typography,
          "section",
        ),
      };
    }
    const dateColumn = geometry.columns.date;
    const cells = {
      exhibit: rectangle(geometry.columns.exhibit.x, top, geometry.columns.exhibit.width, row.height),
      ...(dateColumn ? { date: rectangle(dateColumn.x, top, dateColumn.width, row.height) } : {}),
      description: rectangle(geometry.columns.description.x, top, geometry.columns.description.width, row.height),
      pageReference: rectangle(geometry.columns.pageReference.x, top, geometry.columns.pageReference.width, row.height),
    };
    return {
      kind: "exhibit",
      id: row.input.id,
      exhibitLabel: row.input.exhibitLabel,
      description: row.input.description,
      pageLabel: row.input.pageLabel,
      pageNumber: assignedPage,
      bounds,
      height: row.height,
      cells,
      exhibitLines: textLines(row.exhibitLines, geometry.columns.exhibit, top, typography.exhibitFontSize, typography.lineHeight, typography, "exhibit"),
      ...(dateColumn ? { dateLines: textLines(row.dateLines, dateColumn, top, typography.descriptionFontSize, typography.lineHeight, typography, "date") } : {}),
      descriptionLines: textLines(row.descriptionLines, geometry.columns.description, top, typography.descriptionFontSize, typography.lineHeight, typography, "description"),
      pageReferenceLines: textLines(row.pageReferenceLines, geometry.columns.pageReference, top, row.pageReferenceFontSize, typography.lineHeight, typography, "page-reference", "right"),
      pageReferenceFontSize: row.pageReferenceFontSize,
      linkTargetId: row.input.linkTargetId ?? null,
      linkRectangle: row.input.linkTargetId ? bounds : null,
    };
  });
  const pages: PlannedIndexPage[] = pageRows.map((rowIds, index) => ({
    pageNumber: index + 1,
    rowIds: [...rowIds],
    usedHeight: pageUsed[index],
    remainingHeight: round(pageCapacity - pageUsed[index]),
  }));
  return deepFreeze({
    ok: true,
    plan: {
      schemaVersion: "1.0",
      coordinateSystem: "top-left",
      geometry,
      typography,
      pageCount: pages.length,
      pages,
      rows,
    },
  });
}
