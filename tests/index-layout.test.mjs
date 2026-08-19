import assert from "node:assert/strict";
import test from "node:test";

const measure = (text, size) => text.length * size * 0.5;

test("creates a deterministic immutable plan with wrapped text and bounded links", async () => {
  const { createIndexLayoutPlan } = await import("../app/lib/index-layout.ts");
  const rows = [
    { kind: "section", id: "agreements", title: "Agreements and amendments" },
    {
      kind: "exhibit",
      id: "e1",
      exhibitLabel: "AH 1",
      description: "Executed supply agreement with a deliberately long description that wraps into more than one line",
      pageLabel: "LV-0137GH-LV-0138GH",
      linkTargetId: "page-e1",
    },
  ];
  const first = createIndexLayoutPlan({ rows, measureText: measure });
  const second = createIndexLayoutPlan({ rows, measureText: measure });
  assert.equal(first.ok, true);
  assert.deepEqual(first, second);
  assert.equal(Object.isFrozen(first), true);
  assert.equal(Object.isFrozen(first.plan), true);
  assert.equal(Object.isFrozen(first.plan.rows[1]), true);
  assert.equal(first.plan.coordinateSystem, "top-left");
  assert.equal(first.plan.pages[0].rowIds.join(","), "agreements,e1");
  const exhibit = first.plan.rows[1];
  assert.equal(exhibit.kind, "exhibit");
  assert.ok(exhibit.descriptionLines.length > 1);
  assert.ok(exhibit.pageReferenceLines.length > 1);
  assert.ok(exhibit.pageReferenceFontSize >= first.plan.typography.minimumPageReferenceFontSize);
  assert.deepEqual(exhibit.linkRectangle, exhibit.bounds);
  assert.equal(exhibit.linkTargetId, "page-e1");
  for (const line of exhibit.pageReferenceLines) {
    assert.ok(line.x >= exhibit.cells.pageReference.x);
    assert.ok(line.x + line.width <= exhibit.cells.pageReference.x + exhibit.cells.pageReference.width);
  }
  assert.deepEqual(rows[1].description, "Executed supply agreement with a deliberately long description that wraps into more than one line");
});

test("moves a section and its first exhibit together to prevent an orphan heading", async () => {
  const { createIndexLayoutPlan } = await import("../app/lib/index-layout.ts");
  const geometry = {
    id: "small-test",
    pageWidth: 220,
    pageHeight: 160,
    contentTop: 20,
    contentBottom: 120,
    tableX: 10,
    tableWidth: 200,
    rowGap: 2,
    columns: {
      exhibit: { x: 10, width: 40 },
      description: { x: 50, width: 110 },
      pageReference: { x: 160, width: 50 },
    },
  };
  const result = createIndexLayoutPlan({
    geometry,
    measureText: measure,
    typography: { minimumExhibitRowHeight: 46, minimumSectionRowHeight: 30 },
    rows: [
      { kind: "exhibit", id: "e1", exhibitLabel: "1", description: "First", pageLabel: "1" },
      { kind: "section", id: "reports", title: "Reports" },
      { kind: "exhibit", id: "e2", exhibitLabel: "2", description: "Second", pageLabel: "2" },
    ],
  });
  assert.equal(result.ok, true);
  assert.equal(result.plan.pageCount, 2);
  assert.deepEqual(result.plan.pages.map((page) => page.rowIds), [["e1"], ["reports", "e2"]]);
  assert.deepEqual(result.plan.rows.map((row) => row.pageNumber), [1, 2, 2]);
});

test("uses the custom-template geometry profile and omits links for non-local exhibits", async () => {
  const { createIndexLayoutPlan, CUSTOM_TEMPLATE_INDEX_GEOMETRY } = await import("../app/lib/index-layout.ts");
  const result = createIndexLayoutPlan({
    geometry: "custom-template",
    measureText: measure,
    rows: [{ kind: "exhibit", id: "remote", exhibitLabel: "22", description: "Report in another volume", pageLabel: "Volume 3 / 413-470" }],
  });
  assert.equal(result.ok, true);
  assert.equal(result.plan.geometry.id, "custom-template");
  assert.deepEqual(result.plan.geometry, CUSTOM_TEMPLATE_INDEX_GEOMETRY);
  const row = result.plan.rows[0];
  assert.equal(row.kind, "exhibit");
  assert.equal(row.linkRectangle, null);
  assert.equal(row.linkTargetId, null);
  assert.equal(row.cells.pageReference.x, CUSTOM_TEMPLATE_INDEX_GEOMETRY.columns.pageReference.x);
  assert.equal(row.cells.date, undefined);
});

test("returns an explicit error when a page label cannot be rendered safely", async () => {
  const { createIndexLayoutPlan } = await import("../app/lib/index-layout.ts");
  const result = createIndexLayoutPlan({
    measureText: measure,
    typography: { minimumPageReferenceFontSize: 9, maximumPageReferenceLines: 2 },
    rows: [{ kind: "exhibit", id: "bad", exhibitLabel: "1", description: "Description", pageLabel: "X".repeat(200) }],
  });
  assert.deepEqual(result, {
    ok: false,
    error: {
      code: "UNRENDERABLE_PAGE_LABEL",
      message: "Page reference for row bad cannot fit within 2 lines at the minimum font size.",
      rowId: "bad",
      pageLabel: "X".repeat(200),
    },
  });
  assert.equal(Object.isFrozen(result), true);
  assert.equal(Object.isFrozen(result.error), true);
});

test("rejects overlapping custom columns and unavoidable orphan pairs", async () => {
  const { createIndexLayoutPlan } = await import("../app/lib/index-layout.ts");
  const invalidGeometry = {
    id: "overlap",
    pageWidth: 200,
    pageHeight: 200,
    contentTop: 10,
    contentBottom: 190,
    tableX: 10,
    tableWidth: 180,
    rowGap: 0,
    columns: {
      exhibit: { x: 10, width: 50 },
      description: { x: 50, width: 100 },
      pageReference: { x: 150, width: 40 },
    },
  };
  const invalid = createIndexLayoutPlan({ geometry: invalidGeometry, rows: [] });
  assert.equal(invalid.ok, false);
  assert.equal(invalid.error.code, "INVALID_GEOMETRY");

  const tinyGeometry = {
    ...invalidGeometry,
    id: "tiny",
    contentTop: 10,
    contentBottom: 80,
    columns: {
      exhibit: { x: 10, width: 40 },
      description: { x: 50, width: 100 },
      pageReference: { x: 150, width: 40 },
    },
  };
  const orphan = createIndexLayoutPlan({
    geometry: tinyGeometry,
    measureText: measure,
    rows: [
      { kind: "section", id: "s", title: "Section" },
      { kind: "exhibit", id: "e", exhibitLabel: "1", description: "Description", pageLabel: "1" },
    ],
  });
  assert.equal(orphan.ok, false);
  assert.equal(orphan.error.code, "ORPHAN_HEADING_UNAVOIDABLE");
  assert.equal(orphan.error.rowId, "s");
});

test("built-in geometry has a date cell and custom templates without Date do not", async () => {
  const { applyDetectedDateColumn, createIndexLayoutPlan, CUSTOM_TEMPLATE_INDEX_GEOMETRY, detectIndexTemplateDateColumn } = await import("../app/lib/index-layout.ts");
  const row = { kind: "exhibit", id: "e1", exhibitLabel: "1", description: "Supply agreement", pageLabel: "SS3", date: "14 August 2026" };
  const builtIn = createIndexLayoutPlan({ rows: [row], measureText: measure });
  assert.equal(builtIn.ok, true);
  const builtRow = builtIn.plan.rows[0];
  assert.equal(builtRow.kind, "exhibit");
  assert.ok(builtRow.cells.date, "built-in index includes a Date column");
  assert.equal(builtRow.dateLines[0].text, "14 August 2026");
  assert.ok(builtRow.dateLines[0].x >= builtRow.cells.date.x);
  assert.ok(builtRow.dateLines[0].x + builtRow.dateLines[0].width <= builtRow.cells.date.x + builtRow.cells.date.width);

  const threeColumn = detectIndexTemplateDateColumn([
    { str: "No.", x: 60, y: 700, width: 20 },
    { str: "Description", x: 99, y: 700, width: 60 },
    { str: "Page", x: 456, y: 700, width: 30 },
  ]);
  assert.equal(threeColumn, null);

  const withDate = detectIndexTemplateDateColumn([
    { str: "No.", x: 60, y: 700, width: 20 },
    { str: "Date", x: 99, y: 700, width: 28 },
    { str: "Description", x: 190, y: 700, width: 60 },
    { str: "Page", x: 456, y: 700, width: 30 },
  ]);
  assert.deepEqual(withDate, { x: 99, width: 91 });
  const datedGeometry = applyDetectedDateColumn(CUSTOM_TEMPLATE_INDEX_GEOMETRY, withDate);
  assert.ok(datedGeometry?.columns.date);
  const dated = createIndexLayoutPlan({
    geometry: datedGeometry,
    measureText: measure,
    rows: [row],
  });
  assert.equal(dated.ok, true);
  const datedRow = dated.plan.rows[0];
  assert.equal(datedRow.kind, "exhibit");
  assert.ok(datedRow.cells.date);
  assert.equal(datedRow.dateLines[0].text, "14 August 2026");
  assert.equal(datedGeometry.columns.pageReference.x, CUSTOM_TEMPLATE_INDEX_GEOMETRY.columns.pageReference.x);

  assert.equal(detectIndexTemplateDateColumn([{ str: "Date", x: 99, y: 700, width: 28 }]), null, "Date without companion headers is not a column");
  assert.equal(detectIndexTemplateDateColumn([{ str: "Date", x: 99, y: 100, width: 28 }, { str: "Page", x: 456, y: 100, width: 20 }]), null, "body-band Date is ignored");
  assert.equal(detectIndexTemplateDateColumn([
    { str: "Date", x: 200, y: 800, width: 28 },
    { str: "No.", x: 60, y: 700, width: 20 },
    { str: "Description", x: 99, y: 700, width: 60 },
    { str: "Page", x: 456, y: 700, width: 30 },
  ]), null, "a letterhead Date on a different row does not invent a fourth column");
  assert.ok(detectIndexTemplateDateColumn([
    { str: "No.", x: 60, y: 700, width: 20 },
    { str: "Date", x: 400, y: 700, width: 40 },
  ]).width >= 36);

  assert.equal(applyDetectedDateColumn(CUSTOM_TEMPLATE_INDEX_GEOMETRY, { x: 450, width: 10 }), null);
  assert.equal(applyDetectedDateColumn(CUSTOM_TEMPLATE_INDEX_GEOMETRY, { x: 99, width: 400 }), null);
  const shifted = applyDetectedDateColumn(CUSTOM_TEMPLATE_INDEX_GEOMETRY, { x: 50, width: 80 });
  assert.ok(shifted.columns.date.x >= CUSTOM_TEMPLATE_INDEX_GEOMETRY.columns.exhibit.x + CUSTOM_TEMPLATE_INDEX_GEOMETRY.columns.exhibit.width);
});

test("adds a group-break gap before the first unheaded exhibit after a named section", async () => {
  const { createIndexLayoutPlan, BUILT_IN_INDEX_GEOMETRY, CUSTOM_TEMPLATE_INDEX_GEOMETRY, DEFAULT_INDEX_TYPOGRAPHY } = await import("../app/lib/index-layout.ts");
  const exhibit = (id, extra = {}) => ({
    kind: "exhibit",
    id,
    exhibitLabel: id.toUpperCase(),
    description: `Description ${id}`,
    pageLabel: "1",
    ...extra,
  });
  const rows = [
    { kind: "section", id: "agreements", title: "Agreements" },
    exhibit("e1"),
    exhibit("e2"),
    exhibit("e3", { precedingGroupBreak: true }),
    exhibit("e4"),
  ];
  for (const geometry of [BUILT_IN_INDEX_GEOMETRY, CUSTOM_TEMPLATE_INDEX_GEOMETRY]) {
    const result = createIndexLayoutPlan({ geometry, measureText: measure, rows });
    assert.equal(result.ok, true, geometry.id);
    assert.equal(result.plan.rows.some((row) => row.kind === "section" && /no heading/i.test(row.title)), false);
    const [section, first, second, unheaded, trailing] = result.plan.rows;
    assert.equal(section.kind, "section");
    assert.equal(first.kind, "exhibit");
    assert.equal(second.bounds.top, first.bounds.top + first.bounds.height, `${geometry.id} intra-section spacing stays rowGap`);
    assert.equal(
      unheaded.bounds.top,
      second.bounds.top + second.bounds.height + DEFAULT_INDEX_TYPOGRAPHY.minimumSectionRowHeight,
      `${geometry.id} first unheaded exhibit sits one section-row below the headed run`,
    );
    assert.equal(trailing.bounds.top, unheaded.bounds.top + unheaded.bounds.height, `${geometry.id} later unheaded rows stay tight`);
  }
});

test("page-break arithmetic includes the unheaded group-break gap", async () => {
  const { createIndexLayoutPlan } = await import("../app/lib/index-layout.ts");
  const geometry = {
    id: "group-break-page",
    pageWidth: 220,
    pageHeight: 200,
    contentTop: 20,
    contentBottom: 142,
    tableX: 10,
    tableWidth: 200,
    rowGap: 0,
    columns: {
      exhibit: { x: 10, width: 40 },
      description: { x: 50, width: 110 },
      pageReference: { x: 160, width: 50 },
    },
  };
  const result = createIndexLayoutPlan({
    geometry,
    measureText: measure,
    typography: { minimumExhibitRowHeight: 46, minimumSectionRowHeight: 30 },
    rows: [
      { kind: "section", id: "agreements", title: "Agreements" },
      { kind: "exhibit", id: "e1", exhibitLabel: "1", description: "First", pageLabel: "1" },
      { kind: "exhibit", id: "e2", exhibitLabel: "2", description: "Unheaded", pageLabel: "2", precedingGroupBreak: true },
    ],
  });
  assert.equal(result.ok, true);
  assert.equal(result.plan.pageCount, 2);
  assert.deepEqual(result.plan.pages.map((page) => page.rowIds), [["agreements", "e1"], ["e2"]]);
});
