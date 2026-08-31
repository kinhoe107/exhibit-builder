import assert from "node:assert/strict";
import test from "node:test";
import { PDFDocument, PDFHexString, PDFName, PDFNull } from "pdf-lib";

const { assertPdfActionsSafe, formatUnsafePdfAction, PDF_ACTION_INSPECTION_LIMIT, unsafePdfActions } = await import("../app/lib/pdf-action-safety.ts");

async function onePagePdf() {
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([595.28, 841.89]);
  page.drawText("Document-level action fixture");
  return { pdf, page };
}

function launchAction(pdf) {
  return pdf.context.obj({ Type: "Action", S: "Launch", F: PDFHexString.fromText("calc.exe") });
}

test("formatUnsafePdfAction omits page text for document-level findings", () => {
  assert.equal(formatUnsafePdfAction({ page: 1, location: "page", action: "Launch" }), "page 1 page /Launch");
  assert.equal(formatUnsafePdfAction({ location: "catalog", action: "Launch" }), "catalog /Launch");
  assert.equal(formatUnsafePdfAction({ page: 0, location: "catalog", action: "Launch" }), "catalog /Launch");
});

test("catalog OpenAction Launch is unsafe and does not use page 0", async () => {
  const { pdf } = await onePagePdf();
  pdf.catalog.set(PDFName.of("OpenAction"), pdf.context.register(launchAction(pdf)));
  const findings = unsafePdfActions(pdf);
  assert.deepEqual(findings, [{ location: "catalog", action: "Launch" }]);
  assert.throws(
    () => assertPdfActionsSafe(pdf, "Catalog_Launch.pdf"),
    (error) => {
      assert.match(error.message, /catalog \/Launch/);
      assert.doesNotMatch(error.message, /page 0/);
      return true;
    },
  );
});

test("catalog OpenAction destination arrays, named destinations and GoTo stay safe", async () => {
  const { pdf, page } = await onePagePdf();
  pdf.catalog.set(PDFName.of("OpenAction"), pdf.context.obj([page.ref, PDFName.of("Fit")]));
  assert.deepEqual(unsafePdfActions(pdf), []);

  const named = await onePagePdf();
  named.pdf.catalog.set(PDFName.of("OpenAction"), PDFName.of("Somewhere"));
  assert.deepEqual(unsafePdfActions(named.pdf), []);

  const stringNamed = await onePagePdf();
  stringNamed.pdf.catalog.set(PDFName.of("OpenAction"), PDFHexString.fromText("Somewhere"));
  assert.deepEqual(unsafePdfActions(stringNamed.pdf), []);

  const goTo = await onePagePdf();
  const action = goTo.pdf.context.obj({ Type: "Action", S: "GoTo", D: [goTo.page.ref, PDFName.of("Fit")] });
  goTo.pdf.catalog.set(PDFName.of("OpenAction"), goTo.pdf.context.register(action));
  assert.deepEqual(unsafePdfActions(goTo.pdf), []);
  assert.doesNotThrow(() => assertPdfActionsSafe(goTo.pdf, "Safe_OpenAction.pdf"));
});

test("catalog AA action arrays are not treated as destinations", async () => {
  const { pdf, page } = await onePagePdf();
  pdf.catalog.set(PDFName.of("AA"), pdf.context.obj({
    WC: pdf.context.obj([page.ref, PDFName.of("Fit"), pdf.context.register(launchAction(pdf))]),
  }));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "catalog", action: "Launch" }]);
});

test("catalog OpenAction fake page dictionary is not a destination", async () => {
  const { pdf } = await onePagePdf();
  const fakePage = pdf.context.obj({ Type: "Page", S: "Launch" });
  pdf.catalog.set(PDFName.of("OpenAction"), pdf.context.obj([pdf.context.register(fakePage), PDFName.of("Fit")]));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "catalog", action: "Launch" }]);
  assert.throws(
    () => assertPdfActionsSafe(pdf, "Catalog_Fake_Page_Destination.pdf"),
    (error) => {
      assert.match(error.message, /catalog \/Launch/);
      return true;
    },
  );
});

test("malformed destination arrays with trailing actions are unsafe", async () => {
  const { pdf, page } = await onePagePdf();
  pdf.catalog.set(
    PDFName.of("OpenAction"),
    pdf.context.obj([page.ref, PDFName.of("Fit"), pdf.context.register(launchAction(pdf))]),
  );
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "catalog", action: "Launch" }]);
});

test("fit-specific destination shapes stay safe", async () => {
  const fitH = await onePagePdf();
  fitH.pdf.catalog.set(PDFName.of("OpenAction"), fitH.pdf.context.obj([fitH.page.ref, PDFName.of("FitH"), 0]));
  assert.deepEqual(unsafePdfActions(fitH.pdf), []);

  const xyz = await onePagePdf();
  xyz.pdf.catalog.set(PDFName.of("OpenAction"), xyz.pdf.context.obj([xyz.page.ref, PDFName.of("XYZ"), 0, 0, 0]));
  assert.deepEqual(unsafePdfActions(xyz.pdf), []);

  const fitR = await onePagePdf();
  fitR.pdf.catalog.set(
    PDFName.of("OpenAction"),
    fitR.pdf.context.obj([fitR.page.ref, PDFName.of("FitR"), 0, 0, 100, 100]),
  );
  assert.deepEqual(unsafePdfActions(fitR.pdf), []);

  const xyzNull = await onePagePdf();
  xyzNull.pdf.catalog.set(
    PDFName.of("OpenAction"),
    xyzNull.pdf.context.obj([xyzNull.page.ref, PDFName.of("XYZ"), PDFNull, PDFNull, PDFNull]),
  );
  assert.deepEqual(unsafePdfActions(xyzNull.pdf), []);

  const fitV = await onePagePdf();
  fitV.pdf.catalog.set(PDFName.of("OpenAction"), fitV.pdf.context.obj([fitV.page.ref, PDFName.of("FitV"), 0]));
  assert.deepEqual(unsafePdfActions(fitV.pdf), []);

  const fitBH = await onePagePdf();
  fitBH.pdf.catalog.set(PDFName.of("OpenAction"), fitBH.pdf.context.obj([fitBH.page.ref, PDFName.of("FitBH"), 0]));
  assert.deepEqual(unsafePdfActions(fitBH.pdf), []);

  const fitBV = await onePagePdf();
  fitBV.pdf.catalog.set(PDFName.of("OpenAction"), fitBV.pdf.context.obj([fitBV.page.ref, PDFName.of("FitBV"), 0]));
  assert.deepEqual(unsafePdfActions(fitBV.pdf), []);
});

test("catalog OpenAction FitH with a Launch operand is not a destination", async () => {
  const { pdf, page } = await onePagePdf();
  pdf.catalog.set(
    PDFName.of("OpenAction"),
    pdf.context.obj([page.ref, PDFName.of("FitH"), pdf.context.register(launchAction(pdf))]),
  );
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "catalog", action: "Launch" }]);
  assert.throws(
    () => assertPdfActionsSafe(pdf, "Catalog_FitH_Launch.pdf"),
    (error) => {
      assert.match(error.message, /catalog \/Launch/);
      return true;
    },
  );
});

test("outline item with self-referential Launch action is unsafe", async () => {
  const { pdf } = await onePagePdf();
  const item = pdf.context.obj({ Title: PDFHexString.fromText("Open calculator") });
  const itemRef = pdf.context.register(item);
  item.set(PDFName.of("S"), PDFName.of("Launch"));
  item.set(PDFName.of("A"), itemRef);
  item.set(PDFName.of("Parent"), itemRef);
  pdf.catalog.set(PDFName.of("Outlines"), pdf.context.obj({
    Type: "Outlines",
    First: itemRef,
    Last: itemRef,
    Count: 1,
  }));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "outline", action: "Launch" }]);
  assert.throws(() => assertPdfActionsSafe(pdf, "Outline_Self_Launch.pdf"), /outline \/Launch/);
});

test("AcroForm field with self-referential Launch action is unsafe", async () => {
  const { pdf } = await onePagePdf();
  const field = pdf.context.obj({ FT: "Btn", T: PDFHexString.fromText("Go") });
  const fieldRef = pdf.context.register(field);
  field.set(PDFName.of("S"), PDFName.of("Launch"));
  field.set(PDFName.of("A"), fieldRef);
  pdf.catalog.set(PDFName.of("AcroForm"), pdf.context.obj({ Fields: [fieldRef] }));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "form", action: "Launch" }]);
  assert.throws(() => assertPdfActionsSafe(pdf, "Field_Self_Launch.pdf"), /form \/Launch/);
});

test("catalog OpenAction array with Launch then Fit is not treated as a destination", async () => {
  const { pdf } = await onePagePdf();
  const launchRef = pdf.context.register(launchAction(pdf));
  pdf.catalog.set(PDFName.of("OpenAction"), pdf.context.obj([launchRef, PDFName.of("Fit")]));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "catalog", action: "Launch" }]);
  assert.throws(
    () => assertPdfActionsSafe(pdf, "Catalog_Launch_Lookalike.pdf"),
    (error) => {
      assert.match(error.message, /catalog \/Launch/);
      assert.doesNotMatch(error.message, /page 0/);
      return true;
    },
  );
});

test("typed catalog and page AA maps still expose Launch children", async () => {
  const { pdf } = await onePagePdf();
  pdf.catalog.set(PDFName.of("AA"), pdf.context.obj({
    Type: "NotAnAction",
    WC: pdf.context.register(launchAction(pdf)),
  }));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "catalog", action: "Launch" }]);

  const pageCase = await onePagePdf();
  pageCase.page.node.set(PDFName.of("AA"), pageCase.pdf.context.obj({
    Type: "NotAnAction",
    O: pageCase.pdf.context.register(launchAction(pageCase.pdf)),
  }));
  assert.deepEqual(unsafePdfActions(pageCase.pdf), [{ page: 1, location: "page", action: "Launch" }]);
});

test("catalog additional-actions Launch is unsafe", async () => {
  const { pdf } = await onePagePdf();
  pdf.catalog.set(PDFName.of("AA"), pdf.context.obj({ WC: pdf.context.register(launchAction(pdf)) }));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "catalog", action: "Launch" }]);
});

test("document JavaScript name trees including Kids are unsafe", async () => {
  const { pdf } = await onePagePdf();
  const jsAction = pdf.context.obj({ Type: "Action", S: "JavaScript", JS: PDFHexString.fromText("app.alert(1)") });
  const leaf = pdf.context.obj({ Names: [PDFHexString.fromText("evil"), pdf.context.register(jsAction)] });
  const tree = pdf.context.obj({ Kids: [pdf.context.register(leaf)] });
  pdf.catalog.set(PDFName.of("Names"), pdf.context.obj({ JavaScript: pdf.context.register(tree) }));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "javascript", action: "JavaScript" }]);

  const scriptDict = await onePagePdf();
  const jsDict = scriptDict.pdf.context.obj({ JS: PDFHexString.fromText("app.alert(1)") });
  const names = scriptDict.pdf.context.obj({ Names: [PDFHexString.fromText("boot"), scriptDict.pdf.context.register(jsDict)] });
  scriptDict.pdf.catalog.set(PDFName.of("Names"), scriptDict.pdf.context.obj({ JavaScript: scriptDict.pdf.context.register(names) }));
  assert.deepEqual(unsafePdfActions(scriptDict.pdf), [{ location: "javascript", action: "JavaScript" }]);
});

test("outline item Launch actions are unsafe", async () => {
  const { pdf } = await onePagePdf();
  const item = pdf.context.obj({
    Title: PDFHexString.fromText("Open calculator"),
    A: pdf.context.register(launchAction(pdf)),
  });
  const itemRef = pdf.context.register(item);
  item.set(PDFName.of("Parent"), itemRef);
  item.set(PDFName.of("First"), itemRef);
  item.set(PDFName.of("Next"), itemRef);
  pdf.catalog.set(PDFName.of("Outlines"), pdf.context.obj({
    Type: "Outlines",
    First: itemRef,
    Last: itemRef,
    Count: 1,
  }));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "outline", action: "Launch" }]);
});

test("AcroForm field additional-actions and Kids are unsafe", async () => {
  const { pdf } = await onePagePdf();
  const child = pdf.context.obj({
    FT: "Btn",
    T: PDFHexString.fromText("Child"),
    AA: { D: pdf.context.register(launchAction(pdf)) },
  });
  const parent = pdf.context.obj({
    T: PDFHexString.fromText("Parent"),
    Kids: [pdf.context.register(child)],
  });
  pdf.catalog.set(PDFName.of("AcroForm"), pdf.context.obj({ Fields: [pdf.context.register(parent)] }));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "form", action: "Launch" }]);
});

test("non-GoTo catalog actions remain refused", async () => {
  for (const action of ["URI", "SubmitForm", "GoToR", "Named"]) {
    const { pdf } = await onePagePdf();
    pdf.catalog.set(PDFName.of("OpenAction"), pdf.context.register(pdf.context.obj({ Type: "Action", S: action })));
    assert.deepEqual(unsafePdfActions(pdf), [{ location: "catalog", action }]);
  }
});

test("cyclic additional-actions finish without hanging and still report Launch", async () => {
  const { pdf } = await onePagePdf();
  const aa = pdf.context.obj({});
  const aaRef = pdf.context.register(aa);
  aa.set(PDFName.of("O"), aaRef);
  pdf.catalog.set(PDFName.of("AA"), aaRef);
  pdf.catalog.set(PDFName.of("OpenAction"), pdf.context.register(launchAction(pdf)));
  assert.deepEqual(unsafePdfActions(pdf), [{ location: "catalog", action: "Launch" }]);
});

test("inspection stops fail-closed when the object bound is exceeded", async () => {
  const { pdf } = await onePagePdf();
  let node = pdf.context.obj({ Names: [] });
  for (let index = 0; index < PDF_ACTION_INSPECTION_LIMIT + 2; index += 1) {
    node = pdf.context.obj({ Kids: [pdf.context.register(node)] });
  }
  pdf.catalog.set(PDFName.of("Names"), pdf.context.obj({ JavaScript: pdf.context.register(node) }));
  const findings = unsafePdfActions(pdf);
  assert.ok(findings.some((finding) => finding.location === "document" && finding.action === "InspectionLimit"));
  assert.throws(
    () => assertPdfActionsSafe(pdf, "Bounded.pdf"),
    (error) => {
      assert.match(error.message, /document \/InspectionLimit/);
      assert.doesNotMatch(error.message, /page 0/);
      return true;
    },
  );
});
