import assert from "node:assert/strict";
import JSZip from "jszip";
import test from "node:test";
import { renderWordTemplateHtml } from "../app/lib/template-converter.ts";

test("renders a DOCX template into self-contained A4 HTML", async () => {
  const zip = new JSZip();
  zip.file(
    "word/document.xml",
    [
      "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
      "<w:body>",
      "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val=\"28\"/></w:rPr><w:t>Cover &amp; index</w:t></w:r></w:p>",
      "<w:tbl><w:tblPr><w:tblBorders><w:insideH w:val=\"single\"/></w:tblBorders></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>Exhibit 1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Pages 1-4</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
      "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr>",
      "</w:body></w:document>",
    ].join(""),
  );
  const bytes = await zip.generateAsync({ type: "uint8array" });
  const html = await renderWordTemplateHtml(new File([bytes], "Cover_Template.docx"));

  assert.match(html, /@page \{ size: 210mm 297mm; margin: 0; \}/);
  assert.match(html, /Cover &amp; index/);
  assert.match(html, /font-weight:700/);
  assert.match(html, /Exhibit 1/);
  assert.match(html, /Pages 1-4/);
  assert.doesNotMatch(html, /https?:\/\//);
});

test("rejects a legacy binary DOC template with a clear offline message", async () => {
  const bytes = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]);
  const file = new File([bytes], "Legacy_Template.doc");
  await assert.rejects(
    import("../app/lib/template-converter.ts").then(({ convertWordTemplate }) => convertWordTemplate(file)),
    /Legacy \.doc templates cannot be rendered faithfully offline/,
  );
});

test("rejects an oversized DOCX archive before expansion", async () => {
  const oversized = new Uint8Array(25 * 1024 * 1024 + 1);
  oversized[0] = 0x50;
  oversized[1] = 0x4b;
  await assert.rejects(
    renderWordTemplateHtml(new File([oversized], "Oversized_Template.docx")),
    /25 MB offline safety limit/,
  );
});
