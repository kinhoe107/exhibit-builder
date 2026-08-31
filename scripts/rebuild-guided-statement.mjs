import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sampleDir = path.join(root, "public", "guided-sample");
const imageDir = path.join(sampleDir, "readme-images");
const outputPath = path.join(sampleDir, "Witness statement", "01_GUIDED_SAMPLE_Witness_Statement.docx");
const templatePath = outputPath;

const IMAGES = [
  {
    file: "home.png",
    caption: "Home screen. Open the guided sample folder, or click Run guided sample and follow the highlighted buttons.",
  },
  {
    file: "choose-files.png",
    caption: "Choose the witness statement and evidence files, then click Analyse files.",
  },
  {
    file: "review-confirm.png",
    caption: "Review. Confirm this document, or Confirm all proposed matches.",
  },
  {
    file: "unused-file.png",
    caption: "Unused files stay out of the bundle until you add them.",
  },
  {
    file: "finalise-build.png",
    caption: "Finalise. Click Build exhibit bundle when the matter is ready.",
  },
  {
    file: "save-download.png",
    caption: "Save exhibit project stays in the header. Download bundle PDF appears after the bundle is built.",
  },
];

const NUMBERED_PARAGRAPHS = [
  "The witness statement supplied to Exhibit Builder must be a DOCX file. Evidence exhibits can be PDF, DOCX, EML or XLSX files. This first example is a PDF: I refer to the SAMPLE AGREEMENT dated 1 August 2026 [Exhibit].",
  "The tool recognises general placeholders labelled Exhibit and placeholders with letters such as Exhib xx when they are placed in square brackets. I refer to the SAMPLE INVOICE PDF dated 2 August 2026 [Exhib xx].",
  "Several exhibits can be detected separately in the same paragraph, including different file types. I refer to the SAMPLE PROJECT REPORT in DOCX format dated 3 August 2026, the SAMPLE CLAIMANT EMAIL in EML format dated 4 August 2026 and the SAMPLE COST WORKBOOK in XLSX format dated 5 August 2026 [AH-xx; AH-xx; AH-xx].",
  "For an XLSX exhibit, sheet selection means choosing which worksheet tabs and detected cell ranges Microsoft Excel will print into the bundle. Tick the sheets that belong in the exhibit; unticked sheets are omitted. The preview shows the planned A4 pages, saved formula results are used without recalculation, and the source workbook is not edited.",
  "This guided project also supplies a PDF cover template and a PDF index template. The index template contains a complete three-column table with placeholder headings for the item number, exhibit description and bundle pages. The tool places the final index entries inside those columns.",
  "A repeated reference can use the already selected source instead of creating a duplicate exhibit. I refer again to the SAMPLE AGREEMENT dated 1 August 2026 [Exhibit]. Every match still requires human confirmation, and the source statement remains unchanged.",
  "The SAMPLE UNREFERENCED CHECKLIST is deliberately not cited. Use it to practise the warning shown when a user deliberately adds an exhibit that is not referred to in the witness statement.",
  "At the final-order stage, add an index heading and place exhibits beneath it. The heading appears in the index and becomes a parent PDF bookmark. Manual moves apply immediately; an automatic sort is only a preview until you choose Use this order, and existing headings are not silently discarded.",
  "If a bundle is split into volumes, every volume contains the complete index. The requested maximum includes the cover, repeated index and exhibit pages, and the default printed numbering continues across volumes unless the user deliberately chooses a different scheme.",
  "A supplied template may already contain a case or matter number, party names or other matter details. Exhibit Builder preserves the supplied PDF appearance but does not assume those details are correct: review the exact rendered PDF and confirm the detected matter details, visible placeholders and any disagreement between templates separately.",
  "The current automatic citation matching and local OCR are tested for English. Original PDF artwork and text remain in the source language, but translated placeholder words, non-English OCR and fully multilingual generated index text are not claimed in this release and require careful manual review.",
];

function escapeXml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function pngSize(bytes) {
  if (bytes[0] !== 0x89 || bytes[1] !== 0x50) throw new Error("A guided-sample screenshot must be a PNG file.");
  return {
    width: bytes.readUInt32BE(16),
    height: bytes.readUInt32BE(20),
  };
}

function textParagraph(text, extras = "") {
  return `<w:p>${extras}<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function titledParagraph(text, size, color, bold) {
  const boldXml = bold ? "<w:b/>" : "";
  return `<w:p><w:pPr><w:spacing w:after="80"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${size}"/><w:color w:val="${color}"/>${boldXml}</w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function numberedParagraph(text, numId) {
  return `<w:p><w:pPr><w:keepTogether/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function pictureParagraph(relId, name, widthPx, heightPx, docPrId) {
  const maxCx = 5486400;
  const cx = maxCx;
  const cy = Math.round((heightPx / widthPx) * maxCx);
  return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docPrId}" name="${escapeXml(name)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="${escapeXml(name)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

function decimalNumId(numberingXml) {
  const abstracts = [...numberingXml.matchAll(/<w:abstractNum w:abstractNumId="(\d+)"[\s\S]*?<\/w:abstractNum>/g)]
    .filter((match) => /<w:lvlText w:val="%1\."/.test(match[0]));
  const preferred = [...abstracts].reverse().find((match) => /w:hanging="270"/.test(match[0])) ?? abstracts.at(-1);
  if (!preferred) throw new Error("The statement template is missing a %1. numbering definition.");
  const instance = [...numberingXml.matchAll(/<w:num w:numId="(\d+)"[\s\S]*?<\/w:num>/g)]
    .find((match) => new RegExp(`<w:abstractNumId w:val="${preferred[1]}"`).test(match[0]));
  if (!instance) throw new Error("The statement template is missing a %1. numbering instance.");
  return instance[1];
}

function sectPrXml(documentXml) {
  const match = documentXml.match(/<w:sectPr[\s\S]*<\/w:sectPr>/);
  if (!match) throw new Error("The statement template is missing section properties.");
  return match[0];
}

async function main() {
  const template = await JSZip.loadAsync(await readFile(templatePath));
  const numberingXml = await template.file("word/numbering.xml").async("text");
  const originalDocument = await template.file("word/document.xml").async("text");
  const numId = decimalNumId(numberingXml);
  const imageParts = [];
  for (const [index, image] of IMAGES.entries()) {
    const bytes = await readFile(path.join(imageDir, image.file));
    const size = pngSize(bytes);
    const relId = `rIdImg${index + 1}`;
    const mediaName = `image${index + 1}.png`;
    imageParts.push({ ...image, bytes, size, relId, mediaName, docPrId: index + 1 });
  }

  const body = [
    titledParagraph("GUIDED TUTORIAL", "20", "2E74B5", true),
    titledParagraph("Guided Sample Witness Statement", "50", "0B2545", true),
    titledParagraph("This is an instructional guide, not a realistic witness statement.", "26", "5C6670", false),
    textParagraph("Open the guided sample folder and read this statement together with the sample exhibits. Then choose them with Choose witness statement and Choose evidence files, and click Analyse files. Run guided sample highlights those buttons through to Download bundle PDF and Save exhibit project."),
    ...imageParts.slice(0, 2).flatMap((image) => [
      pictureParagraph(image.relId, image.file, image.size.width, image.size.height, image.docPrId),
      textParagraph(image.caption),
    ]),
    titledParagraph("What the tool should recognise", "32", "2E74B5", true),
    ...NUMBERED_PARAGRAPHS.map((text) => numberedParagraph(text, numId)),
    ...imageParts.slice(2).flatMap((image) => [
      pictureParagraph(image.relId, image.file, image.size.width, image.size.height, image.docPrId),
      textParagraph(image.caption),
    ]),
    textParagraph("Review rule: Several placeholders in one paragraph are detected separately, but each proposed source, worksheet selection, template detail, grouping and final order remains subject to human review."),
    sectPrXml(originalDocument),
  ].join("");

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:body>${body}</w:body></w:document>`;

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>${imageParts.map((image) => `<Relationship Id="${image.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${image.mediaName}"/>`).join("")}</Relationships>`;

  const zip = await JSZip.loadAsync(await readFile(templatePath));
  zip.file("word/document.xml", documentXml);
  zip.file("word/_rels/document.xml.rels", rels);
  for (const image of imageParts) {
    zip.file(`word/media/${image.mediaName}`, image.bytes);
  }
  let contentTypes = await zip.file("[Content_Types].xml").async("text");
  if (!contentTypes.includes('Extension="png"')) {
    contentTypes = contentTypes.replace(
      "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>",
      "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/><Default Extension=\"png\" ContentType=\"image/png\"/>",
    );
  }
  zip.file("[Content_Types].xml", contentTypes);
  const core = await zip.file("docProps/core.xml").async("text");
  zip.file(
    "docProps/core.xml",
    core
      .replace(/<dc:creator>[^<]*<\/dc:creator>/, "<dc:creator></dc:creator>")
      .replace(/<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/, "<cp:lastModifiedBy></cp:lastModifiedBy>")
      .replace(/<dc:description>[^<]*<\/dc:description>/, "<dc:description>Instructional pictured README for the Exhibit Builder guided sample</dc:description>"),
  );

  const bytes = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await writeFile(outputPath, bytes);
  console.log(`Wrote ${outputPath}`);
}

await main();
