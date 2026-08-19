# Exhibit Builder

Exhibit Builder is an offline Windows desktop application that turns exhibit
references in a final DOCX witness statement into a reviewed, indexed and
paginated PDF exhibit bundle.

> **Licence:** This repository is publicly viewable proprietary software. It is
> source-available for transparency, security review and evaluation, but it is
> not open-source software. See [LICENSE](LICENSE). The packaged application is
> licensed separately under [build/EULA.txt](build/EULA.txt).

## What it does

- Finds exhibit placeholders and document references in one DOCX statement,
  including several exhibits in the same paragraph.
- Accepts PDF, DOCX, EML and XLSX evidence.
- Requires the reviewer to confirm proposed document matches and review the
  final exhibit order before building.
- Prints XLSX worksheets through locally installed Microsoft Excel to preserve
  their saved colours, fonts, borders, number formats and print layout.
- Builds an indexed PDF with bookmarks that use the index number, internal
  links, visible page numbers and copy-ready statement-reference suggestions.
- Supports optional cover, index, exhibit-cover and divider templates.
- Records source and output hashes, exclusions, manual additions, warnings and
  reviewer-approved technical exceptions in a JSON manifest.
- Can split a large bundle into separate physical volumes while keeping one
  continuous page-number sequence by default.

The original source documents are not modified. The reviewer remains
responsible for comparing every converted exhibit with its source and checking
the completed PDF before it is filed, served or relied upon.

## Privacy and security model

- Processing takes place on the user's computer.
- There is no account, telemetry, hosted API, remote model call or cloud upload.
- The application blocks non-local network requests.
- Its internal loopback server listens only on `127.0.0.1` and closes with the
  application.
- Automatic recovery metadata can be deleted from the home screen. It contains
  project metadata and local file paths, not copies of source-document content
  or email-attachment bytes. Attachment decisions are stored as hashes and
  provenance only; the parent `.eml` is re-read to rebuild children.

The application is not code-signed, so Windows may show a SmartScreen warning.
Check the SHA-256 published with the release before installation.

## Guided sample

The installer includes a deliberately obvious fictional guided sample. It
demonstrates accepted file types, several exhibits in one paragraph, workbook
sheet selection, an uncited exhibit, a saved email with one text attachment,
and optional cover/index templates. The sample starts on the standard cover
and index; a chosen template is named exactly. It is a product tutorial, not
a realistic witness statement or legal precedent.

The guided-sample button can be hidden and restored. If the packaged sample is
missing, normal use with the user's own documents remains available.

Custom PDF covers, dividers and exhibit covers are proportionally fitted to A4
without cropping. Two cover treatments are mutually exclusive:

- Finish from this template (default): the layout stays as the page
  background. The reviewer can correct a misread name, case number, forum or
  title, and those confirmed corrections are printed on the finished cover
  with the template's inferred alignment (centred party names stay centred).
  Ordinary bundle page numbers apply; a split bundle also receives a volume
  label.
- Use this cover as supplied: names are not rewritten. The page is still
  fitted to A4. A page number or volume label is added only if the reviewer
  ticks that option.

A custom index is never an as-supplied page. Exhibit Builder preserves the
background and table borders, then writes exhibit rows into the columns that
template already has. A Date column is filled only if that template already
has one. The built-in index includes a Date column from the reviewer-editable
document date. If matter text appears on a custom index, confirmed corrections
are printed in the same way as on a finished cover.

Printed page numbers, the index page column and copy-ready statement
suggestions are taken from the finished work-product PDF, not from a second
numbering planner. Copy suggestions uses the desktop clipboard.

On review, the progress strip and statement-reference cards come first. Each
exhibit card shows the document title and **Cited at paragraph**. The bundle
mark still prints as `[AH1/page]`; it is not repeated on the card. Index
description and document date appear once on the main card. Choosing a source
does not confirm it. Confirming a document does not jump the page; the next
included Confirm stays in view. Repeat exhibits are named in the progress
strip and the sticky continuation line. A confirmed email collapses once every
attachment has a choice; **Minimise** remains available. **Add exhibit** is
available from Review, Reconcile and Finalise for an uncited local file; it
does not invent a statement reference. Uncited exhibits are listed at the end
of the copy-ready suggestions. Technical exceptions are recorded after
the cards and are not shown as completed review. Other requirements is a
compact disclosure after the cards.

Bundle cover and index sit as a stable pair. Choosing a cover does not shove
the index off the screen. Finish from this template and Use this cover as
supplied stay in the cover card.

On Finalise, Index heading is a short list of heading names. Named headings
can be dragged as a block, including when collapsed. Drop targets highlight
while a heading is dragged. Exhibits left on No heading still print. Empty
named headings stay saved and do not print. The printed index leaves a gap
before unheaded exhibits that follow a named section. Assigning an exhibit
still moves it to the end of that heading.

When a bundle is split, **Download all volumes (.zip)** is the primary
download. Each volume PDF is available as a secondary download. **Download
bundle PDF** is shown only when there is a single unsplit PDF.

If no custom cover is chosen, a small built-in matter panel (title, parties,
case number and forum) drives the generated cover and index heading.

Every selected custom template must be opened as the exact PDF that will be
used in the bundle. A Word template also requires confirmation of the locally
converted appearance. Possible placeholders and differences between selected
templates are shown for explicit review. These checks are review aids only:
the application never decides which legal details are correct. Source
exhibits and the witness statement remain unchanged. Writing a corrected
party name onto the cover is not a licence to touch an exhibit.

## Requirements

- 64-bit Windows 10 or Windows 11.
- Microsoft Excel is required when XLSX evidence is included, including Excel
  files attached to a local email. The build stops if native Excel printing is
  unavailable; it does not substitute a lower-fidelity spreadsheet reconstruction.
- Node.js 22.13 or later and pnpm 11.9 are required only to build from source.

## Build and test from source

```text
pnpm install --frozen-lockfile
pnpm test
pnpm run package:win
```

`pnpm test` is self-contained and uses only synthetic fixtures committed to
this repository. Maintainers can additionally run `pnpm run test:adversarial`
when the separate local ICC-style stress pack is present beside the repository;
that stress pack is not part of the public source release.

The mandatory release command is:

```text
pnpm run release:verify
```

It enforces TypeScript-aware linting and measured coverage thresholds for the
critical ordering, planning, index-layout and template-persistence modules,
then rebuilds and tests the source, packages Windows, smoke-tests the unpacked
application, temporarily installs it, smoke-tests the installed application and
guided sample (including the local PDF template preview), checks the installed
licence and third-party notices, uninstalls the verification copy and records
the installer SHA-256. A release must not be promoted if any stage fails.

## Supported boundaries

- Witness statement: DOCX only.
- Evidence: PDF, DOCX, EML and XLSX.
- Templates: PDF or DOCX. PDF is recommended where exact layout is critical.
- Scanned PDFs use local OCR. OCR can contain errors.
- DOCX evidence uses a simplified local renderer and may not reproduce floating
  objects, fields, headers or complex Word layout exactly.
- Non-A4 PDF pages can be proportionately converted to A4 or retained at their
  original size. Pages containing PDF annotations are retained rather than
  geometrically altered; unsafe rotated-annotation combinations are blocked.
- Safe internal PDF page links are preserved. Launch actions, JavaScript,
  form submission and other active PDF actions are blocked; flatten or remove
  them in a trusted PDF application, then replace the source file.
- Local `.eml` files can expose their MIME attachments. Each attachment must
  be printed with the parent email, added as its own uncited exhibit, or left
  out. Supported children are PDF, DOCX, EML, XLSX and TXT. Nested `.eml`
  files follow the same rules and are not silently expanded. Unsupported types
  can only be left out. Attachment bytes stay in memory, are bounded (depth 3,
  500 children, 128 MiB extracted), count toward the 500-evidence project cap,
  and are rederived from the parent `.eml` on open or build. They are not
  copied into recovery metadata or the project archive. PST, IMAP and cloud
  mail are not supported. HTML and attachment content are never executed.
  XLSX children, whether printed with the email or added separately, use the
  Sheets stage and locally installed Microsoft Excel; there is no fallback
  spreadsheet reconstruction.
- The application does not provide redaction, legal advice or guaranteed
  compliance with any procedural rule.

### Language support

- English DOCX statement recognition, matching and local OCR are the tested
  release capability. OCR uses the bundled English Tesseract data only.
- The statement parser can retain Unicode text and may find the documented
  square-bracket exhibit placeholders inside statements written partly in
  another language. That does not make matching multilingual.
- Matching vocabulary, date interpretation and inferred descriptions are
  English-oriented. Non-Latin and right-to-left matching are not supported.
- Existing PDF page artwork is embedded without translating it. However,
  automatically generated index text, cover/divider text and labels are not a
  general multilingual typesetting system. The reviewer must not rely on this
  release to generate non-English legal text without checking the finished PDF.
- Full multilingual matching, OCR language packs, fonts and right-to-left
  layout are a separate potential release, not an implied feature of this one.

### Project scale

- A project may contain up to 500 evidence files, including extracted or listed
  email children. The release gate includes a deterministic 200-exhibit,
  long-statement capacity test using text-native PDFs; document pickers load
  their long option list only when opened.
- PDFs are limited to 256 MiB and 1,500 pages per source. Automatic OCR has
  lower limits (100 MiB and 200 pages) so a scanned source cannot exhaust the
  desktop process; an affected source requires the existing explicit visual-
  review route.
- Saved project archives are limited to 192 MiB compressed, 256 MiB expanded
  and 128 MiB per embedded source. A large matter can therefore be analysable
  while still being too large to save as one portable project archive.
- A volume page target counts the volume cover, the complete repeated index and
  exhibit pages. An exhibit is never split merely to hit the target. Every
  volume contains the full index; index links and child bookmarks point only to
  exhibits physically present in that volume.

## Reports, security and contributions

- Use GitHub Issues for non-confidential defects and usability reports.
- Use GitHub private vulnerability reporting for suspected security problems;
  see [SECURITY.md](SECURITY.md).
- Pull requests are not accepted unless invited under agreed written terms; see
  [CONTRIBUTING.md](CONTRIBUTING.md).
- Third-party component terms are reproduced in
  [THIRD_PARTY_LICENSES.txt](THIRD_PARTY_LICENSES.txt).
