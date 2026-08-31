# Exhibit Builder

Exhibit Builder is an offline Windows desktop application for turning exhibit
references in a final DOCX witness statement into a reviewed, indexed and
paginated PDF exhibit bundle.

[Download the latest official Windows release](https://github.com/kinhoe107/exhibit-builder/releases/latest)

> **Licence:** Exhibit Builder is free for individual use under the
> [End User Licence Agreement](build/EULA.txt). This public repository contains
> proprietary source code made available for inspection and security review. It
> is not open-source software, and public access does not grant permission to
> compile, run, modify or redistribute the source. See the
> [Source Code Notice](LICENSE).

## What it does

- Finds exhibit references and document descriptions in a final DOCX witness
  statement.
- Accepts PDF, DOCX, EML and XLSX evidence.
- Presents proposed document matches for confirmation rather than silently
  deciding them.
- Builds a paginated PDF bundle with an index, bookmarks and internal index
  links.
- Supports optional cover, index, exhibit-cover and divider templates.
- Handles local email attachments and native Microsoft Excel printing.
- Can divide large bundles into separate physical volumes while retaining
  continuous pagination.
- Records source and output hashes, exclusions, warnings and reviewer decisions
  in a JSON manifest.

## Fidelity and review

Exhibit Builder does not modify the witness statement or original evidence
files. Converted exhibits and the completed bundle must still be checked against
their sources before the bundle is filed, served or relied upon.

Confirmed exhibit order controls index numbering and bookmark titles. Printed
page numbers, index references and copy-ready statement suggestions are taken
from the finished PDF.

Potentially active PDF content—including JavaScript, launch actions and form
submission—is blocked. Unsupported or unsafe material is not silently accepted.

## Privacy

Processing takes place on the user's computer.

Exhibit Builder has:

- no user account;
- no telemetry;
- no cloud upload;
- no hosted API or remote AI call; and
- no access to cloud mail or remote document stores.

The application blocks non-local network requests. Local recovery information
can be deleted from the home screen.

## Requirements

- 64-bit Windows 10 or Windows 11.
- Microsoft Excel when XLSX evidence is included, including spreadsheets
  attached to an email.
- A final witness statement in DOCX format.

The tested recognition, matching and OCR capability is English. Scanned
documents and converted Word files require careful visual checking. Exhibit
Builder does not provide redaction, legal advice or guaranteed compliance with
any court, tribunal or procedural rule.

## Download and verification

Use only the installer published on the
[official GitHub Releases page](https://github.com/kinhoe107/exhibit-builder/releases/latest).

The application is not code-signed, so Windows may display a SmartScreen
warning. Compare the installer's SHA-256 checksum with the checksum published
alongside the release before installing it.

Official releases are subjected to automated source, coverage, packaging,
installation, guided-sample and adversarial testing. A release is not promoted
if the maintainer's verification process fails.

## Security and reports

- Use [GitHub Issues](https://github.com/kinhoe107/exhibit-builder/issues) for
  non-confidential defects and usability reports.
- Use
  [GitHub private vulnerability reporting](https://github.com/kinhoe107/exhibit-builder/security)
  for suspected security problems.
- See [SECURITY.md](SECURITY.md) for the security policy.
- See [THIRD_PARTY_LICENSES.txt](THIRD_PARTY_LICENSES.txt) for third-party
  component notices.
- Pull requests are accepted only by prior invitation under agreed written
  terms; see [CONTRIBUTING.md](CONTRIBUTING.md).
