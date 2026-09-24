# Changelog

All notable changes to this project are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-09-24

First public release. The library was extracted from a private codebase where it
had been writing exports in production, so the surface below is the one that
survived that use rather than a first draft.

### Added

- `SpreadsheetPackage` — an in-memory OPC package, written as `.xlsx` with no
  temporary file, read back through the same API.
- Worksheets, cells and ranges addressed either as `[row, column]` or as `"A1"`
  and `"A1:C9"`, with everything outside Excel's grid refused rather than written.
- Styles: font, fill, borders, number format, alignment, wrapping and cell
  locking, applied per sheet, per column, per row or per cell, merged in that
  order and de-duplicated into one style table.
- Data validation: dropdowns, numeric, date and time bounds, text length, and
  input or error messages.
- Tables, defined names, pictures (PNG, JPEG, GIF, BMP, TIFF), freeze panes,
  auto-fit columns, sheet protection and macro-enabled workbooks.
- `SpreadsheetGrid` — one column description, written as either `.xlsx` or CSV,
  with CSV formula injection neutralised.
- Targets `net8.0`, `net9.0` and `net10.0`.

[Unreleased]: https://github.com/sinanoran/sheetwright/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/sinanoran/sheetwright/releases/tag/v0.1.0
