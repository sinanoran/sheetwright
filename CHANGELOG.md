# Changelog

All notable changes to this project are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] - 2026-10-01

### Added

- `SpreadsheetStreamWriter` — a single worksheet written straight to a stream,
  one row at a time, for an export too large to hold in memory.
  `SpreadsheetPackage` keeps every cell until it is asked for the bytes, which is
  what lets styles merge and columns auto-fit, and what runs a worker out of
  memory on a report of several hundred thousand rows. The streaming writer keeps
  nothing but the row in hand: text goes in as an inline string so there is no
  shared string table, and the bytes reach the destination as the rows arrive.
  Rows and columns go in in ascending order, columns and tables are declared
  before the first row, and `Complete()` finishes the workbook.
- A table added to a streamed sheet may be left open at the bottom — pass `null`
  for its last row and it ends at the last row written, because the row count is
  the one thing a streaming caller rarely knows in advance.

### Fixed

- `NaN` and the infinities are written as text rather than as numbers. The file
  format has one numeric type and no spelling for any of the three, so
  `<v>NaN</v>` produced a workbook Excel offered to repair rather than open. An
  average over an empty set is the usual way a caller arrives there.

### Changed

- A streamed table and a table on an in-memory sheet resolve their column names
  through the same code, so an empty heading becomes its column letter and a
  repeated one gains a suffix identically in both — and in both the resolved
  heading is what ends up in the header cell, which is what stops Excel offering
  to repair the file.

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

[Unreleased]: https://github.com/sinanoran/sheetwright/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/sinanoran/sheetwright/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/sinanoran/sheetwright/releases/tag/v0.1.0
