# Changelog

All notable changes to this project are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.4.0] - 2026-11-20

### Added

- Readers for files this library did not write, for an import: `SpreadsheetWorkbookReader` reads any
  `.xlsx` a row at a time and `SpreadsheetCsvReader` reads any CSV a row at a time. Both return
  `SpreadsheetRow`s of `SpreadsheetCellValue`s (text, number as `decimal`, boolean, date, error, empty) and
  are configured by `SpreadsheetReaderOptions`. Neither loads the file: the workbook reader advances an
  XML reader over the worksheet and holds only the shared-string table, and the CSV reader holds one record.
- Workbook reading: dates are recognised by their cell number format (built-in and custom formats; the 1900
  system's non-existent 29 February and the 1904 system are both handled); a formula is read as its stored
  value and never evaluated; rich text is joined without its phonetic guide; inline strings, booleans,
  errors and ISO dates are read; the strict namespaces read like the transitional ones; sheets are listed
  with their visibility; cells with no reference follow one another.
- CSV reading: RFC 4180 quoting including line breaks in quotes; CRLF, LF and lone CR; the encoding from a
  UTF-8, UTF-16 LE or UTF-16 BE byte-order mark (UTF-8 otherwise, strictly: invalid bytes are refused, not
  replaced) or from the options; the separator from a `sep=` line or by measuring the first rows (comma,
  semicolon, tab, pipe); a stream that cannot seek.
- The apostrophe rule: the CSV reader takes off the one apostrophe `SpreadsheetGridCsv` puts in front of
  text that starts with `=`, `+`, `-`, `@`, a tab, a carriage return or a full-width form of the first
  four, so a file the writer wrote reads back as the values that went in (`UnescapeFormulaGuard`). The
  writer and the reader share one list of leader characters. One text cannot be told apart from a guarded
  one: a text that itself begins with an apostrophe and then a leader.
- Limits, checked while the file is read and before the memory is allocated, each refused with a
  `SpreadsheetReadLimitException` that names it: `MaxUncompressedBytes`, `MaxCompressionRatio`,
  `MaxSharedStrings`, `MaxSharedStringCharacters`, `MaxParts`, `MaxRows`, `MaxColumns` and
  `MaxCellCharacters`. No document type declaration is processed, no external resource is read, a
  relationship that leaves the package is refused, and a tag, comment or run of text longer than any cell
  can be is refused before the XML reader holds it.
- `net11.0` is a target framework.

### Changed

- Built with the .NET 11 SDK; `xunit.v3` 4.0.3; the GitHub Actions the workflows use are at their latest
  major versions. `DocumentFormat.OpenXml` stays at 3.5.1, the latest stable.

## [0.3.2] - 2026-11-17

### Fixed

- CSV formula guard: text beginning with the full-width forms of `=`, `+`, `-`
  and `@` (U+FF1D, U+FF0B, U+FF0D, U+FF20) is now prefixed with an apostrophe
  like the ASCII forms, in values and headings.

## [0.3.1] - 2026-10-11

### Changed

- The README and the package description describe the API on its own terms
  rather than by comparison with another library. No code changes.

## [0.3.0] - 2026-10-11

### Added

- Conditional formatting — `sheet.ConditionalFormatting` and
  `range.ConditionalFormatting`. Cells can be formatted by comparing their value
  (eight comparisons, `Between` and `NotBetween` inclusive), by a formula, as
  duplicate or unique values, or shaded on a two- or three-colour scale whose
  points can be anchored to the lowest or highest value, a number, a percentage,
  a percentile or a formula. Every rule but a colour scale carries a font
  (bold, italic, underline, strikethrough, colour), a fill, borders and a number
  format, and applies only what was set — so turning bold off is written as a
  change rather than left out. Values to compare with are typed as cells are and
  written in the invariant culture; identical formatting resolves to one entry in
  the differential format table across every sheet; rules take priority in the
  order they were added. A comparison with nothing to compare against, or a
  second bound on a comparison that takes one, is refused before any sheet is
  committed, so the rule can be fixed and the workbook written after all.
- `SpreadsheetReader` and `SpreadsheetRowReader` — typed reading of existing workbooks,
  cell coordinates, sheet names, and tabular ranges with typed getters for text,
  integers, decimals, doubles, booleans, and dates.
- `SpreadsheetTemplateHydrator` — stamps typed values and collections into pre-formatted
  template workbooks by coordinate or defined name while preserving layouts, formulas,
  and formatting.
- `SpreadsheetNumberFormats` — standard invariant number and currency format string presets.

### Fixed

- Table header cells containing numbers or dates are coerced to strings so OpenXML
  `<tableColumn name="...">` and `<c t="s">` types match, preventing Excel repair warnings.

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

[Unreleased]: https://github.com/sinanoran/sheetwright/compare/v0.4.0...HEAD
[0.4.0]: https://github.com/sinanoran/sheetwright/compare/v0.3.2...v0.4.0
[0.3.2]: https://github.com/sinanoran/sheetwright/compare/v0.3.1...v0.3.2
[0.3.1]: https://github.com/sinanoran/sheetwright/compare/v0.3.0...v0.3.1
[0.3.0]: https://github.com/sinanoran/sheetwright/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/sinanoran/sheetwright/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/sinanoran/sheetwright/releases/tag/v0.1.0
