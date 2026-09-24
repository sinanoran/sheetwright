# Sheetwright

[![CI](https://github.com/sinanoran/sheetwright/actions/workflows/ci.yml/badge.svg)](https://github.com/sinanoran/sheetwright/actions/workflows/ci.yml)
[![NuGet](https://img.shields.io/nuget/v/Sheetwright.svg)](https://www.nuget.org/packages/Sheetwright)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

**Sheetwright writes `.xlsx` files, and reads back the ones it wrote.**

It is a small API over [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK),
which is MIT-licensed and maintained by Microsoft. Sheetwright is MIT-licensed
too — including for commercial use, with no second licence to buy when the
project stops being a hobby.

```bash
dotnet add package Sheetwright
```

Targets `net8.0`, `net9.0` and `net10.0`. One dependency.

## Why it exists

Every administrative product is asked for an export, and "export" almost always
means Excel rather than CSV — because the person asking wants column widths, a
frozen header row, a date that reads as a date, and a dropdown on the column they
are going to send back filled in. CSV gives you none of that, and a
culture-dependent decimal separator besides.

The surface deliberately reads like EPPlus's — `package.Workbook.Worksheets.Add`,
`sheet.Cells[1, 1].Value` — because that is the vocabulary anybody who has
written an Excel export already has. The difference is the licence.

## Writing a workbook

```csharp
using Sheetwright;

using SpreadsheetPackage package = new();

SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Invoices");
sheet.Cells[1, 1].Value = "Reference";
sheet.Cells[1, 2].Value = "Issued";
sheet.Cells[1, 3].Value = "Amount";
sheet.Cells["A1:C1"].Style.Font.SetBold(true);

int row = 2;
foreach (Invoice invoice in invoices)
{
    sheet.Cells[row, 1].Value = invoice.Reference;
    sheet.Cells[row, 2].Value = invoice.IssuedOn;
    sheet.Cells[row, 3].Value = invoice.Amount;
    row++;
}

sheet.Column(3).Style.NumberFormat.SetFormat("#,##0.00");
sheet.View.FreezePanes(2, 1);
sheet.Cells.AutoFitColumns();

byte[] workbook = package.GetAsByteArray();
```

`GetAsByteArray` and `SaveAs` both **finalise** the package: styles are resolved,
the shared string table is closed, the zip is written. Calling either again
returns the same bytes, and any further edit throws. Build the whole workbook,
then ask for it.

A collection can go in wholesale, which is the usual shape of an export:

```csharp
sheet.Cells[1, 1].LoadFromCollection(rows, printHeaders: true);
```

Two things about `LoadFromCollection` that matter in practice. Column order
follows `Type.GetProperties()`, which the runtime does not guarantee — project
into a `record` whose properties are the columns you want, in the order you want
them, and do not rely on that record's shape by accident. And the header text is
the property name, so `IssuedOn` appears as `IssuedOn`; write the header row
yourself when a person is going to read it.

### From an ASP.NET Core endpoint

The bytes are a file like any other:

```csharp
app.MapGet("/invoices/export", async (InvoiceQueries queries, CancellationToken cancellationToken) =>
{
    IReadOnlyList<InvoiceRow> rows = await queries.ForExportAsync(cancellationToken);

    using SpreadsheetPackage package = new();
    SpreadsheetWorksheet sheet = package.Workbook.Worksheets.Add("Invoices");
    sheet.Cells[1, 1].LoadFromCollection(rows, printHeaders: true);

    return Results.File(package.GetAsByteArray(), SpreadsheetGrid.ContentType, "invoices.xlsx");
});
```

The whole workbook is held in memory, so an export of unbounded size is an
unbounded allocation. Page the query, cap the row count, or move the work to a
background job and hand back a stored file.

## Exporting a grid

`SpreadsheetGrid` describes columns once and writes them as either format:

```csharp
byte[] workbook = SpreadsheetGrid.Write(
    "Members",
    [
        new SpreadsheetGridColumn<Member>("Name", row => row.DisplayName),
        new SpreadsheetGridColumn<Member>("Email", row => row.Email),
        new SpreadsheetGridColumn<Member>("Joined", row => row.JoinedOn)
        {
            NumberFormat = "yyyy-mm-dd hh:mm",
        },
    ],
    rows);
```

`SpreadsheetGrid.TryWriteAs(format, ...)` is the endpoint-facing form: it takes
`"xlsx"` or `"csv"`, and returns `false` for a format it does not know rather
than throwing, so the caller raises its own validation error. Falling back to
Excel for an unrecognised format would be the friendlier spelling and the wrong
one — a typo would return a workbook to something expecting CSV, which fails
later and somewhere else.

Return the underlying value from a column rather than a string wherever there is
one. A date written as text sorts April before January, and a number written as
text cannot be summed, which is most of why the person asked for Excel instead
of CSV.

**CSV escaping here is a security control, not just a formatting one.** A cell in
a workbook is typed and Excel never evaluates a string; CSV has no types, so a
display name of `=1+1` becomes a formula and one calling `DDE` is code running on
the machine of whoever opened the file. Text beginning `=`, `+`, `-`, `@`, tab or
carriage return is prefixed with an apostrophe, which is what Excel itself writes
to mean "this is text". Numbers and dates are written by the library and are not
touched, so a negative number keeps its minus sign.

## Values, and how they are typed

| Written | Stored as | Read back as |
|---|---|---|
| `string`, `char` | shared string | `string` |
| `int`, `long`, `decimal`, `double`, … | number | `double` |
| `bool` | boolean | `bool` |
| `DateTime` | serial number, formatted `dd/mm/yyyy` | `DateTime` |
| `null`, `DBNull.Value` | empty cell | `null` |
| anything else | shared string of `Convert.ToString` | `string` |

Numbers all come back as `double`: the file format has one numeric type and the
reader does not know what the writer started with. A `decimal` written and read
is a `decimal` no longer, which matters if you round-trip money through a
workbook — don't.

A `DateTime` written with no explicit number format gets `dd/mm/yyyy`, because a
date cell with no format shows Excel's serial number and the export looks broken.
Set a format yourself and yours wins.

Everything is converted with `InvariantCulture`, including the fallback that
renders a type the writer does not know. A workbook is a file rather than a
screen: the same object has to produce the same bytes on a Turkish machine and an
American one, or an export becomes a function of who ran it. To render a value
for a particular reader, format it yourself and pass the string.

## What it refuses

An address outside Excel's grid throws rather than being written:

```csharp
sheet.Cells["A0"].Value = "x";        // ArgumentOutOfRangeException
sheet.Cells["C1:A1"].Merge = true;    // ArgumentOutOfRangeException
```

Rows run from 1 to 1,048,576 and columns from A to XFD. Row zero used to travel
all the way through and be written as `<row r="0">`, which is not a row Excel has
— the caller got a package produced without error that could not be opened. An
inverted range is the quieter version of the same fault: every loop over it runs
zero times, so styling or merging one did nothing at all and reported success.

A validation address is parsed too, so `"Sheet1!A1"` (a sheet qualifier is not
legal in `sqref`) and anything mistyped fail on the line that wrote them. A list
of rectangles is still accepted, which is how one dropdown reaches two columns:

```csharp
sheet.DataValidations.AddAnyValidation("B1:B9 E1:E9");
```

## What else it does

- **Styles** — font, fill, borders, number format, alignment, wrapping and cell
  locking, applied per sheet, per column, per row or per cell. The layers merge in
  that order, so a cell format beats the column it sits in. Identical formatting
  resolves to one entry in the style table; that is what keeps a large export
  openable.
- **Data validation** — dropdowns, numeric and date bounds, text length, and
  input or error messages. The workbook you send out can enforce its own rules
  before it comes back.
- **Tables** — a range promoted to an Excel table with a style and an autofilter.
  Column names come from the first row of the range. An empty heading becomes its
  column letter and a repeated one gains a suffix, because Excel refuses
  duplicates — and in both cases the resolved heading is written back into the
  header cell, since Excel checks the two against each other and offers to repair
  the file when they disagree.
- **Pictures** — PNG, JPEG, GIF, BMP and TIFF, sized from the image header.
- **Defined names** — a name for a range, which is how a dropdown on one sheet
  refers to a list on another.
- **Sheet protection** — a locked sheet with unlocked cells for the parts you
  want filled in. The password uses Excel's own 16-bit hash: it stops a
  well-meaning person editing the wrong cell, and **it is not a security
  control**. Do not put anything behind it that matters.
- **Macro-enabled workbooks** — `new SpreadsheetPackage(macroEnabled: true)` and
  `SetVbaProject`, for a template that has to carry an existing VBA project.

## What it deliberately does not do

- **Calculate.** A formula is stored as text; Excel evaluates it on open. Reading
  a formula cell out of a package this library wrote gives you the formula, not a
  value.
- **Edit an existing workbook.** `new SpreadsheetPackage(stream)` opens
  read-only, and every write throws. Reading is there so a test — or an import —
  can look at what was produced.
- **Charts, pivot tables, conditional formatting, comments.** None of them are
  here. Each is a fair thing to add when something needs one, and none should be
  added on the argument that a spreadsheet library ought to have it.

## How it is built

One consequence of the design is worth knowing before you read the code: the OPC
package is assembled by hand. `SpreadsheetMemoryPackage` is a
`System.IO.Packaging.Package` over a `MemoryStream`, and
`SpreadsheetPackageArchiveWriter` writes the zip and `[Content_Types].xml`
itself. That is not incidental complexity — it is what keeps the whole workbook
in memory with no temporary file — but it does mean the file format is this
repository's responsibility rather than the SDK's.

That is why `tests/Sheetwright.Tests` opens almost everything it writes. 241
tests, no container, no database; they run anywhere the SDK does.

```bash
dotnet build Sheetwright.slnx
```

## Contributing

Issues and pull requests are welcome. A change to what gets written needs a test
that opens the result — through the library for behaviour, through
`DocumentFormat.OpenXml` for the parts of the XML that Excel refuses to open and
no public API exposes.

## Licence

MIT. See [LICENSE](LICENSE).
