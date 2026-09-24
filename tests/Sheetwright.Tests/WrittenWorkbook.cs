using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Sheetwright.Tests;

/// <summary>
/// A workbook that has been written and then opened again.
/// </summary>
/// <remarks>
/// <para>
/// Every test here writes a real package and reads it back, because the thing
/// under test is a file format. An assertion against the in-memory object graph
/// proves the library remembered what it was told; only the bytes prove it
/// produced something Excel will open.
/// </para>
/// <para>
/// Two readers, on purpose. <see cref="SpreadsheetPackage"/> is used wherever the
/// question is about behaviour a caller can observe. DocumentFormat.OpenXml is
/// used for the rest — the schema order, the style table, the shared strings —
/// which is exactly the part no public API exposes and Excel refuses to guess at.
/// </para>
/// </remarks>
internal sealed class WrittenWorkbook : IDisposable
{
    private readonly MemoryStream _stream;
    private readonly SpreadsheetDocument _document;

    private WrittenWorkbook(byte[] bytes)
    {
        Bytes = bytes;
        _stream = new MemoryStream(bytes, writable: false);
        _document = SpreadsheetDocument.Open(_stream, isEditable: false);
    }

    /// <summary>Builds a workbook, writes it, and opens what was written.</summary>
    internal static WrittenWorkbook From(Action<SpreadsheetPackage> build)
    {
        using SpreadsheetPackage package = new();
        build(package);
        return new WrittenWorkbook(package.GetAsByteArray());
    }

    /// <summary>Opens bytes that were written somewhere else.</summary>
    /// <remarks>
    /// For <see cref="SpreadsheetGrid"/>, which finalises its own package and
    /// hands back an array rather than letting a test build one.
    /// </remarks>
    internal static WrittenWorkbook Opening(byte[] bytes) => new(bytes);

    internal byte[] Bytes { get; }

    internal WorkbookPart WorkbookPart => _document.WorkbookPart
        ?? throw new InvalidOperationException("The written package has no workbook part.");

    internal Workbook Workbook => WorkbookPart.Workbook!;

    internal Stylesheet Stylesheet => WorkbookPart.WorkbookStylesPart?.Stylesheet
        ?? throw new InvalidOperationException("The written package has no stylesheet.");

    internal Worksheet Sheet(string name)
    {
        Sheet sheet = WorkbookPart.Workbook!.Sheets?.Elements<Sheet>()
                .FirstOrDefault(x => string.Equals(x.Name?.Value, name, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"The written package has no sheet named '{name}'.");

        return ((WorksheetPart)WorkbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;
    }

    internal Cell? Cell(string sheetName, string reference) =>
        Sheet(sheetName).GetFirstChild<SheetData>()?.Elements<Row>()
            .SelectMany(row => row.Elements<Cell>())
            .FirstOrDefault(cell => string.Equals(cell.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));

    /// <summary>The displayed text of a cell, with a shared string resolved.</summary>
    internal string? Text(string sheetName, string reference)
    {
        Cell? cell = Cell(sheetName, reference);
        string? raw = cell?.CellValue?.InnerText;
        if (raw is null)
        {
            return null;
        }

        return cell!.DataType?.Value == CellValues.SharedString
            ? SharedStrings[int.Parse(raw, System.Globalization.CultureInfo.InvariantCulture)]
            : raw;
    }

    internal IReadOnlyList<string> SharedStrings =>
        [.. WorkbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>()
            .Select(item => item.InnerText) ?? []];

    /// <summary>The cell format a cell's style index points at.</summary>
    internal CellFormat FormatOf(Cell cell)
    {
        int index = (int)(cell.StyleIndex?.Value ?? 0U);
        return Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt(index)!;
    }

    /// <summary>The local names of a worksheet's children, in document order.</summary>
    /// <remarks>
    /// CT_Worksheet is a sequence, not a choice: elements out of order are a file
    /// Excel offers to repair rather than open. Nothing in the public API shows
    /// this, so the order is asserted directly.
    /// </remarks>
    internal IReadOnlyList<string> ChildOrder(string sheetName) =>
        [.. Sheet(sheetName).ChildElements.Select(element => element.LocalName)];

    public void Dispose()
    {
        _document.Dispose();
        _stream.Dispose();
    }
}
