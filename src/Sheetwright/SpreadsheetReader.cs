using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace Sheetwright;

/// <summary>
/// Reads typed data from Excel workbooks, worksheets, and individual cells.
/// </summary>
public sealed class SpreadsheetReader : IDisposable
{
    private readonly MemoryStream _stream;
    private readonly SpreadsheetDocument _document;
    private readonly WorkbookPart _workbookPart;
    private readonly Workbook _openXmlWorkbook;
    private readonly string[] _sharedStrings;
    private bool _disposed;

    public SpreadsheetReader(Stream stream)
    {
        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        _stream = new MemoryStream();
        if (stream.CanSeek)
        {
            stream.Position = 0;
        }
        stream.CopyTo(_stream);
        _stream.Position = 0;

        _document = SpreadsheetDocument.Open(_stream, isEditable: false);
        _workbookPart = _document.WorkbookPart
            ?? throw new InvalidDataException("The spreadsheet is missing a WorkbookPart.");
        _openXmlWorkbook = _workbookPart.Workbook
            ?? throw new InvalidDataException("The spreadsheet is missing its Workbook.");

        _sharedStrings = _workbookPart.SharedStringTablePart?.SharedStringTable?
            .Elements<SharedStringItem>()
            .Select(x => x.InnerText)
            .ToArray() ?? [];
    }

    public SpreadsheetReader(byte[] bytes)
        : this(new MemoryStream(bytes ?? throw new ArgumentNullException(nameof(bytes))))
    {
    }

    public IReadOnlyList<string> SheetNames =>
        _openXmlWorkbook.Sheets?.Elements<Sheet>()
            .Select(s => s.Name?.Value ?? string.Empty)
            .Where(name => !string.IsNullOrEmpty(name))
            .ToList() ?? [];

    public string? GetText(string sheetName, string cellReference)
    {
        Cell? cell = FindCell(sheetName, cellReference);
        return ExtractText(cell);
    }

    public decimal? GetDecimal(string sheetName, string cellReference)
    {
        Cell? cell = FindCell(sheetName, cellReference);
        if (cell?.CellValue?.InnerText is not string raw)
        {
            return null;
        }

        return decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result)
            ? result
            : null;
    }

    public int? GetInt(string sheetName, string cellReference)
    {
        Cell? cell = FindCell(sheetName, cellReference);
        if (cell?.CellValue?.InnerText is not string raw)
        {
            return null;
        }

        return int.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out int result)
            ? result
            : null;
    }

    public DateTime? GetDateTime(string sheetName, string cellReference)
    {
        Cell? cell = FindCell(sheetName, cellReference);
        if (cell?.CellValue?.InnerText is not string raw)
        {
            return null;
        }

        if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double oaDate))
        {
            try
            {
                return DateTime.FromOADate(oaDate);
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        return DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsed)
            ? parsed
            : null;
    }

    public bool? GetBoolean(string sheetName, string cellReference)
    {
        Cell? cell = FindCell(sheetName, cellReference);
        if (cell?.CellValue?.InnerText is not string raw)
        {
            return null;
        }

        return raw is "1" or "true" or "True";
    }

    public IReadOnlyList<T> ReadRows<T>(
        string sheetName,
        int startRow,
        int? endRow,
        Func<SpreadsheetRowReader, T> mapRow)
    {
        EnsureNotDisposed();

        if (mapRow is null)
        {
            throw new ArgumentNullException(nameof(mapRow));
        }

        WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
        Worksheet worksheet = worksheetPart.Worksheet
            ?? throw new InvalidDataException($"Worksheet '{sheetName}' is missing content.");

        SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData is null)
        {
            return [];
        }

        List<T> results = [];
        int maxRow = endRow ?? int.MaxValue;

        foreach (Row row in sheetData.Elements<Row>())
        {
            uint rowIndex = row.RowIndex?.Value ?? 0U;
            if (rowIndex < (uint)startRow || rowIndex > (uint)maxRow)
            {
                continue;
            }

            SpreadsheetRowReader reader = new(row, _sharedStrings);
            results.Add(mapRow(reader));
        }

        return results;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _document.Dispose();
        _stream.Dispose();
        _disposed = true;
    }

    private void EnsureNotDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    private WorksheetPart GetWorksheetPart(string sheetName)
    {
        EnsureNotDisposed();

        Sheet? sheet = _openXmlWorkbook.Sheets?.Elements<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));

        if (sheet is null || sheet.Id?.Value is null)
        {
            throw new KeyNotFoundException($"Worksheet '{sheetName}' was not found in the workbook.");
        }

        return (WorksheetPart)_workbookPart.GetPartById(sheet.Id.Value);
    }

    private Cell? FindCell(string sheetName, string cellReference)
    {
        EnsureNotDisposed();

        WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
        Worksheet worksheet = worksheetPart.Worksheet
            ?? throw new InvalidDataException($"Worksheet '{sheetName}' is missing content.");

        SpreadsheetAddress address = SpreadsheetAddress.Parse(cellReference);
        uint rowIndex = (uint)address.FromRow;

        SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
        Row? row = sheetData?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);

        string targetRef = $"{SpreadsheetAddress.GetColumnName(address.FromColumn)}{address.FromRow}";
        return row?.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, targetRef, StringComparison.OrdinalIgnoreCase));
    }

    private string? ExtractText(Cell? cell)
    {
        if (cell?.CellValue?.InnerText is not string raw)
        {
            return null;
        }

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            return int.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out int index)
                && index >= 0 && index < _sharedStrings.Length
                ? _sharedStrings[index]
                : raw;
        }

        return raw;
    }
}
