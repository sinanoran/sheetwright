using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace Sheetwright;

/// <summary>
/// Provides typed cell reading for a single row during spreadsheet reading.
/// </summary>
public sealed class SpreadsheetRowReader
{
    private readonly Row _row;
    private readonly string[] _sharedStrings;
    private readonly Dictionary<int, Cell> _cellsByColumn = [];

    internal SpreadsheetRowReader(Row row, string[] sharedStrings)
    {
        _row = row;
        _sharedStrings = sharedStrings;
        IndexCells();
    }

    public int RowIndex => (int)(_row.RowIndex?.Value ?? 0U);

    public bool HasValue(int column) => _cellsByColumn.ContainsKey(column);

    public bool HasValue(string columnLetter) => HasValue(SpreadsheetAddress.GetColumnNumber(columnLetter));

    public string? GetText(int column)
    {
        if (!_cellsByColumn.TryGetValue(column, out Cell? cell) || cell.CellValue?.InnerText is not string raw)
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

    public string? GetText(string columnLetter) => GetText(SpreadsheetAddress.GetColumnNumber(columnLetter));

    public decimal? GetDecimal(int column)
    {
        if (!_cellsByColumn.TryGetValue(column, out Cell? cell) || cell.CellValue?.InnerText is not string raw)
        {
            return null;
        }

        return decimal.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result)
            ? result
            : null;
    }

    public decimal? GetDecimal(string columnLetter) => GetDecimal(SpreadsheetAddress.GetColumnNumber(columnLetter));

    public int? GetInt(int column)
    {
        if (!_cellsByColumn.TryGetValue(column, out Cell? cell) || cell.CellValue?.InnerText is not string raw)
        {
            return null;
        }

        return int.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out int result)
            ? result
            : null;
    }

    public int? GetInt(string columnLetter) => GetInt(SpreadsheetAddress.GetColumnNumber(columnLetter));

    public DateTime? GetDateTime(int column)
    {
        if (!_cellsByColumn.TryGetValue(column, out Cell? cell) || cell.CellValue?.InnerText is not string raw)
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

    public DateTime? GetDateTime(string columnLetter) => GetDateTime(SpreadsheetAddress.GetColumnNumber(columnLetter));

    public bool? GetBoolean(int column)
    {
        if (!_cellsByColumn.TryGetValue(column, out Cell? cell) || cell.CellValue?.InnerText is not string raw)
        {
            return null;
        }

        return raw is "1" or "true" or "True";
    }

    public bool? GetBoolean(string columnLetter) => GetBoolean(SpreadsheetAddress.GetColumnNumber(columnLetter));

    private void IndexCells()
    {
        foreach (Cell cell in _row.Elements<Cell>())
        {
            if (cell.CellReference?.Value is string reference)
            {
                SpreadsheetAddress parsed = SpreadsheetAddress.Parse(reference);
                _cellsByColumn[parsed.FromColumn] = cell;
            }
        }
    }
}
