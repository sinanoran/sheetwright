namespace Sheetwright;

public sealed class SpreadsheetColumnNumberFormatStyle
{
    private readonly SpreadsheetWorksheet _worksheet;
    private readonly int _column;

    internal SpreadsheetColumnNumberFormatStyle(SpreadsheetWorksheet worksheet, int column)
    {
        _worksheet = worksheet;
        _column = column;
    }

    public void SetFormat(string format)
    {
        if (format is null)
        {
            throw new ArgumentNullException(nameof(format));
        }

        _worksheet.ApplyColumnStyle(
            _column,
            new SpreadsheetStylePatch { NumberFormat = SpreadsheetTextSanitizer.Sanitize(format) });
    }
}
