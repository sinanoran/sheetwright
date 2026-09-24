namespace Sheetwright;

public sealed class SpreadsheetColumnStyle
{
    private readonly SpreadsheetWorksheet _worksheet;
    private readonly int _column;
    private SpreadsheetColumnNumberFormatStyle? _numberFormat;

    internal SpreadsheetColumnStyle(SpreadsheetWorksheet worksheet, int column)
    {
        _worksheet = worksheet;
        _column = column;
    }

    public SpreadsheetColumnNumberFormatStyle NumberFormat => _numberFormat ??= new SpreadsheetColumnNumberFormatStyle(_worksheet, _column);

    public void SetWrapText(bool wrapText)
    {
        _worksheet.ApplyColumnStyle(_column, new SpreadsheetStylePatch { WrapText = wrapText });
    }
}
