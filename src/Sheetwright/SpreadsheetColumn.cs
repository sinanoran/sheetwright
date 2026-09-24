namespace Sheetwright;

public sealed class SpreadsheetColumn
{
    private readonly SpreadsheetWorksheet _worksheet;
    private readonly int _column;
    private SpreadsheetColumnStyle? _style;

    internal SpreadsheetColumn(SpreadsheetWorksheet worksheet, int column)
    {
        _worksheet = worksheet;
        _column = column;
    }

    public SpreadsheetColumnStyle Style => _style ??= new SpreadsheetColumnStyle(_worksheet, _column);

    public void SetWidth(double width)
    {
        _worksheet.SetColumnWidth(_column, width);
    }

    public void SetHidden(bool hidden)
    {
        _worksheet.SetColumnHidden(_column, hidden);
    }

    public void AutoFit()
    {
        _worksheet.AutoFitColumn(_column);
    }
}
