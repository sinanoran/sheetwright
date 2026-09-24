namespace Sheetwright;

public sealed class SpreadsheetColumns
{
    private readonly SpreadsheetWorksheet _worksheet;

    internal SpreadsheetColumns(SpreadsheetWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public SpreadsheetColumn this[int column] => new(_worksheet, column);
}
