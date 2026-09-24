namespace Sheetwright;

public sealed class SpreadsheetWorksheetView
{
    private readonly SpreadsheetWorksheet _worksheet;

    internal SpreadsheetWorksheetView(SpreadsheetWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public void FreezePanes(int row, int column)
    {
        _worksheet.FreezePanes(row, column);
    }
}
