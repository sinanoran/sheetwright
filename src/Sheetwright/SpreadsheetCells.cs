namespace Sheetwright;

public sealed class SpreadsheetCells
{
    private readonly SpreadsheetWorksheet _worksheet;
    private SpreadsheetWorksheetDefaultStyle? _style;

    internal SpreadsheetCells(SpreadsheetWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public SpreadsheetWorksheetDefaultStyle Style => _style ??= new SpreadsheetWorksheetDefaultStyle(_worksheet);

    public SpreadsheetRange this[int row, int column] => new(_worksheet, new SpreadsheetAddress(row, column, row, column));

    public SpreadsheetRange this[int fromRow, int fromColumn, int toRow, int toColumn] => new(_worksheet, new SpreadsheetAddress(fromRow, fromColumn, toRow, toColumn));

    public SpreadsheetRange this[string address] => new(_worksheet, SpreadsheetAddress.Parse(address));

    public void AutoFitColumns()
    {
        _worksheet.AutoFitColumns();
    }
}
