namespace Sheetwright;

public sealed class SpreadsheetRangePosition
{
    internal SpreadsheetRangePosition(int row, int column)
    {
        Row = row;
        Column = column;
    }

    public int Row { get; }

    public int Column { get; }
}
