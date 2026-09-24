namespace Sheetwright;

public sealed class SpreadsheetDimension
{
    internal SpreadsheetDimension(int fromRow, int fromColumn, int toRow, int toColumn)
    {
        Start = new SpreadsheetRangePosition(fromRow, fromColumn);
        End = new SpreadsheetRangePosition(toRow, toColumn);
        Reference = new SpreadsheetAddress(fromRow, fromColumn, toRow, toColumn).Reference;
    }

    public SpreadsheetRangePosition Start { get; }

    public SpreadsheetRangePosition End { get; }

    public string Reference { get; }

    public int Rows => End.Row - Start.Row + 1;

    public int Columns => End.Column - Start.Column + 1;
}
