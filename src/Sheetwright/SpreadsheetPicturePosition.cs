namespace Sheetwright;

public sealed class SpreadsheetPicturePosition
{
    internal SpreadsheetPicturePosition(int row)
    {
        Row = row;
    }

    public int Row { get; }
}
