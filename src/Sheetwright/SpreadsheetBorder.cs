using System.Drawing;

namespace Sheetwright;

public sealed class SpreadsheetBorder
{
    private readonly SpreadsheetRange _range;

    internal SpreadsheetBorder(SpreadsheetRange range)
    {
        _range = range;
    }

    public void BorderAround(SpreadsheetBorderStyle borderStyle, Color color)
    {
        _range.ApplyBorderAround(borderStyle, $"{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}");
    }
}
