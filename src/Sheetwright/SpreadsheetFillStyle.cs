namespace Sheetwright;

public sealed class SpreadsheetFillStyle
{
    internal SpreadsheetFillStyle(SpreadsheetRange range)
    {
        BackgroundColor = new SpreadsheetColor(value => range.ApplyStyle(new SpreadsheetStylePatch { FillColor = value }));
    }

    public SpreadsheetColor BackgroundColor { get; }
}
