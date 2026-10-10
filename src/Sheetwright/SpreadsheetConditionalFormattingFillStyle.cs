namespace Sheetwright;

public sealed class SpreadsheetConditionalFormattingFillStyle
{
    internal SpreadsheetConditionalFormattingFillStyle(SpreadsheetDifferentialFormat format)
    {
        BackgroundColor = new SpreadsheetColor(value => format.FillColor = value);
    }

    public SpreadsheetColor BackgroundColor { get; }
}
