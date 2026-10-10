namespace Sheetwright;

/// <summary>
/// The formatting a conditional format applies to the cells its rule matches.
/// </summary>
/// <remarks>
/// Only what is set here is applied; anything left alone stays as the cell's own style has it.
/// A rule with nothing set is still written — with Stop If True, that is how Excel itself keeps
/// lower-priority rules off a set of cells — but it changes nothing on its own.
/// </remarks>
public sealed class SpreadsheetConditionalFormattingStyle
{
    private SpreadsheetConditionalFormattingFontStyle? _font;
    private SpreadsheetConditionalFormattingFillStyle? _fill;
    private SpreadsheetConditionalFormattingBorderStyle? _border;
    private SpreadsheetConditionalFormattingNumberFormatStyle? _numberFormat;

    internal SpreadsheetConditionalFormattingStyle()
    {
    }

    public SpreadsheetConditionalFormattingFontStyle Font => _font ??= new SpreadsheetConditionalFormattingFontStyle(Format);

    public SpreadsheetConditionalFormattingFillStyle Fill => _fill ??= new SpreadsheetConditionalFormattingFillStyle(Format);

    public SpreadsheetConditionalFormattingBorderStyle Border => _border ??= new SpreadsheetConditionalFormattingBorderStyle(Format);

    public SpreadsheetConditionalFormattingNumberFormatStyle NumberFormat =>
        _numberFormat ??= new SpreadsheetConditionalFormattingNumberFormatStyle(Format);

    internal SpreadsheetDifferentialFormat Format { get; } = new();
}
