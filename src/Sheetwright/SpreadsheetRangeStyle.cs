namespace Sheetwright;

public sealed class SpreadsheetRangeStyle
{
    private readonly SpreadsheetRange _range;
    private SpreadsheetFontStyle? _font;
    private SpreadsheetFillStyle? _fill;
    private SpreadsheetBorder? _border;
    private SpreadsheetNumberFormatStyle? _numberFormat;

    internal SpreadsheetRangeStyle(SpreadsheetRange range)
    {
        _range = range;
    }

    public SpreadsheetFontStyle Font => _font ??= new SpreadsheetFontStyle(_range);

    public SpreadsheetFillStyle Fill => _fill ??= new SpreadsheetFillStyle(_range);

    public SpreadsheetBorder Border => _border ??= new SpreadsheetBorder(_range);

    public SpreadsheetNumberFormatStyle NumberFormat => _numberFormat ??= new SpreadsheetNumberFormatStyle(_range);

    public void SetHorizontalAlignment(SpreadsheetHorizontalAlignment horizontalAlignment)
    {
        _range.ApplyStyle(new SpreadsheetStylePatch { HorizontalAlignment = horizontalAlignment });
    }

    public void SetWrapText(bool wrapText)
    {
        _range.ApplyStyle(new SpreadsheetStylePatch { WrapText = wrapText });
    }

    public void SetLocked(bool locked)
    {
        _range.ApplyStyle(new SpreadsheetStylePatch { Locked = locked });
    }
}
