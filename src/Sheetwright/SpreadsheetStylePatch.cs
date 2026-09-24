namespace Sheetwright;

internal sealed class SpreadsheetStylePatch
{
    internal string? FontName { get; set; }

    internal double? FontSize { get; set; }

    internal bool? Bold { get; set; }

    internal bool? Italic { get; set; }

    internal string? FontColor { get; set; }

    internal string? FillColor { get; set; }

    internal string? NumberFormat { get; set; }

    internal SpreadsheetHorizontalAlignment? HorizontalAlignment { get; set; }

    internal bool? WrapText { get; set; }

    internal bool? Locked { get; set; }

    internal SpreadsheetBorderStyle? LeftBorder { get; set; }

    internal SpreadsheetBorderStyle? RightBorder { get; set; }

    internal SpreadsheetBorderStyle? TopBorder { get; set; }

    internal SpreadsheetBorderStyle? BottomBorder { get; set; }

    internal string? BorderColor { get; set; }

    internal void Merge(SpreadsheetStylePatch patch)
    {
        FontName = patch.FontName ?? FontName;
        FontSize = patch.FontSize ?? FontSize;
        Bold = patch.Bold ?? Bold;
        Italic = patch.Italic ?? Italic;
        FontColor = patch.FontColor ?? FontColor;
        FillColor = patch.FillColor ?? FillColor;
        NumberFormat = patch.NumberFormat ?? NumberFormat;
        HorizontalAlignment = patch.HorizontalAlignment ?? HorizontalAlignment;
        WrapText = patch.WrapText ?? WrapText;
        Locked = patch.Locked ?? Locked;
        LeftBorder = patch.LeftBorder ?? LeftBorder;
        RightBorder = patch.RightBorder ?? RightBorder;
        TopBorder = patch.TopBorder ?? TopBorder;
        BottomBorder = patch.BottomBorder ?? BottomBorder;
        BorderColor = patch.BorderColor ?? BorderColor;
    }

    internal SpreadsheetStylePatch Clone()
    {
        SpreadsheetStylePatch clone = new();
        clone.Merge(this);
        return clone;
    }
}
