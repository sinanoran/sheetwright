using System.Drawing;

namespace Sheetwright;

/// <summary>
/// The borders a conditional format draws around each cell its rule matches.
/// </summary>
/// <remarks>
/// A differential format applies cell by cell, so <see cref="BorderAround"/> here outlines every
/// matching cell — unlike <see cref="SpreadsheetBorder.BorderAround"/>, which outlines a range.
/// </remarks>
public sealed class SpreadsheetConditionalFormattingBorderStyle
{
    private readonly SpreadsheetDifferentialFormat _format;

    internal SpreadsheetConditionalFormattingBorderStyle(SpreadsheetDifferentialFormat format)
    {
        _format = format;
    }

    public void BorderAround(SpreadsheetBorderStyle borderStyle, Color color)
    {
        SetLeft(borderStyle, color);
        SetRight(borderStyle, color);
        SetTop(borderStyle, color);
        SetBottom(borderStyle, color);
    }

    public void SetLeft(SpreadsheetBorderStyle borderStyle, Color color)
    {
        _format.LeftBorder = borderStyle;
        _format.LeftBorderColor = ToRgb(color);
    }

    public void SetRight(SpreadsheetBorderStyle borderStyle, Color color)
    {
        _format.RightBorder = borderStyle;
        _format.RightBorderColor = ToRgb(color);
    }

    public void SetTop(SpreadsheetBorderStyle borderStyle, Color color)
    {
        _format.TopBorder = borderStyle;
        _format.TopBorderColor = ToRgb(color);
    }

    public void SetBottom(SpreadsheetBorderStyle borderStyle, Color color)
    {
        _format.BottomBorder = borderStyle;
        _format.BottomBorderColor = ToRgb(color);
    }

    private static string ToRgb(Color color)
    {
        return $"{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
    }
}
