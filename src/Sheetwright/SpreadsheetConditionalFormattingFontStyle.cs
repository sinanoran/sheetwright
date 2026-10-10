namespace Sheetwright;

/// <summary>
/// The font changes a conditional format applies.
/// </summary>
/// <remarks>
/// Each setter is a change rather than a description, so <c>SetBold(false)</c> is not the same as
/// never calling it: it takes the bold off a cell whose own style is bold.
/// </remarks>
public sealed class SpreadsheetConditionalFormattingFontStyle
{
    private readonly SpreadsheetDifferentialFormat _format;

    internal SpreadsheetConditionalFormattingFontStyle(SpreadsheetDifferentialFormat format)
    {
        _format = format;
        Color = new SpreadsheetColor(value => _format.FontColor = value);
    }

    public SpreadsheetColor Color { get; }

    public void SetBold(bool bold)
    {
        _format.Bold = bold;
    }

    public void SetItalic(bool italic)
    {
        _format.Italic = italic;
    }

    public void SetUnderline(bool underline)
    {
        _format.Underline = underline;
    }

    public void SetStrikethrough(bool strikethrough)
    {
        _format.Strikethrough = strikethrough;
    }
}
