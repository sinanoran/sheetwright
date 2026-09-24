namespace Sheetwright;

public sealed class SpreadsheetFontStyle
{
    private readonly SpreadsheetRange _range;

    internal SpreadsheetFontStyle(SpreadsheetRange range)
    {
        _range = range;
        Color = new SpreadsheetColor(value => _range.ApplyStyle(new SpreadsheetStylePatch { FontColor = value }));
    }

    public SpreadsheetColor Color { get; }

    public void SetName(string name)
    {
        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        _range.ApplyStyle(new SpreadsheetStylePatch { FontName = SpreadsheetTextSanitizer.Sanitize(name) });
    }

    public void SetSize(double size)
    {
        _range.ApplyStyle(new SpreadsheetStylePatch { FontSize = size });
    }

    public void SetBold(bool bold)
    {
        _range.ApplyStyle(new SpreadsheetStylePatch { Bold = bold });
    }

    public void SetItalic(bool italic)
    {
        _range.ApplyStyle(new SpreadsheetStylePatch { Italic = italic });
    }
}
