namespace Sheetwright;

public sealed class SpreadsheetNumberFormatStyle
{
    private readonly SpreadsheetRange _range;

    internal SpreadsheetNumberFormatStyle(SpreadsheetRange range)
    {
        _range = range;
    }

    public void SetFormat(string format)
    {
        if (format is null)
        {
            throw new ArgumentNullException(nameof(format));
        }

        _range.ApplyStyle(new SpreadsheetStylePatch { NumberFormat = SpreadsheetTextSanitizer.Sanitize(format) });
    }
}
