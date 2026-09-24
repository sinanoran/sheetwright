namespace Sheetwright;

public sealed class SpreadsheetDefaultFontStyle
{
    private readonly SpreadsheetWorksheet _worksheet;

    internal SpreadsheetDefaultFontStyle(SpreadsheetWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public void SetName(string name)
    {
        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        _worksheet.ApplyDefaultStyle(new SpreadsheetStylePatch { FontName = SpreadsheetTextSanitizer.Sanitize(name) });
    }

    public void SetSize(double size)
    {
        _worksheet.ApplyDefaultStyle(new SpreadsheetStylePatch { FontSize = size });
    }
}
