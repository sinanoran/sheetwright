namespace Sheetwright;

public sealed class SpreadsheetWorksheetDefaultStyle
{
    private readonly SpreadsheetWorksheet _worksheet;
    private SpreadsheetDefaultFontStyle? _font;

    internal SpreadsheetWorksheetDefaultStyle(SpreadsheetWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public SpreadsheetDefaultFontStyle Font => _font ??= new SpreadsheetDefaultFontStyle(_worksheet);

    public void SetLocked(bool locked)
    {
        _worksheet.ApplyDefaultStyle(new SpreadsheetStylePatch { Locked = locked });
    }

    public void SetWrapText(bool wrapText)
    {
        _worksheet.ApplyDefaultStyle(new SpreadsheetStylePatch { WrapText = wrapText });
    }
}
