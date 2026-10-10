namespace Sheetwright;

public sealed class SpreadsheetConditionalFormattingNumberFormatStyle
{
    private readonly SpreadsheetDifferentialFormat _format;

    internal SpreadsheetConditionalFormattingNumberFormatStyle(SpreadsheetDifferentialFormat format)
    {
        _format = format;
    }

    /// <exception cref="ArgumentException">
    /// The format is empty. A cell format with no number format is simply the default, but a
    /// differential format has to name one, and an empty <c>formatCode</c> is not one.
    /// </exception>
    public void SetFormat(string format)
    {
        if (format is null)
        {
            throw new ArgumentNullException(nameof(format));
        }

        string sanitized = SpreadsheetTextSanitizer.Sanitize(format);
        if (sanitized.Trim().Length == 0)
        {
            throw new ArgumentException("A conditional number format cannot be empty.", nameof(format));
        }

        _format.NumberFormat = sanitized;
    }
}
