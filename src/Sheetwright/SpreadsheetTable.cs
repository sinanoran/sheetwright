namespace Sheetwright;

public sealed class SpreadsheetTable
{
    private readonly SpreadsheetTableDefinition _definition;

    internal SpreadsheetTable(SpreadsheetTableDefinition definition)
    {
        _definition = definition;
    }

    public string TableStyle
    {
        get => _definition.StyleName;
        set => _definition.StyleName = SpreadsheetTextSanitizer.Sanitize(
            value ?? throw new ArgumentNullException(nameof(value)));
    }
}
