namespace Sheetwright;

internal sealed class SpreadsheetTableDefinition
{
    internal SpreadsheetTableDefinition(string reference, string name, string styleName)
    {
        Reference = reference;
        Name = name;
        StyleName = styleName;
    }

    internal string Reference { get; }

    internal string Name { get; }

    internal string StyleName { get; set; }
}
