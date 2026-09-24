namespace Sheetwright;

public sealed class SpreadsheetTables
{
    private readonly SpreadsheetWorksheet _worksheet;
    private int _count;

    internal SpreadsheetTables(SpreadsheetWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public int Count => _count;

    public SpreadsheetTable Add(SpreadsheetRange range, string name)
    {
        if (range is null)
        {
            throw new ArgumentNullException(nameof(range));
        }

        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        _count++;
        SpreadsheetTableDefinition definition = _worksheet.AddTable(range, name, "TableStyleMedium2");
        return new SpreadsheetTable(definition);
    }
}
