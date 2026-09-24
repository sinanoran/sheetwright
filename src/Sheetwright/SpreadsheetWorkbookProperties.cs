namespace Sheetwright;

public sealed class SpreadsheetWorkbookProperties
{
    private readonly SpreadsheetPackage _package;

    internal SpreadsheetWorkbookProperties(SpreadsheetPackage package)
    {
        _package = package;
    }

    public string Author
    {
        get => _package.GetAuthor();
        set => _package.SetAuthor(value ?? throw new ArgumentNullException(nameof(value)));
    }
}
