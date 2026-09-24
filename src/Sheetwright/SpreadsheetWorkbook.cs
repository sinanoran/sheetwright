using DocumentFormat.OpenXml.Packaging;

namespace Sheetwright;

public sealed class SpreadsheetWorkbook
{
    internal SpreadsheetWorkbook(SpreadsheetPackage package, WorkbookPart workbookPart)
    {
        Worksheets = new SpreadsheetWorksheetCollection(package, workbookPart);
        Names = new SpreadsheetDefinedNameCollection(workbookPart);
        Properties = new SpreadsheetWorkbookProperties(package);
    }

    public SpreadsheetWorksheetCollection Worksheets { get; }

    public SpreadsheetDefinedNameCollection Names { get; }

    public SpreadsheetWorkbookProperties Properties { get; }
}
