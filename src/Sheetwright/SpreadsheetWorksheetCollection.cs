using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace Sheetwright;

public sealed class SpreadsheetWorksheetCollection : IEnumerable<SpreadsheetWorksheet>
{
    private readonly SpreadsheetPackage _package;
    private readonly WorkbookPart _workbookPart;
    private readonly Workbook _workbook;
    private readonly List<SpreadsheetWorksheet> _worksheets = new();

    internal SpreadsheetWorksheetCollection(SpreadsheetPackage package, WorkbookPart workbookPart)
    {
        _package = package;
        _workbookPart = workbookPart;
        _workbook = workbookPart.Workbook ?? throw new InvalidDataException("The workbook is missing.");
        Sheets sheets = GetSheets();
        foreach (Sheet sheet in sheets.Elements<Sheet>())
        {
            if (sheet.Id?.Value is string id && workbookPart.GetPartById(id) is WorksheetPart part)
            {
                _worksheets.Add(new SpreadsheetWorksheet(package, part, sheet));
            }
        }
    }

    public int Count => _worksheets.Count;

    public SpreadsheetWorksheet this[int oneBasedIndex] => _worksheets[oneBasedIndex - 1];

    public SpreadsheetWorksheet? this[string name] => name is null
        ? null
        : _worksheets.FirstOrDefault(
            worksheet => string.Equals(
                worksheet.Name,
                SpreadsheetTextSanitizer.Sanitize(name),
                StringComparison.OrdinalIgnoreCase));

    public SpreadsheetWorksheet Add(string name)
    {
        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        WorksheetPart worksheetPart = _workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        Sheets sheets = GetSheets();
        uint sheetId = sheets.Elements<Sheet>().Select(x => x.SheetId?.Value ?? 0U).DefaultIfEmpty().Max() + 1U;
        Sheet sheet = new()
        {
            Id = _workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = SpreadsheetTextSanitizer.Sanitize(name)
        };
        sheets.Append(sheet);
        SpreadsheetWorksheet worksheet = new(_package, worksheetPart, sheet);
        _worksheets.Add(worksheet);
        return worksheet;
    }

    public IEnumerator<SpreadsheetWorksheet> GetEnumerator()
    {
        return _worksheets.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private Sheets GetSheets()
    {
        if (_workbook.Sheets is Sheets sheets)
        {
            return sheets;
        }

        Sheets newSheets = new();
        _workbook.Append(newSheets);
        return newSheets;
    }
}
