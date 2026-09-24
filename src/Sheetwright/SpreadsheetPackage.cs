using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace Sheetwright;

public sealed class SpreadsheetPackage : IDisposable
{
    private readonly MemoryStream _stream;
    private readonly SpreadsheetDocument _document;
    private readonly System.IO.Packaging.Package? _customPackage;
    private readonly WorkbookPart _workbookPart;
    private readonly Workbook _openXmlWorkbook;
    private readonly SharedStringTablePart? _sharedStringPart;
    private readonly Dictionary<string, int> _sharedStrings = new(StringComparer.Ordinal);
    private readonly SpreadsheetStyleRepository? _styles;
    private readonly bool _editable;
    private string[]? _sharedStringIndex;
    private bool _finalized;
    private uint _tableId;

    public SpreadsheetPackage() : this(macroEnabled: false)
    {
    }

    public SpreadsheetPackage(bool macroEnabled)
    {
        _stream = new MemoryStream();
        _customPackage = new SpreadsheetMemoryPackage(_stream);
        _document = SpreadsheetDocument.Create(
            _customPackage,
            macroEnabled ? SpreadsheetDocumentType.MacroEnabledWorkbook : SpreadsheetDocumentType.Workbook,
            autoSave: true);
        _editable = true;
        _workbookPart = _document.AddWorkbookPart();
        _openXmlWorkbook = new Workbook(new Sheets());
        _workbookPart.Workbook = _openXmlWorkbook;
        _sharedStringPart = _workbookPart.AddNewPart<SharedStringTablePart>();
        _sharedStringPart.SharedStringTable = new SharedStringTable();
        _styles = new SpreadsheetStyleRepository(_workbookPart);
        Workbook = new SpreadsheetWorkbook(this, _workbookPart);
    }

    public SpreadsheetPackage(Stream source)
    {
        if (source is null)
        {
            throw new ArgumentNullException(nameof(source));
        }

        _stream = new MemoryStream();
        if (source.CanSeek)
        {
            source.Position = 0;
        }
        source.CopyTo(_stream);
        _stream.Position = 0;
        _document = SpreadsheetDocument.Open(_stream, false);
        _editable = false;
        _workbookPart = _document.WorkbookPart ?? throw new InvalidDataException("The workbook part is missing.");
        _openXmlWorkbook = _workbookPart.Workbook ?? throw new InvalidDataException("The workbook is missing.");
        _sharedStringPart = _workbookPart.SharedStringTablePart;
        Workbook = new SpreadsheetWorkbook(this, _workbookPart);
    }

    public SpreadsheetWorkbook Workbook { get; }

    public byte[] GetAsByteArray()
    {
        FinalizePackage();
        return _stream.ToArray();
    }

    public void SaveAs(Stream destination)
    {
        if (destination is null)
        {
            throw new ArgumentNullException(nameof(destination));
        }

        FinalizePackage();
        _stream.Position = 0;
        _stream.CopyTo(destination);
    }

    public void SetVbaProject(Stream vbaProject)
    {
        if (vbaProject is null)
        {
            throw new ArgumentNullException(nameof(vbaProject));
        }

        EnsureEditable();
        if (_document.DocumentType != SpreadsheetDocumentType.MacroEnabledWorkbook)
        {
            throw new InvalidOperationException("The package must be created as macro-enabled before attaching a VBA project.");
        }
        VbaProjectPart part = _workbookPart.VbaProjectPart ?? _workbookPart.AddNewPart<VbaProjectPart>();
        if (vbaProject.CanSeek)
        {
            vbaProject.Position = 0;
        }
        part.FeedData(vbaProject);

        WorkbookProperties properties = _openXmlWorkbook.WorkbookProperties
            ?? _openXmlWorkbook.InsertAt(new WorkbookProperties(), 0);
        properties.CodeName = "ThisWorkbook";
        foreach (SpreadsheetWorksheet worksheet in Workbook.Worksheets)
        {
            worksheet.SetCodeName(worksheet.Name);
        }
    }

    internal int GetSharedStringIndex(string text)
    {
        EnsureEditable();
        text = SpreadsheetTextSanitizer.Sanitize(text);
        if (_sharedStrings.TryGetValue(text, out int existing))
        {
            return existing;
        }

        int index = _sharedStrings.Count;
        _sharedStringPart!.SharedStringTable!.AppendChild(new SharedStringItem(new Text(text)
        {
            Space = text.Length != text.Trim().Length ? DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve : null
        }));
        _sharedStrings[text] = index;
        _sharedStringIndex = null;
        return index;
    }

    internal object? ReadCellValue(Cell cell)
    {
        string? raw = cell.CellValue?.InnerText;
        if (raw is null && cell.InlineString?.InnerText is string inline)
        {
            return inline;
        }

        if (raw is null)
        {
            return null;
        }

        CellValues? dataType = cell.DataType?.Value;
        if (dataType == CellValues.SharedString)
        {
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)
                ? ReadSharedString(index)
                : raw;
        }

        if (dataType == CellValues.Boolean)
        {
            return raw == "1" || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);
        }

        if (dataType == CellValues.Date && DateTime.TryParse(
                raw,
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out DateTime parsedDate))
        {
            return parsedDate;
        }

        if (dataType == CellValues.String || dataType == CellValues.InlineString || dataType == CellValues.Error)
        {
            return raw;
        }

        if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
        {
            if (!IsDateCell(cell))
            {
                return number;
            }

            DateTime serialDate = DateTime.FromOADate(number);
            return _openXmlWorkbook.WorkbookProperties?.Date1904?.Value == true
                ? serialDate.AddDays(1462)
                : serialDate;
        }

        return raw;
    }

    internal uint NextTableId()
    {
        return ++_tableId;
    }

    internal void SetAuthor(string author)
    {
        _document.PackageProperties.Creator = SpreadsheetTextSanitizer.Sanitize(author);
    }

    internal string GetAuthor()
    {
        return _document.PackageProperties.Creator ?? string.Empty;
    }

    internal SpreadsheetStyleRepository Styles => _styles ?? throw new InvalidOperationException("A read-only workbook cannot be styled.");

    public void Dispose()
    {
        _document.Dispose();
        if (!_finalized)
        {
            _customPackage?.Close();
        }
        _stream.Dispose();
    }

    private void FinalizePackage()
    {
        if (_finalized)
        {
            return;
        }

        EnsureEditable();
        foreach (SpreadsheetWorksheet worksheet in Workbook.Worksheets)
        {
            worksheet.Commit(Styles);
        }
        Styles.Save();
        SharedStringTable sharedStringTable = _sharedStringPart?.SharedStringTable ?? throw new InvalidOperationException("The shared string table is missing.");
        sharedStringTable.Count = null;
        sharedStringTable.UniqueCount = (uint)_sharedStrings.Count;
        sharedStringTable.Save();
        _openXmlWorkbook.Save();
        _document.Dispose();
        if (_customPackage != null)
        {
            _customPackage.Flush();
            _customPackage.Close();
        }
        _finalized = true;
    }

    private bool IsDateCell(Cell cell)
    {
        if (cell.StyleIndex?.Value is not uint styleIndex)
        {
            return false;
        }

        Stylesheet? stylesheet = _workbookPart.WorkbookStylesPart?.Stylesheet;
        CellFormat? format = stylesheet?.CellFormats?.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
        uint numberFormatId = format?.NumberFormatId?.Value ?? 0;
        if ((numberFormatId >= 14 && numberFormatId <= 22) 
            || (numberFormatId >= 27 && numberFormatId <= 36)
            || (numberFormatId >= 45 && numberFormatId <= 47)
            || (numberFormatId >= 50 && numberFormatId <= 58))
        {
            return true;
        }

        string? customFormat = stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
            .FirstOrDefault(x => x.NumberFormatId?.Value == numberFormatId)?.FormatCode?.Value;
        if (customFormat is null || customFormat.Length == 0)
        {
            return false;
        }

        return IsDateFormatCode(customFormat);
    }

    private string? ReadSharedString(int index)
    {
        _sharedStringIndex ??= _sharedStringPart?.SharedStringTable?
            .Elements<SharedStringItem>()
            .Select(x => x.InnerText)
            .ToArray()
            ?? Array.Empty<string>();

        return index >= 0 && index < _sharedStringIndex.Length
            ? _sharedStringIndex[index]
            : null;
    }

    private static bool IsDateFormatCode(string formatCode)
    {
        bool inQuotedLiteral = false;

        for (int index = 0; index < formatCode.Length; index++)
        {
            char character = formatCode[index];
            if (character == '"')
            {
                inQuotedLiteral = !inQuotedLiteral;
                continue;
            }
            if (inQuotedLiteral)
            {
                continue;
            }
            if (character is '\\' or '_' or '*')
            {
                index++;
                continue;
            }
            if (character == '[')
            {
                int closingBracket = formatCode.IndexOf(']', index + 1);
                if (closingBracket < 0)
                {
                    break;
                }

                string bracketContent = formatCode.Substring(index + 1, closingBracket - index - 1).Trim();
                if (bracketContent.Length > 0 && bracketContent.All(value => value is 'h' or 'H' or 'm' or 'M' or 's' or 'S'))
                {
                    return true;
                }

                index = closingBracket;
                continue;
            }

            if (character is 'y' or 'Y' or 'm' or 'M' or 'd' or 'D' or 'h' or 'H' or 's' or 'S')
            {
                return true;
            }
        }

        return false;
    }

    private void EnsureEditable()
    {
        if (!_editable || _finalized)
        {
            throw new InvalidOperationException("The spreadsheet package is not editable.");
        }
    }
}
