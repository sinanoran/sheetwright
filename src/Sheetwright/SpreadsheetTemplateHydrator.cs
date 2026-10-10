using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Sheetwright;

/// <summary>
/// Hydrates existing Excel workbook templates with data while preserving existing
/// styles, formulas, charts, conditional formatting, and workbook structure.
/// </summary>
public sealed class SpreadsheetTemplateHydrator : IDisposable
{
    private static readonly Regex DefinedNameReferenceRegex = new(
        @"^(?:'?(?<sheet>[^'!]+)'?!)?\$?(?<col>[A-Za-z]+)\$?(?<row>\d+)",
        RegexOptions.Compiled);

    private readonly MemoryStream _stream;
    private readonly SpreadsheetDocument _document;
    private readonly WorkbookPart _workbookPart;
    private readonly Workbook _openXmlWorkbook;
    private readonly Dictionary<string, int> _sharedStrings = new(StringComparer.Ordinal);
    private SharedStringTablePart? _sharedStringPart;
    private bool _disposed;

    public SpreadsheetTemplateHydrator(Stream templateStream)
    {
        if (templateStream is null)
        {
            throw new ArgumentNullException(nameof(templateStream));
        }

        _stream = new MemoryStream();
        if (templateStream.CanSeek)
        {
            templateStream.Position = 0;
        }
        templateStream.CopyTo(_stream);
        _stream.Position = 0;

        _document = SpreadsheetDocument.Open(_stream, isEditable: true);
        _workbookPart = _document.WorkbookPart
            ?? throw new InvalidDataException("The template workbook is missing a WorkbookPart.");
        _openXmlWorkbook = _workbookPart.Workbook
            ?? throw new InvalidDataException("The template workbook is missing its Workbook.");

        IndexSharedStrings();
    }

    public SpreadsheetTemplateHydrator(byte[] templateBytes)
        : this(new MemoryStream(templateBytes ?? throw new ArgumentNullException(nameof(templateBytes))))
    {
    }

    public void SetCellValue(string cellAddress, object? value)
    {
        if (cellAddress is null)
        {
            throw new ArgumentNullException(nameof(cellAddress));
        }

        string sheetName;
        string reference;

        int exclamationIndex = cellAddress.IndexOf('!');
        if (exclamationIndex >= 0)
        {
            sheetName = cellAddress[..exclamationIndex].Trim('\'');
            reference = cellAddress[(exclamationIndex + 1)..];
        }
        else
        {
            sheetName = GetDefaultSheetName();
            reference = cellAddress;
        }

        SetCellValue(sheetName, reference, value);
    }

    public void SetCellValue(string sheetName, string cellAddress, object? value)
    {
        EnsureNotDisposed();

        if (sheetName is null)
        {
            throw new ArgumentNullException(nameof(sheetName));
        }

        if (cellAddress is null)
        {
            throw new ArgumentNullException(nameof(cellAddress));
        }

        WorksheetPart worksheetPart = GetWorksheetPart(sheetName);
        Worksheet worksheet = worksheetPart.Worksheet
            ?? throw new InvalidDataException($"Worksheet '{sheetName}' is missing content.");

        SheetData sheetData = worksheet.GetFirstChild<SheetData>()
            ?? worksheet.AppendChild(new SheetData());

        SpreadsheetAddress address = SpreadsheetAddress.Parse(cellAddress);
        Cell cell = GetOrCreateCell(sheetData, address);

        ApplyCellValue(cell, value);
    }

    public void SetNamedCellValue(string definedName, object? value)
    {
        EnsureNotDisposed();

        if (definedName is null)
        {
            throw new ArgumentNullException(nameof(definedName));
        }

        DefinedName? match = _openXmlWorkbook.DefinedNames?.Elements<DefinedName>()
            .FirstOrDefault(d => string.Equals(d.Name?.Value, definedName, StringComparison.OrdinalIgnoreCase));

        if (match is null || string.IsNullOrWhiteSpace(match.InnerText))
        {
            throw new KeyNotFoundException($"Defined name '{definedName}' was not found in the workbook template.");
        }

        Match regexMatch = DefinedNameReferenceRegex.Match(match.InnerText);
        if (!regexMatch.Success)
        {
            throw new InvalidOperationException(
                $"Defined name '{definedName}' target formula '{match.InnerText}' could not be parsed to a single cell reference.");
        }

        string sheetName = regexMatch.Groups["sheet"].Success
            ? regexMatch.Groups["sheet"].Value
            : GetDefaultSheetName();

        string reference = $"{regexMatch.Groups["col"].Value}{regexMatch.Groups["row"].Value}";
        SetCellValue(sheetName, reference, value);
    }

    public byte[] ToByteArray()
    {
        EnsureNotDisposed();

        if (_sharedStringPart?.SharedStringTable is not null)
        {
            _sharedStringPart.SharedStringTable.Count = null;
            _sharedStringPart.SharedStringTable.UniqueCount = (uint)_sharedStrings.Count;
            _sharedStringPart.SharedStringTable.Save();
        }

        _openXmlWorkbook.Save();
        _document.Dispose();
        _disposed = true;

        return _stream.ToArray();
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _document.Dispose();
        _stream.Dispose();
        _disposed = true;
    }

    private void EnsureNotDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    private string GetDefaultSheetName()
    {
        Sheet? sheet = _openXmlWorkbook.Sheets?.Elements<Sheet>().FirstOrDefault();
        return sheet?.Name?.Value
            ?? throw new InvalidOperationException("The workbook contains no worksheets.");
    }

    private WorksheetPart GetWorksheetPart(string sheetName)
    {
        Sheet? sheet = _openXmlWorkbook.Sheets?.Elements<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));

        if (sheet is null || sheet.Id?.Value is null)
        {
            throw new KeyNotFoundException($"Worksheet '{sheetName}' was not found in the template.");
        }

        return (WorksheetPart)_workbookPart.GetPartById(sheet.Id.Value);
    }

    private static Cell GetOrCreateCell(SheetData sheetData, SpreadsheetAddress address)
    {
        uint rowIndex = (uint)address.FromRow;
        string cellReference = $"{SpreadsheetAddress.GetColumnName(address.FromColumn)}{address.FromRow}";

        Row? row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row is null)
        {
            row = new Row { RowIndex = rowIndex };

            Row? successor = sheetData.Elements<Row>()
                .FirstOrDefault(r => r.RowIndex?.Value > rowIndex);

            if (successor is not null)
            {
                sheetData.InsertBefore(row, successor);
            }
            else
            {
                sheetData.Append(row);
            }
        }

        Cell? cell = row.Elements<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));

        if (cell is null)
        {
            cell = new Cell { CellReference = cellReference };

            Cell? successor = row.Elements<Cell>()
                .FirstOrDefault(c => IsColumnAfter(c.CellReference?.Value, address.FromColumn));

            if (successor is not null)
            {
                row.InsertBefore(cell, successor);
            }
            else
            {
                row.Append(cell);
            }
        }

        return cell;
    }

    private static bool IsColumnAfter(string? otherRef, int targetColumn)
    {
        if (string.IsNullOrEmpty(otherRef))
        {
            return false;
        }

        SpreadsheetAddress parsed = SpreadsheetAddress.Parse(otherRef);
        return parsed.FromColumn > targetColumn;
    }

    private void ApplyCellValue(Cell cell, object? value)
    {
        cell.CellFormula = null;

        if (value is null || value == DBNull.Value)
        {
            cell.CellValue = null;
            cell.InlineString = null;
            cell.DataType = null;
            return;
        }

        switch (value)
        {
            case string text:
                cell.CellValue = new CellValue(GetOrAddSharedString(text).ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.SharedString;
                break;

            case char character:
                cell.CellValue = new CellValue(GetOrAddSharedString(character.ToString()).ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.SharedString;
                break;

            case DateTime dateTime:
                cell.CellValue = new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
                break;

            case bool boolean:
                cell.CellValue = new CellValue(boolean ? "1" : "0");
                cell.DataType = CellValues.Boolean;
                break;

            case byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal:
                cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                cell.DataType = CellValues.Number;
                break;

            default:
                string converted = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                cell.CellValue = new CellValue(GetOrAddSharedString(converted).ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.SharedString;
                break;
        }
    }

    private void IndexSharedStrings()
    {
        _sharedStringPart = _workbookPart.SharedStringTablePart;
        if (_sharedStringPart?.SharedStringTable is null)
        {
            return;
        }

        int index = 0;
        foreach (SharedStringItem item in _sharedStringPart.SharedStringTable.Elements<SharedStringItem>())
        {
            string text = item.InnerText;
            if (!_sharedStrings.ContainsKey(text))
            {
                _sharedStrings[text] = index;
            }
            index++;
        }
    }

    private int GetOrAddSharedString(string text)
    {
        if (_sharedStrings.TryGetValue(text, out int existingIndex))
        {
            return existingIndex;
        }

        if (_sharedStringPart is null)
        {
            _sharedStringPart = _workbookPart.AddNewPart<SharedStringTablePart>();
            _sharedStringPart.SharedStringTable = new SharedStringTable();
        }
        else if (_sharedStringPart.SharedStringTable is null)
        {
            _sharedStringPart.SharedStringTable = new SharedStringTable();
        }

        int newIndex = _sharedStrings.Count;
        _sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
        _sharedStrings[text] = newIndex;
        return newIndex;
    }
}
