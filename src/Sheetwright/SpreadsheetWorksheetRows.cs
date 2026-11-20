using System.Globalization;
using System.Text;
using System.Xml;

namespace Sheetwright;

/// <summary>
/// The rows of one worksheet part, read forward one at a time.
/// </summary>
/// <remarks>
/// <para>
/// The worksheet is never loaded: the XML reader is advanced to <c>sheetData</c> and each <c>row</c> is read,
/// turned into cells and handed on before the next is touched, so a sheet of a million rows costs the memory of
/// one.
/// </para>
/// <para>
/// What a cell's value is follows its type attribute: <c>s</c> an index into the shared strings, <c>inlineStr</c>
/// text written in the cell, <c>str</c> the text a formula calculated, <c>b</c> a flag, <c>e</c> an error,
/// <c>d</c> an ISO date, and a number otherwise (a date, when its style says it is a date). A formula is not
/// evaluated: the cell holds the value Excel stored beside it, and a formula with none is empty.
/// </para>
/// </remarks>
internal sealed class SpreadsheetWorksheetRows
{
    // Excel's last row. A row number past it is not a spreadsheet's.
    private const int LastExcelRow = 1_048_576;

    private readonly XmlReader _reader;
    private readonly string _partName;
    private readonly SpreadsheetReaderOptions _options;
    private readonly string[] _sharedStrings;
    private readonly bool[] _dateStyles;
    private readonly bool _date1904;
    private static readonly string[] IsoDateFormats =
    [
        "yyyy-MM-dd",
        "yyyy-MM-ddTHH:mm",
        "yyyy-MM-ddTHH:mm:ss",
        "yyyy-MM-ddTHH:mm:ss.FFFFFFF",
        "yyyy-MM-ddTHH:mmK",
        "yyyy-MM-ddTHH:mm:ssK",
        "yyyy-MM-ddTHH:mm:ss.FFFFFFFK",
    ];

    private readonly StringBuilder _value = new();
    private readonly StringBuilder _inline = new();
    private readonly List<(int Column, SpreadsheetCellValue Value)> _cells = [];

    internal SpreadsheetWorksheetRows(
        XmlReader reader,
        string partName,
        SpreadsheetReaderOptions options,
        string[] sharedStrings,
        bool[] dateStyles,
        bool date1904)
    {
        _reader = reader;
        _partName = partName;
        _options = options;
        _sharedStrings = sharedStrings;
        _dateStyles = dateStyles;
        _date1904 = date1904;
    }

    internal IEnumerable<SpreadsheetRow> Read()
    {
        {
            bool inSheetData = false;
            int sheetDataDepth = -1;
            int previousRow = 0;
            int returned = 0;

            while (SpreadsheetXml.Read(_reader, _partName))
            {
                if (_reader.NodeType == XmlNodeType.Element)
                {
                    if (!inSheetData)
                    {
                        if (_reader.LocalName == "sheetData")
                        {
                            if (_reader.IsEmptyElement)
                            {
                                yield break;
                            }

                            inSheetData = true;
                            sheetDataDepth = _reader.Depth;
                        }

                        continue;
                    }

                    if (_reader.LocalName != "row" || _reader.Depth != sheetDataDepth + 1)
                    {
                        continue;
                    }

                    int rowNumber = ReadRowNumber(previousRow);
                    previousRow = rowNumber;

                    SpreadsheetRow cells = ReadRow(rowNumber);
                    if (cells.IsEmpty && !_options.IncludeEmptyRows)
                    {
                        continue;
                    }

                    if (returned >= _options.MaxRows)
                    {
                        throw new SpreadsheetReadLimitException(
                            SpreadsheetReadLimit.MaxRows,
                            _options.MaxRows,
                            $"the sheet has more rows than that (row {rowNumber} is the first one past it).");
                    }

                    returned++;
                    yield return cells;
                }
                else if (_reader.NodeType == XmlNodeType.EndElement && inSheetData && _reader.Depth == sheetDataDepth)
                {
                    yield break;
                }
            }
        }
    }

    private int ReadRowNumber(int previousRow)
    {
        string? declared = _reader.GetAttribute("r");
        int rowNumber = previousRow + 1;
        if (declared is not null && !SpreadsheetXml.TryParseInt(declared, out rowNumber))
        {
            throw new InvalidDataException($"Part '{_partName}' has a row with a row number that is not a number.");
        }

        if (rowNumber <= previousRow)
        {
            throw new InvalidDataException($"Part '{_partName}' has rows that are not in order (row {rowNumber} after row {previousRow}).");
        }

        if (rowNumber > LastExcelRow)
        {
            throw new InvalidDataException($"Part '{_partName}' has a row numbered past the last row a spreadsheet has.");
        }

        return rowNumber;
    }

    /// <summary>
    /// The cells of the row the reader is on that have a value, as a row that ends at the last of them.
    /// </summary>
    private SpreadsheetRow ReadRow(int rowNumber)
    {
        _cells.Clear();

        if (!_reader.IsEmptyElement)
        {
            int rowDepth = _reader.Depth;
            int previousColumn = 0;

            while (SpreadsheetXml.Read(_reader, _partName))
            {
                if (_reader.NodeType == XmlNodeType.Element && _reader.LocalName == "c" && _reader.Depth == rowDepth + 1)
                {
                    int column = ReadColumn(previousColumn, rowNumber);
                    previousColumn = column;

                    SpreadsheetCellValue value = ReadCell(rowNumber);
                    if (!value.IsEmpty)
                    {
                        _cells.Add((column, value));
                    }
                }
                else if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == rowDepth)
                {
                    break;
                }
            }
        }

        int[] columns = new int[_cells.Count];
        SpreadsheetCellValue[] values = new SpreadsheetCellValue[_cells.Count];
        for (int index = 0; index < columns.Length; index++)
        {
            columns[index] = _cells[index].Column - 1;
            values[index] = _cells[index].Value;
        }

        return new SpreadsheetRow(rowNumber, columns, values);
    }

    private int ReadColumn(int previousColumn, int rowNumber)
    {
        string? reference = _reader.GetAttribute("r");
        int column = previousColumn + 1;

        if (reference is not null)
        {
            column = 0;
            int index = 0;
            while (index < reference.Length && char.IsAsciiLetter(reference[index]))
            {
                column = (column * 26) + (char.ToUpperInvariant(reference[index]) - 'A' + 1);
                if (column > _options.MaxColumns)
                {
                    throw ColumnLimit(rowNumber);
                }

                index++;
            }

            if (column == 0)
            {
                throw new InvalidDataException($"Part '{_partName}' has a cell in row {rowNumber.ToString(CultureInfo.InvariantCulture)} with no column in its reference.");
            }
        }

        if (column > _options.MaxColumns)
        {
            throw ColumnLimit(rowNumber);
        }

        if (column <= previousColumn)
        {
            throw new InvalidDataException(
                $"Part '{_partName}' has cells that are not in order in row {rowNumber.ToString(CultureInfo.InvariantCulture)}.");
        }

        return column;
    }

    private SpreadsheetReadLimitException ColumnLimit(int rowNumber)
    {
        return new SpreadsheetReadLimitException(
            SpreadsheetReadLimit.MaxColumns,
            _options.MaxColumns,
            $"row {rowNumber.ToString(CultureInfo.InvariantCulture)} has a cell past that column.");
    }

    private SpreadsheetCellValue ReadCell(int rowNumber)
    {
        string? type = _reader.GetAttribute("t");
        int style = SpreadsheetXml.TryParseInt(_reader.GetAttribute("s"), out int parsedStyle) ? parsedStyle : 0;

        _value.Clear();
        _inline.Clear();
        bool hasValue = false;
        bool hasInline = false;

        if (!_reader.IsEmptyElement)
        {
            int cellDepth = _reader.Depth;

            while (SpreadsheetXml.Read(_reader, _partName))
            {
                if (_reader.NodeType == XmlNodeType.Element && _reader.Depth == cellDepth + 1)
                {
                    if (_reader.LocalName == "v")
                    {
                        hasValue = true;
                        SpreadsheetXml.ReadText(_reader, _partName, _value, _options, "a cell", rowNumber);
                    }
                    else if (_reader.LocalName == "is")
                    {
                        hasInline = true;
                        SpreadsheetXml.ReadStringItem(_reader, _partName, _inline, _options, "a cell", rowNumber);
                    }
                }
                else if (_reader.NodeType == XmlNodeType.EndElement && _reader.Depth == cellDepth)
                {
                    break;
                }
            }
        }

        return Interpret(type, style, hasValue, hasInline);
    }

    private SpreadsheetCellValue Interpret(string? type, int style, bool hasValue, bool hasInline)
    {
        switch (type)
        {
            case "inlineStr":
                if (hasInline)
                {
                    return Text(_inline);
                }

                return hasValue ? Text(_value) : default;

            case "s":
                if (!hasValue)
                {
                    return default;
                }

                if (!SpreadsheetXml.TryParseInt(_value.ToString().Trim(), out int index) || index >= _sharedStrings.Length)
                {
                    throw new InvalidDataException($"Part '{_partName}' has a cell that names a shared string that does not exist.");
                }

                return _sharedStrings[index].Length == 0 ? default : SpreadsheetCellValue.FromText(_sharedStrings[index]);

            case "str":
                return hasValue ? Text(_value) : default;

            case "b":
                return hasValue ? SpreadsheetCellValue.FromBoolean(SpreadsheetXml.Flag(_value.ToString().Trim())) : default;

            case "e":
                return hasValue ? SpreadsheetCellValue.FromError(_value.ToString()) : default;

            case "d":
                if (!hasValue)
                {
                    return default;
                }

                string iso = _value.ToString().Trim();
                return DateTime.TryParseExact(iso, IsoDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime parsed)
                    ? SpreadsheetCellValue.FromDate(parsed)
                    : Text(iso);

            default:
                return Number(style, hasValue);
        }
    }

    private SpreadsheetCellValue Number(int style, bool hasValue)
    {
        if (!hasValue)
        {
            return default;
        }

        string raw = _value.ToString().Trim();
        if (raw.Length == 0)
        {
            return default;
        }

        if (!decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal number)
            || (number == 0m && HasNonZeroDigit(raw)))
        {
            // NaN, an infinity, a magnitude a decimal cannot hold in either direction (1E-30 parses, as 0).
            // Handing back text keeps what the file said.
            return SpreadsheetCellValue.FromText(raw);
        }

        if (style >= 0
            && style < _dateStyles.Length
            && _dateStyles[style]
            && SpreadsheetDateFormats.TryConvert(number, _date1904, out DateTime date))
        {
            return SpreadsheetCellValue.FromDate(date, number);
        }

        return SpreadsheetCellValue.FromNumber(number);
    }

    private static SpreadsheetCellValue Text(string text)
    {
        return text.Length == 0 ? default : SpreadsheetCellValue.FromText(text);
    }

    private static SpreadsheetCellValue Text(StringBuilder text)
    {
        return text.Length == 0 ? default : SpreadsheetCellValue.FromText(text.ToString());
    }

    // A number that parsed as zero has a digit other than zero in its mantissa only if it underflowed.
    private static bool HasNonZeroDigit(string raw)
    {
        foreach (char character in raw)
        {
            if (character is 'e' or 'E')
            {
                return false;
            }

            if (character is >= '1' and <= '9')
            {
                return true;
            }
        }

        return false;
    }
}
