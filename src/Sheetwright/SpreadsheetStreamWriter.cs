using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Sheetwright;

/// <summary>
/// Writes a single-worksheet workbook straight to a stream, one row at a time, so memory stays
/// flat however many rows are written.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="SpreadsheetPackage"/> holds every cell as an OpenXML element until it is asked for
/// the bytes, which is what makes the rest of this library possible — styles merge, columns
/// auto-fit, a table reads its headings back out of its own header row — and what makes an export
/// of unbounded size an unbounded allocation. A report of several hundred thousand rows runs a
/// worker out of memory. This writer is the other trade: nothing is kept but the current row, and
/// nothing can be revisited.
/// </para>
/// <para>
/// What that costs. Text is written as an inline string, so there is no shared string table to
/// hold open — a workbook of repeated values is larger than the same workbook from
/// <see cref="SpreadsheetPackage"/>. Columns and tables have to be declared before the first row,
/// because the column definitions and the content types are written ahead of the rows. Rows go in
/// in ascending order and cells in ascending column order within a row. There is one worksheet,
/// and no pictures, data validation, merged cells, freeze panes or auto-fit.
/// </para>
/// <para>
/// Call <see cref="Complete"/> to finish the workbook. Disposing without completing still leaves a
/// readable ZIP archive, but one with no workbook, styles or table parts: it is not a workbook and
/// should be discarded. The destination stream is left open either way, so the caller keeps
/// ownership of it.
/// </para>
/// <para>
/// Values are typed and converted exactly as <see cref="SpreadsheetPackage"/> types and converts
/// them, including the <c>dd/mm/yyyy</c> a <see cref="DateTime"/> gets when the caller gave no
/// format of their own, so the same rows produce the same cells whichever writer wrote them.
/// </para>
/// </remarks>
public sealed class SpreadsheetStreamWriter : IDisposable
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string CorePropertiesNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private const string DublinCoreNamespace = "http://purl.org/dc/elements/1.1/";
    private const string RelationshipTypeBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    private const string ContentTypeBase = "application/vnd.openxmlformats-officedocument.spreadsheetml.";

    /// <summary>The style <see cref="SpreadsheetTables"/> gives a table it is not told about.</summary>
    private const string DefaultTableStyle = "TableStyleMedium2";

    private readonly ZipArchive _archive;
    private readonly string _sheetName;
    private readonly string? _author;
    private readonly SortedDictionary<int, ColumnSettings> _columns = new();
    private readonly Dictionary<Style, uint> _styleIndexes = new();
    private readonly List<Style> _styles = new();
    private readonly List<TablePart> _tables = new();

    /// <summary>
    /// The resolved heading of every table header cell, by row and then column.
    /// </summary>
    /// <remarks>
    /// A heading invented for an empty cell or made unique against a duplicate has to end up in
    /// the header cell as well as in the table part, because Excel checks the two against each
    /// other and offers to repair the file when they disagree. The worksheet writer settles this
    /// after the fact; here the header row may be long gone by the time the table part is written,
    /// so the headings are resolved when the table is added and the header cells are written from
    /// them rather than from what the caller passes to <see cref="WriteCell"/>.
    /// </remarks>
    private readonly Dictionary<int, SortedDictionary<int, string>> _headerCells = new();

    /// <summary>
    /// The header rows of tables that have been added and not yet written.
    /// </summary>
    /// <remarks>
    /// A table whose header row is never started names columns that are not in the sheet, which is
    /// the same repair dialogue as a header row that disagrees with its table. It cannot be put
    /// right at <see cref="Complete"/> — the rows are gone — so it is refused there instead, where
    /// a caller still has a stack trace rather than a file somebody cannot open.
    /// </remarks>
    private readonly HashSet<int> _unwrittenHeaderRows = new();

    private Stream? _sheetStream;
    private XmlWriter? _sheetWriter;
    private SortedDictionary<int, string>? _rowHeadings;
    private int _currentRow;
    private int _currentColumn;
    private bool _rowOpen;
    private bool _completed;

    /// <param name="destination">
    /// Where the workbook is written. It is left open, so the caller disposes it.
    /// </param>
    /// <param name="sheetName">The name of the single worksheet.</param>
    /// <param name="author">The workbook's author, or <see langword="null"/> for none.</param>
    public SpreadsheetStreamWriter(Stream destination, string sheetName, string? author = null)
    {
        if (destination is null)
        {
            throw new ArgumentNullException(nameof(destination));
        }

        if (!destination.CanWrite)
        {
            throw new ArgumentException("A workbook cannot be written to a read-only stream.", nameof(destination));
        }

        if (sheetName is null)
        {
            throw new ArgumentNullException(nameof(sheetName));
        }

        _archive = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
        _sheetName = SpreadsheetTextSanitizer.Sanitize(sheetName);
        _author = SpreadsheetTextSanitizer.SanitizeNullable(author);

        // Index 0 is the default style, so a caller who never asks for one still has a valid
        // cellXfs entry to point at.
        GetStyleIndex();
    }

    /// <summary>
    /// Returns the style index for the given formatting, registering it on first use. Index 0 is
    /// the default style, and identical formatting always returns the same index.
    /// </summary>
    /// <remarks>
    /// This is the whole of the styling available here. The style table is written at
    /// <see cref="Complete"/>, so a style may be registered at any point before then.
    /// </remarks>
    public uint GetStyleIndex(bool bold = false, bool italic = false, string? numberFormat = null, bool wrapText = false)
    {
        EnsureNotCompleted();

        Style style = new(
            bold,
            italic,
            string.IsNullOrWhiteSpace(numberFormat) ? null : SpreadsheetTextSanitizer.Sanitize(numberFormat!),
            wrapText);

        if (_styleIndexes.TryGetValue(style, out uint existing))
        {
            return existing;
        }

        uint index = (uint)_styles.Count;
        _styles.Add(style);
        _styleIndexes[style] = index;
        return index;
    }

    /// <summary>
    /// Sets a column's width, visibility and default style. Columns must be set before the first
    /// row is written.
    /// </summary>
    /// <param name="column">The one-based column.</param>
    /// <param name="width">The width in characters, or <see langword="null"/> for Excel's default.</param>
    /// <param name="hidden">Whether the column is hidden.</param>
    /// <param name="styleIndex">A style from <see cref="GetStyleIndex"/>, applied to cells of this column that carry no style of their own.</param>
    public void SetColumn(int column, double? width, bool hidden = false, uint styleIndex = 0)
    {
        EnsureNotCompleted();

        if (_sheetWriter is not null)
        {
            throw new InvalidOperationException("Columns must be set before the first row is written.");
        }

        ThrowIfOutsideTheGrid(column, nameof(column), SpreadsheetAddress.MaximumColumn);
        ThrowIfUnknownStyle(styleIndex, nameof(styleIndex));

        if (width.HasValue && width.Value <= 0D)
        {
            throw new ArgumentOutOfRangeException(nameof(width), width, "A column width must be greater than zero.");
        }

        _columns[column] = new ColumnSettings(width, hidden, styleIndex);
    }

    /// <summary>
    /// Adds a table over the given range, whose first row is the header row. Tables must be added
    /// before the first row is written.
    /// </summary>
    /// <param name="fromRow">The one-based row of the header row.</param>
    /// <param name="fromColumn">The one-based first column of the range.</param>
    /// <param name="toRow">
    /// The one-based last row of the range, header row included, or <see langword="null"/> for the
    /// last row written — which is the usual answer when the rows are a query nobody has counted.
    /// The table part is written at <see cref="Complete"/>, so this one thing does not have to be
    /// known in advance.
    /// </param>
    /// <param name="toColumn">The one-based last column of the range.</param>
    /// <param name="name">The table's name, which Excel also shows as its display name.</param>
    /// <param name="headings">
    /// One heading per column of the range. An empty heading becomes its column letter and a
    /// repeated one gains a suffix, because Excel refuses duplicates; the resolved heading is what
    /// gets written into the header cell, whatever <see cref="WriteCell"/> is given for it.
    /// </param>
    /// <param name="styleName">The name of an Excel table style.</param>
    public void AddTable(
        int fromRow,
        int fromColumn,
        int? toRow,
        int toColumn,
        string name,
        IReadOnlyList<string> headings,
        string styleName = DefaultTableStyle)
    {
        EnsureNotCompleted();

        if (_sheetWriter is not null)
        {
            throw new InvalidOperationException("Tables must be added before the first row is written.");
        }

        if (name is null)
        {
            throw new ArgumentNullException(nameof(name));
        }

        if (styleName is null)
        {
            throw new ArgumentNullException(nameof(styleName));
        }

        if (headings is null)
        {
            throw new ArgumentNullException(nameof(headings));
        }

        // An address is built here, before anything is derived from the range, so an impossible
        // one is refused on the line that wrote it and with the message the rest of the library
        // uses. A range left open at the bottom is validated as the header row alone, which is
        // the smallest it can turn out to be.
        _ = new SpreadsheetAddress(fromRow, fromColumn, toRow ?? fromRow, toColumn);

        if (headings.Count != toColumn - fromColumn + 1)
        {
            throw new ArgumentException("A heading is required for every column of the table.", nameof(headings));
        }

        HashSet<string> taken = new(StringComparer.OrdinalIgnoreCase);
        List<string> resolved = new(headings.Count);

        if (!_headerCells.TryGetValue(fromRow, out SortedDictionary<int, string>? headerRow))
        {
            headerRow = new SortedDictionary<int, string>();
            _headerCells[fromRow] = headerRow;
        }

        _unwrittenHeaderRows.Add(fromRow);

        for (int index = 0; index < headings.Count; index++)
        {
            string heading = SpreadsheetTableHeadings.Resolve(
                SpreadsheetTextSanitizer.SanitizeNullable(headings[index]),
                fromColumn + index,
                taken);

            resolved.Add(heading);
            headerRow[fromColumn + index] = heading;
        }

        _tables.Add(new TablePart(
            fromRow,
            fromColumn,
            toRow,
            toColumn,
            SpreadsheetTextSanitizer.Sanitize(name),
            SpreadsheetTextSanitizer.Sanitize(styleName),
            resolved));
    }

    /// <summary>
    /// Starts a row, closing the previous one. Rows must be written in ascending order, and a row
    /// that is never started is simply absent from the sheet.
    /// </summary>
    public void StartRow(int rowNumber)
    {
        EnsureNotCompleted();
        ThrowIfOutsideTheGrid(rowNumber, nameof(rowNumber), SpreadsheetAddress.MaximumRow);

        if (rowNumber <= _currentRow)
        {
            throw new InvalidOperationException(
                $"Rows must be written in ascending order, and row {rowNumber.ToString(CultureInfo.InvariantCulture)} follows row {_currentRow.ToString(CultureInfo.InvariantCulture)}.");
        }

        EnsureSheetStarted();
        EndRow();

        _currentRow = rowNumber;
        _currentColumn = 0;
        _rowOpen = true;
        _headerCells.TryGetValue(rowNumber, out _rowHeadings);
        _unwrittenHeaderRows.Remove(rowNumber);

        _sheetWriter!.WriteStartElement("row", MainNamespace);
        _sheetWriter.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Writes a cell in the current row. <see langword="null"/> and <see cref="DBNull"/> leave the
    /// cell empty, and cells must be written in ascending column order.
    /// </summary>
    /// <param name="column">The one-based column.</param>
    /// <param name="value">The value, typed as <see cref="SpreadsheetPackage"/> types it.</param>
    /// <param name="styleIndex">A style from <see cref="GetStyleIndex"/>.</param>
    public void WriteCell(int column, object? value, uint styleIndex = 0)
    {
        EnsureNotCompleted();

        if (!_rowOpen)
        {
            throw new InvalidOperationException("A row must be started before a cell is written.");
        }

        ThrowIfOutsideTheGrid(column, nameof(column), SpreadsheetAddress.MaximumColumn);
        ThrowIfUnknownStyle(styleIndex, nameof(styleIndex));

        if (column <= _currentColumn)
        {
            throw new InvalidOperationException(
                $"Cells must be written in ascending column order, and column {column.ToString(CultureInfo.InvariantCulture)} follows column {_currentColumn.ToString(CultureInfo.InvariantCulture)}.");
        }

        WriteHeadingsBefore(column);

        if (_rowHeadings is not null && _rowHeadings.TryGetValue(column, out string? heading))
        {
            value = heading;
        }

        _currentColumn = column;
        WriteCellCore(column, value, styleIndex);
    }

    /// <summary>
    /// Finishes the workbook: the sheet is closed and the workbook, style, table and property
    /// parts are written. Nothing can be written afterwards.
    /// </summary>
    public void Complete()
    {
        EnsureNotCompleted();

        if (_unwrittenHeaderRows.Count > 0)
        {
            string rows = string.Join(
                ", ",
                _unwrittenHeaderRows.Order().Select(row => row.ToString(CultureInfo.InvariantCulture)));

            throw new InvalidOperationException(
                $"A table's header row was never written: row {rows}.");
        }

        EnsureSheetStarted();
        EndRow();

        XmlWriter sheetWriter = _sheetWriter!;
        sheetWriter.WriteEndElement(); // sheetData

        if (_tables.Count > 0)
        {
            sheetWriter.WriteStartElement("tableParts", MainNamespace);
            sheetWriter.WriteAttributeString("count", _tables.Count.ToString(CultureInfo.InvariantCulture));
            for (int index = 1; index <= _tables.Count; index++)
            {
                sheetWriter.WriteStartElement("tablePart", MainNamespace);
                sheetWriter.WriteAttributeString("id", RelationshipNamespace, RelationshipId(index));
                sheetWriter.WriteEndElement();
            }

            sheetWriter.WriteEndElement();
        }

        sheetWriter.WriteEndElement(); // worksheet
        sheetWriter.WriteEndDocument();

        // One ZIP entry at a time: the sheet has to be closed before the remaining parts are
        // written, and the writer has to be flushed before the entry is.
        sheetWriter.Dispose();
        _sheetWriter = null;
        _sheetStream!.Dispose();
        _sheetStream = null;

        WriteTables();
        WriteStyles();
        WriteWorkbook();
        WriteCoreProperties();

        _completed = true;
        _archive.Dispose();
    }

    public void Dispose()
    {
        _sheetWriter?.Dispose();
        _sheetWriter = null;
        _sheetStream?.Dispose();
        _sheetStream = null;
        _archive.Dispose();
    }

    private void EnsureSheetStarted()
    {
        if (_sheetWriter is not null)
        {
            return;
        }

        // Every part is known once the columns and tables are in, so the content types go in as
        // the first entry — which is what SpreadsheetPackageArchiveWriter does, and what Excel
        // itself writes.
        WriteContentTypes();

        _sheetStream = _archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Optimal).Open();
        _sheetWriter = CreateXmlWriter(_sheetStream);
        _sheetWriter.WriteStartDocument();
        _sheetWriter.WriteStartElement("worksheet", MainNamespace);
        _sheetWriter.WriteAttributeString("xmlns", "r", null, RelationshipNamespace);

        if (_columns.Count > 0)
        {
            _sheetWriter.WriteStartElement("cols", MainNamespace);
            foreach (KeyValuePair<int, ColumnSettings> column in _columns)
            {
                string index = column.Key.ToString(CultureInfo.InvariantCulture);
                _sheetWriter.WriteStartElement("col", MainNamespace);
                _sheetWriter.WriteAttributeString("min", index);
                _sheetWriter.WriteAttributeString("max", index);
                if (column.Value.Width is double width)
                {
                    _sheetWriter.WriteAttributeString("width", width.ToString("R", CultureInfo.InvariantCulture));
                    _sheetWriter.WriteAttributeString("customWidth", "1");
                }

                if (column.Value.StyleIndex != 0)
                {
                    _sheetWriter.WriteAttributeString("style", column.Value.StyleIndex.ToString(CultureInfo.InvariantCulture));
                }

                if (column.Value.Hidden)
                {
                    _sheetWriter.WriteAttributeString("hidden", "1");
                }

                _sheetWriter.WriteEndElement();
            }

            _sheetWriter.WriteEndElement();
        }

        _sheetWriter.WriteStartElement("sheetData", MainNamespace);
    }

    private void EndRow()
    {
        if (!_rowOpen)
        {
            return;
        }

        // Any header cell the caller skipped still has to be there: the table part names it.
        WriteHeadingsBefore(int.MaxValue);

        _sheetWriter!.WriteEndElement();
        _rowOpen = false;
        _rowHeadings = null;
    }

    /// <summary>
    /// Writes the header cells of the current row that sit before <paramref name="column"/> and
    /// that the caller has not written, so the header row can never be missing a cell the table
    /// part names. They get the default style; a caller who writes every heading itself — which
    /// is the usual shape — keeps whatever style it gave them.
    /// </summary>
    private void WriteHeadingsBefore(int column)
    {
        if (_rowHeadings is null)
        {
            return;
        }

        foreach (KeyValuePair<int, string> heading in _rowHeadings)
        {
            if (heading.Key <= _currentColumn)
            {
                continue;
            }

            if (heading.Key >= column)
            {
                break;
            }

            _currentColumn = heading.Key;
            WriteCellCore(heading.Key, heading.Value, 0);
        }
    }

    private void WriteCellCore(int column, object? value, uint styleIndex)
    {
        if (value is null || value == DBNull.Value)
        {
            return;
        }

        if (value is DateTime)
        {
            styleIndex = GetDateStyleIndex(column, styleIndex);
        }

        XmlWriter writer = _sheetWriter!;
        writer.WriteStartElement("c", MainNamespace);
        writer.WriteAttributeString("r", SpreadsheetAddress.GetCellReference(_currentRow, column));
        if (styleIndex != 0)
        {
            writer.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));
        }

        // Typed exactly as SpreadsheetWorksheet.SetValue types the same value, down to the
        // invariant conversion: a workbook is a file rather than a screen, and the same object
        // written on a Turkish machine and an American one has to produce the same bytes.
        switch (value)
        {
            case string text:
                WriteInlineString(writer, text);
                break;

            case char character:
                WriteInlineString(writer, character.ToString());
                break;

            case DateTime dateTime:
                WriteNumber(writer, dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                break;

            case bool boolean:
                writer.WriteAttributeString("t", "b");
                writer.WriteElementString("v", MainNamespace, boolean ? "1" : "0");
                break;

            // NaN and the infinities have no numeric representation in the file format. Written
            // as a number they produce a cell Excel offers to repair rather than open.
            case double number when double.IsNaN(number) || double.IsInfinity(number):
                WriteInlineString(writer, number.ToString(CultureInfo.InvariantCulture));
                break;

            case float number when float.IsNaN(number) || float.IsInfinity(number):
                WriteInlineString(writer, number.ToString(CultureInfo.InvariantCulture));
                break;

            case byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal:
                WriteNumber(writer, Convert.ToString(value, CultureInfo.InvariantCulture)!);
                break;

            default:
                WriteInlineString(writer, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                break;
        }

        writer.WriteEndElement();
    }

    /// <summary>
    /// The style a <see cref="DateTime"/> cell is written with: the caller's own, or the caller's
    /// own plus <c>dd/mm/yyyy</c> when nothing already formats the cell.
    /// </summary>
    /// <remarks>
    /// A date cell with no number format shows Excel's serial number and the export looks broken,
    /// so <see cref="SpreadsheetWorksheet"/> gives one a format and so does this. A format the
    /// caller set wins. The column's format counts only for a cell that has no style of its own,
    /// because that is the only case where Excel falls back to it: a cell with its own style is
    /// formatted by that style alone, so leaving it General would show the serial number.
    /// </remarks>
    private uint GetDateStyleIndex(int column, uint styleIndex)
    {
        Style style = _styles[(int)styleIndex];
        if (style.NumberFormat is not null)
        {
            return styleIndex;
        }

        if (styleIndex == 0
            && _columns.TryGetValue(column, out ColumnSettings settings)
            && _styles[(int)settings.StyleIndex].NumberFormat is not null)
        {
            return styleIndex;
        }

        return GetStyleIndex(
            style.Bold,
            style.Italic,
            SpreadsheetWorksheet.ImplicitDateNumberFormat,
            style.WrapText);
    }

    // Mirrors SpreadsheetWorksheet.CommitTables — the autoFilter, the column numbering from one
    // and the tableStyleInfo flags; keep the two in step.
    private void WriteTables()
    {
        if (_tables.Count == 0)
        {
            return;
        }

        WritePart("xl/worksheets/_rels/sheet1.xml.rels", writer =>
        {
            writer.WriteStartElement("Relationships", PackageRelationshipNamespace);
            for (int index = 1; index <= _tables.Count; index++)
            {
                WriteRelationship(
                    writer,
                    RelationshipId(index),
                    RelationshipTypeBase + "table",
                    $"../tables/{TablePartName(index)}");
            }

            writer.WriteEndElement();
        });

        for (int index = 1; index <= _tables.Count; index++)
        {
            TablePart table = _tables[index - 1];

            // A range left open at the bottom ends at the last row written, which is only known
            // now.
            string reference = new SpreadsheetAddress(
                table.FromRow,
                table.FromColumn,
                table.ToRow ?? GetOpenEndRow(table),
                table.ToColumn).Reference;

            WritePart($"xl/tables/{TablePartName(index)}", writer =>
            {
                writer.WriteStartElement("table", MainNamespace);
                writer.WriteAttributeString("id", index.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("name", table.Name);
                writer.WriteAttributeString("displayName", table.Name);
                writer.WriteAttributeString("ref", reference);
                writer.WriteAttributeString("totalsRowShown", "0");

                writer.WriteStartElement("autoFilter", MainNamespace);
                writer.WriteAttributeString("ref", reference);
                writer.WriteEndElement();

                writer.WriteStartElement("tableColumns", MainNamespace);
                writer.WriteAttributeString("count", table.Headings.Count.ToString(CultureInfo.InvariantCulture));
                for (int column = 0; column < table.Headings.Count; column++)
                {
                    writer.WriteStartElement("tableColumn", MainNamespace);
                    writer.WriteAttributeString("id", (column + 1).ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("name", table.Headings[column]);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();

                writer.WriteStartElement("tableStyleInfo", MainNamespace);
                writer.WriteAttributeString("name", table.StyleName);
                writer.WriteAttributeString("showFirstColumn", "0");
                writer.WriteAttributeString("showLastColumn", "0");
                writer.WriteAttributeString("showRowStripes", "1");
                writer.WriteAttributeString("showColumnStripes", "0");
                writer.WriteEndElement();

                writer.WriteEndElement();
            });
        }
    }

    /// <summary>
    /// The last row of a table whose range was left open at the bottom: the last row written, or
    /// the row above the next table down when there is one.
    /// </summary>
    /// <remarks>
    /// Two tables cannot overlap — Excel refuses the file rather than picking one — so a table
    /// left open above another one stops where that one starts. Without this, a sheet of stacked
    /// tables would give the first of them every row on the sheet. The header row is the floor: it
    /// has been written, or <see cref="Complete"/> refused.
    /// </remarks>
    private int GetOpenEndRow(TablePart table)
    {
        int lastRow = _currentRow;
        foreach (TablePart other in _tables)
        {
            if (other.FromRow > table.FromRow && other.FromRow - 1 < lastRow)
            {
                lastRow = other.FromRow - 1;
            }
        }

        return Math.Max(table.FromRow, lastRow);
    }

    /// <remarks>
    /// The fixed parts here — one border, two fills, one cell style — are the same base stylesheet
    /// <see cref="SpreadsheetStyleRepository"/> starts from, so a cell written by either writer
    /// resolves against the same defaults.
    /// </remarks>
    private void WriteStyles()
    {
        List<string> numberFormats = _styles
            .Select(style => style.NumberFormat)
            .OfType<string>()
            .Distinct(StringComparer.Ordinal)
            .ToList();

        List<(bool Bold, bool Italic)> fonts = new() { (false, false) };
        foreach (Style style in _styles)
        {
            if (!fonts.Contains((style.Bold, style.Italic)))
            {
                fonts.Add((style.Bold, style.Italic));
            }
        }

        WritePart("xl/styles.xml", writer =>
        {
            writer.WriteStartElement("styleSheet", MainNamespace);

            if (numberFormats.Count > 0)
            {
                writer.WriteStartElement("numFmts", MainNamespace);
                writer.WriteAttributeString("count", numberFormats.Count.ToString(CultureInfo.InvariantCulture));
                for (int index = 0; index < numberFormats.Count; index++)
                {
                    writer.WriteStartElement("numFmt", MainNamespace);
                    writer.WriteAttributeString("numFmtId", NumberFormatId(index));
                    writer.WriteAttributeString("formatCode", numberFormats[index]);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }

            writer.WriteStartElement("fonts", MainNamespace);
            writer.WriteAttributeString("count", fonts.Count.ToString(CultureInfo.InvariantCulture));
            foreach ((bool bold, bool italic) in fonts)
            {
                writer.WriteStartElement("font", MainNamespace);
                if (bold)
                {
                    writer.WriteElementString("b", MainNamespace, null);
                }

                if (italic)
                {
                    writer.WriteElementString("i", MainNamespace, null);
                }

                WriteValueElement(writer, "sz", "11");
                WriteValueElement(writer, "name", "Calibri");
                WriteValueElement(writer, "family", "2");
                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            writer.WriteRaw(
                "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
                + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
                + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");

            writer.WriteStartElement("cellXfs", MainNamespace);
            writer.WriteAttributeString("count", _styles.Count.ToString(CultureInfo.InvariantCulture));
            foreach (Style style in _styles)
            {
                int fontId = fonts.IndexOf((style.Bold, style.Italic));
                string numberFormatId = style.NumberFormat is null
                    ? "0"
                    : NumberFormatId(numberFormats.IndexOf(style.NumberFormat));

                writer.WriteStartElement("xf", MainNamespace);
                writer.WriteAttributeString("numFmtId", numberFormatId);
                writer.WriteAttributeString("fontId", fontId.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("fillId", "0");
                writer.WriteAttributeString("borderId", "0");
                writer.WriteAttributeString("xfId", "0");
                if (style.NumberFormat is not null)
                {
                    writer.WriteAttributeString("applyNumberFormat", "1");
                }

                if (fontId != 0)
                {
                    writer.WriteAttributeString("applyFont", "1");
                }

                if (style.WrapText)
                {
                    writer.WriteAttributeString("applyAlignment", "1");
                    writer.WriteStartElement("alignment", MainNamespace);
                    writer.WriteAttributeString("wrapText", "1");
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();

            writer.WriteRaw(
                "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
                + "<dxfs count=\"0\"/>"
                + "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>");

            writer.WriteEndElement();
        });
    }

    private void WriteWorkbook()
    {
        WritePart("xl/workbook.xml", writer =>
        {
            writer.WriteStartElement("workbook", MainNamespace);
            writer.WriteAttributeString("xmlns", "r", null, RelationshipNamespace);
            writer.WriteStartElement("sheets", MainNamespace);
            writer.WriteStartElement("sheet", MainNamespace);
            writer.WriteAttributeString("name", _sheetName);
            writer.WriteAttributeString("sheetId", "1");
            writer.WriteAttributeString("id", RelationshipNamespace, "rId1");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
        });

        WritePart("xl/_rels/workbook.xml.rels", writer =>
        {
            writer.WriteStartElement("Relationships", PackageRelationshipNamespace);
            WriteRelationship(writer, "rId1", RelationshipTypeBase + "worksheet", "worksheets/sheet1.xml");
            WriteRelationship(writer, "rId2", RelationshipTypeBase + "styles", "styles.xml");
            writer.WriteEndElement();
        });

        WritePart("_rels/.rels", writer =>
        {
            writer.WriteStartElement("Relationships", PackageRelationshipNamespace);
            WriteRelationship(writer, "rId1", RelationshipTypeBase + "officeDocument", "xl/workbook.xml");
            WriteRelationship(writer, "rId2", CorePropertiesRelationshipType, "docProps/core.xml");
            writer.WriteEndElement();
        });
    }

    private void WriteCoreProperties()
    {
        WritePart("docProps/core.xml", writer =>
        {
            writer.WriteStartElement("cp", "coreProperties", CorePropertiesNamespace);
            writer.WriteAttributeString("xmlns", "dc", null, DublinCoreNamespace);
            if (!string.IsNullOrEmpty(_author))
            {
                writer.WriteElementString("dc", "creator", DublinCoreNamespace, _author);
            }

            writer.WriteEndElement();
        });
    }

    private void WriteContentTypes()
    {
        WritePart("[Content_Types].xml", writer =>
        {
            writer.WriteStartElement("Types", ContentTypesNamespace);
            WriteDefaultContentType(writer, "rels", "application/vnd.openxmlformats-package.relationships+xml");
            WriteDefaultContentType(writer, "xml", "application/xml");
            WriteOverrideContentType(writer, "/xl/workbook.xml", ContentTypeBase + "sheet.main+xml");
            WriteOverrideContentType(writer, "/xl/worksheets/sheet1.xml", ContentTypeBase + "worksheet+xml");
            WriteOverrideContentType(writer, "/xl/styles.xml", ContentTypeBase + "styles+xml");
            for (int index = 1; index <= _tables.Count; index++)
            {
                WriteOverrideContentType(writer, $"/xl/tables/{TablePartName(index)}", ContentTypeBase + "table+xml");
            }

            WriteOverrideContentType(writer, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
            writer.WriteEndElement();
        });
    }

    private void WritePart(string path, Action<XmlWriter> write)
    {
        using Stream stream = _archive.CreateEntry(path, CompressionLevel.Optimal).Open();
        using XmlWriter writer = CreateXmlWriter(stream);
        writer.WriteStartDocument();
        write(writer);
        writer.WriteEndDocument();
    }

    private void EnsureNotCompleted()
    {
        if (_completed)
        {
            throw new InvalidOperationException("The workbook has already been completed.");
        }
    }

    private void ThrowIfUnknownStyle(uint styleIndex, string parameterName)
    {
        if (styleIndex >= (uint)_styles.Count)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                styleIndex,
                "A style index must come from GetStyleIndex.");
        }
    }

    private static void ThrowIfOutsideTheGrid(int value, string parameterName, int maximum)
    {
        if (value < 1 || value > maximum)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                value,
                $"{parameterName} must be between 1 and {maximum.ToString(CultureInfo.InvariantCulture)}.");
        }
    }

    private static void WriteInlineString(XmlWriter writer, string text)
    {
        text = SpreadsheetTextSanitizer.Sanitize(text);
        writer.WriteAttributeString("t", "inlineStr");
        writer.WriteStartElement("is", MainNamespace);
        writer.WriteStartElement("t", MainNamespace);

        // Excel drops leading and trailing whitespace from an inline string without this.
        if (text.Length != text.Trim().Length)
        {
            writer.WriteAttributeString("xml", "space", null, "preserve");
        }

        writer.WriteString(text);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteNumber(XmlWriter writer, string number)
    {
        writer.WriteElementString("v", MainNamespace, number);
    }

    private static XmlWriter CreateXmlWriter(Stream stream)
    {
        return XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,

            // Every string written here has been through SpreadsheetTextSanitizer, and the check
            // is a per-character cost on the hot path of a several-hundred-thousand-row export.
            CheckCharacters = false
        });
    }

    private static void WriteValueElement(XmlWriter writer, string name, string value)
    {
        writer.WriteStartElement(name, MainNamespace);
        writer.WriteAttributeString("val", value);
        writer.WriteEndElement();
    }

    private static void WriteRelationship(XmlWriter writer, string id, string type, string target)
    {
        writer.WriteStartElement("Relationship", PackageRelationshipNamespace);
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Type", type);
        writer.WriteAttributeString("Target", target);
        writer.WriteEndElement();
    }

    private static void WriteDefaultContentType(XmlWriter writer, string extension, string contentType)
    {
        writer.WriteStartElement("Default", ContentTypesNamespace);
        writer.WriteAttributeString("Extension", extension);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteOverrideContentType(XmlWriter writer, string partName, string contentType)
    {
        writer.WriteStartElement("Override", ContentTypesNamespace);
        writer.WriteAttributeString("PartName", partName);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static string RelationshipId(int index) => "rId" + index.ToString(CultureInfo.InvariantCulture);

    private static string TablePartName(int index) => $"table{index.ToString(CultureInfo.InvariantCulture)}.xml";

    // Number formats 0 to 163 are the ones Excel defines; a custom format is numbered from 164.
    private static string NumberFormatId(int index) => (164 + index).ToString(CultureInfo.InvariantCulture);

    private readonly record struct Style(bool Bold, bool Italic, string? NumberFormat, bool WrapText);

    private readonly record struct ColumnSettings(double? Width, bool Hidden, uint StyleIndex);

    private sealed record TablePart(
        int FromRow,
        int FromColumn,
        int? ToRow,
        int ToColumn,
        string Name,
        string StyleName,
        IReadOnlyList<string> Headings);
}
