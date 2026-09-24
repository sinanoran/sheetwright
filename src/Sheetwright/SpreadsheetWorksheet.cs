using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Sheetwright;

public sealed class SpreadsheetWorksheet
{
    // Applied to a cell written as a DateTime that the caller never gave an explicit format.
    private const string ImplicitDateNumberFormat = "dd/mm/yyyy";

    private readonly SpreadsheetPackage _package;
    private readonly WorksheetPart _worksheetPart;
    private readonly Sheet _sheet;
    private readonly Dictionary<uint, Row> _rows = new();
    private readonly Dictionary<long, Cell> _cells = new();
    private readonly Dictionary<long, SpreadsheetStylePatch> _cellStyles = new();
    private readonly HashSet<long> _dateCells = new();
    private readonly Dictionary<uint, int> _maximumColumnsByRow = new();
    private readonly Dictionary<int, SpreadsheetStylePatch> _rowStyles = new();
    private readonly Dictionary<int, SpreadsheetStylePatch> _columnStyles = new();
    private readonly Dictionary<int, SpreadsheetColumnSettings> _columnSettings = new();
    private readonly SpreadsheetStylePatch _defaultStyle = new();
    private readonly List<string> _mergedRanges = new();
    private readonly List<SpreadsheetTableDefinition> _tables = new();
    private SpreadsheetSheetProtection? _protection;
    private Dictionary<int, double>? _autoFitWidths;
    private IReadOnlyCollection<SpreadsheetPicture>? _pictures;
    private double? _defaultColumnWidth;
    private uint _maximumRowIndex;
    private int _minimumCellRow = int.MaxValue;
    private int _minimumCellColumn = int.MaxValue;
    private int _maximumCellRow;
    private int _maximumCellColumn;

    private Worksheet OpenXmlWorksheet => _worksheetPart.Worksheet ?? throw new InvalidDataException("The worksheet is missing.");

    internal SpreadsheetWorksheet(SpreadsheetPackage package, WorksheetPart worksheetPart, Sheet sheet)
    {
        _package = package;
        _worksheetPart = worksheetPart;
        _sheet = sheet;
        Cells = new SpreadsheetCells(this);
        Columns = new SpreadsheetColumns(this);
        View = new SpreadsheetWorksheetView(this);
        Drawings = new SpreadsheetDrawingCollection(this);
        DataValidations = new SpreadsheetDataValidationCollection();
        Tables = new SpreadsheetTables(this);

        SheetData sheetData = EnsureSheetData();
        foreach (Row row in sheetData.Elements<Row>())
        {
            if (row.RowIndex?.Value is uint rowIndex)
            {
                _rows[rowIndex] = row;
                _maximumRowIndex = Math.Max(_maximumRowIndex, rowIndex);
            }

            foreach (Cell cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value is string reference)
                {
                    SpreadsheetAddress address = SpreadsheetAddress.Parse(reference);
                    long cellKey = GetCellKey(address.FromRow, address.FromColumn);
                    _cells[cellKey] = cell;
                    RegisterCell(address.FromRow, address.FromColumn);
                }
            }
        }
    }

    public string Name => _sheet.Name?.Value ?? string.Empty;

    public SpreadsheetCells Cells { get; }

    public SpreadsheetColumns Columns { get; }

    public SpreadsheetWorksheetView View { get; }

    public SpreadsheetDrawingCollection Drawings { get; }

    public SpreadsheetDataValidationCollection DataValidations { get; }

    public SpreadsheetTables Tables { get; }

    public SpreadsheetWorksheetVisibility Visibility
    {
        get
        {
            SheetStateValues? state = _sheet.State?.Value;
            if (state == SheetStateValues.Hidden)
            {
                return SpreadsheetWorksheetVisibility.Hidden;
            }
            if (state == SheetStateValues.VeryHidden)
            {
                return SpreadsheetWorksheetVisibility.VeryHidden;
            }
            return SpreadsheetWorksheetVisibility.Visible;
        }
        set => _sheet.State = value switch
        {
            SpreadsheetWorksheetVisibility.Hidden => SheetStateValues.Hidden,
            SpreadsheetWorksheetVisibility.VeryHidden => SheetStateValues.VeryHidden,
            _ => SheetStateValues.Visible
        };
    }

    public double DefaultColumnWidth
    {
        get => _defaultColumnWidth
            ?? OpenXmlWorksheet.GetFirstChild<SheetFormatProperties>()?.DefaultColumnWidth?.Value
            ?? 8.43D;
        set
        {
            _defaultColumnWidth = value;
            SheetFormatProperties properties = OpenXmlWorksheet.GetFirstChild<SheetFormatProperties>()
                ?? OpenXmlWorksheet.InsertBefore(new SheetFormatProperties { DefaultRowHeight = 15D }, EnsureSheetData());
            properties.DefaultColumnWidth = value;
        }
    }

    public SpreadsheetDimension? Dimension
    {
        get
        {
            if (_cells.Count == 0)
            {
                return null;
            }

            return new SpreadsheetDimension(
                _minimumCellRow,
                _minimumCellColumn,
                _maximumCellRow,
                _maximumCellColumn);
        }
    }

    public SpreadsheetColumn Column(int column)
    {
        return Columns[column];
    }

    public void Select(string address)
    {
        if (address is null)
        {
            throw new ArgumentNullException(nameof(address));
        }

        SelectedRange = Cells[address];
    }

    public SpreadsheetRange? SelectedRange { get; private set; }

    public void Protect(bool lockEverything, string password)
    {
        if (string.IsNullOrWhiteSpace(password))
        {
            throw new ArgumentException("A worksheet-protection password is required.", nameof(password));
        }

        _protection = new SpreadsheetSheetProtection(lockEverything, password);
    }

    internal void SetCodeName(string codeName)
    {
        SheetProperties properties = OpenXmlWorksheet.GetFirstChild<SheetProperties>()
            ?? OpenXmlWorksheet.InsertAt(new SheetProperties(), 0);
        properties.CodeName = SpreadsheetTextSanitizer.Sanitize(codeName);
    }

    internal object? GetValue(int row, int column)
    {
        if (!_cells.TryGetValue(GetCellKey(row, column), out Cell? cell))
        {
            return null;
        }

        return _package.ReadCellValue(cell);
    }

    internal string? GetFormula(int row, int column)
    {
        return _cells.TryGetValue(GetCellKey(row, column), out Cell? cell) ? cell.CellFormula?.Text : null;
    }

    internal void SetValue(int row, int column, object? value)
    {
        _autoFitWidths = null;
        long cellKey = GetCellKey(row, column);
        _dateCells.Remove(cellKey);
        Cell cell = GetOrCreateCell(row, column);
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
                cell.CellValue = new CellValue(_package.GetSharedStringIndex(text).ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.SharedString;
                break;

            case char character:
                cell.CellValue = new CellValue(_package.GetSharedStringIndex(character.ToString()).ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.SharedString;
                break;

            case DateTime dateTime:
                cell.CellValue = new CellValue(dateTime.ToOADate().ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
                _dateCells.Add(cellKey);
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
                // Invariant, like every other conversion here. A workbook is a
                // file rather than a screen: the same object written on a Turkish
                // machine and an American one has to produce the same bytes, or an
                // export becomes a function of who ran it. A caller who wants a
                // value rendered for a particular reader formats it themselves and
                // passes a string.
                string converted = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                cell.CellValue = new CellValue(_package.GetSharedStringIndex(converted).ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.SharedString;
                break;
        }
    }

    internal void SetFormula(int row, int column, string? formula)
    {
        _autoFitWidths = null;
        Cell cell = GetOrCreateCell(row, column);
        if (formula is null || formula.Length == 0)
        {
            cell.CellFormula = null;
            return;
        }

        _dateCells.Remove(GetCellKey(row, column));
        cell.CellValue = null;
        cell.InlineString = null;
        cell.CellFormula = new CellFormula(SpreadsheetTextSanitizer.Sanitize(formula.TrimStart('=')));
        cell.DataType = null;
    }

    internal void ApplyDefaultStyle(SpreadsheetStylePatch patch)
    {
        _defaultStyle.Merge(patch);
    }

    internal void ApplyColumnStyle(int column, SpreadsheetStylePatch patch)
    {
        if (!_columnStyles.TryGetValue(column, out SpreadsheetStylePatch? existing))
        {
            existing = new SpreadsheetStylePatch();
            _columnStyles[column] = existing;
        }
        existing.Merge(patch);
    }

    internal void ApplyStyle(SpreadsheetAddress address, SpreadsheetStylePatch patch)
    {
        if (address.IsWholeColumn)
        {
            for (int column = address.FromColumn; column <= address.ToColumn; column++)
            {
                ApplyColumnStyle(column, patch);
            }
            return;
        }

        if (address.IsWholeRow)
        {
            for (int row = address.FromRow; row <= address.ToRow; row++)
            {
                ApplyRowStyle(row, patch);
            }
            return;
        }

        for (int row = address.FromRow; row <= address.ToRow; row++)
        {
            for (int column = address.FromColumn; column <= address.ToColumn; column++)
            {
                ApplyCellStyle(row, column, patch);
            }
        }
    }

    internal void ApplyBorderAround(SpreadsheetAddress address, SpreadsheetBorderStyle borderStyle, string color)
    {
        if (address.IsWholeColumn || address.IsWholeRow)
        {
            throw new InvalidOperationException("Borders around whole rows or columns are not supported.");
        }

        for (int row = address.FromRow; row <= address.ToRow; row++)
        {
            for (int column = address.FromColumn; column <= address.ToColumn; column++)
            {
                ApplyCellStyle(row, column, new SpreadsheetStylePatch
                {
                    LeftBorder = column == address.FromColumn ? borderStyle : null,
                    RightBorder = column == address.ToColumn ? borderStyle : null,
                    TopBorder = row == address.FromRow ? borderStyle : null,
                    BottomBorder = row == address.ToRow ? borderStyle : null,
                    BorderColor = color
                });
            }
        }
    }

    internal void SetMerged(SpreadsheetAddress address, bool merged)
    {
        if (merged)
        {
            if (!_mergedRanges.Contains(address.Reference, StringComparer.OrdinalIgnoreCase))
            {
                _mergedRanges.Add(address.Reference);
            }
            return;
        }

        _mergedRanges.RemoveAll(x => string.Equals(x, address.Reference, StringComparison.OrdinalIgnoreCase));
    }

    internal bool IsMerged(SpreadsheetAddress address)
    {
        if (_mergedRanges.Contains(address.Reference, StringComparer.OrdinalIgnoreCase))
        {
            return true;
        }

        return OpenXmlWorksheet.GetFirstChild<MergeCells>()?.Elements<MergeCell>()
            .Any(merge => string.Equals(merge.Reference?.Value, address.Reference, StringComparison.OrdinalIgnoreCase)) == true;
    }

    internal void FreezePanes(int row, int column)
    {
        // sheetViews follows sheetPr in CT_Worksheet, and the schema order is not
        // advisory: Excel offers to repair a file that gets it wrong. Inserting at 0
        // is only correct while no sheetPr exists, and SetVbaProject writes one for
        // every sheet — so a workbook that attached its macros before freezing a
        // pane produced a file that would not open.
        SheetViews views = OpenXmlWorksheet.GetFirstChild<SheetViews>()
            ?? (OpenXmlWorksheet.GetFirstChild<SheetProperties>() is SheetProperties sheetProperties
                ? OpenXmlWorksheet.InsertAfter(new SheetViews(), sheetProperties)
                : OpenXmlWorksheet.InsertAt(new SheetViews(), 0));
        SheetView view = views.Elements<SheetView>().FirstOrDefault()
            ?? views.AppendChild(new SheetView { WorkbookViewId = 0U });

        string topLeftCell = SpreadsheetAddress.GetCellReference(row, column);
        view.Pane = new Pane
        {
            VerticalSplit = row > 1 ? row - 1 : 0,
            HorizontalSplit = column > 1 ? column - 1 : 0,
            TopLeftCell = topLeftCell,
            ActivePane = row > 1
                ? column > 1 ? PaneValues.BottomRight : PaneValues.BottomLeft
                : column > 1 ? PaneValues.TopRight : PaneValues.TopLeft,
            State = PaneStateValues.Frozen
        };
    }

    internal void SetColumnWidth(int column, double width)
    {
        GetColumnSettings(column).Width = width;
    }

    internal void SetColumnHidden(int column, bool hidden)
    {
        GetColumnSettings(column).Hidden = hidden;
    }

    internal void AutoFitColumn(int column)
    {
        EnsureAutoFitWidths();
        double minimumWidth = _defaultColumnWidth ?? 8.43D;
        double contentWidth = _autoFitWidths!.TryGetValue(column, out double width) ? width : minimumWidth;
        SetColumnWidth(column, Math.Min(Math.Max(contentWidth, minimumWidth), 80D));
    }

    internal void AutoFitColumns()
    {
        int maxColumn = Dimension?.End.Column ?? 0;
        for (int column = 1; column <= maxColumn; column++)
        {
            AutoFitColumn(column);
        }
    }

    internal SpreadsheetTableDefinition AddTable(SpreadsheetRange range, string name, string styleName)
    {
        SpreadsheetTableDefinition definition = new(
            range.Address,
            SpreadsheetTextSanitizer.Sanitize(name),
            SpreadsheetTextSanitizer.Sanitize(styleName));
        _tables.Add(definition);
        return definition;
    }

    internal Xdr.OneCellAnchor AddPicture(string name, Stream image, int row, int column, int pixelTop, int pixelLeft)
    {
        _pictures = null;
        name = SpreadsheetTextSanitizer.Sanitize(name);
        DrawingsPart drawingsPart = _worksheetPart.DrawingsPart ?? _worksheetPart.AddNewPart<DrawingsPart>();
        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

        if (OpenXmlWorksheet.GetFirstChild<Drawing>() is null)
        {
            OpenXmlWorksheet.Append(new Drawing { Id = _worksheetPart.GetIdOfPart(drawingsPart) });
        }

        (PartTypeInfo imageType, long width, long height) = GetImageInfo(image);
        image.Position = 0;
        ImagePart imagePart = drawingsPart.AddImagePart(imageType);
        imagePart.FeedData(image);
        string relationshipId = drawingsPart.GetIdOfPart(imagePart);
        uint drawingId = (uint)(drawingsPart.WorksheetDrawing.Descendants<Xdr.Picture>().Count() + 1);

        Xdr.Picture picture = new(
            new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties { Id = drawingId, Name = name, Description = name },
                new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
            new Xdr.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())),
            new Xdr.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = width, Cy = height }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        Xdr.OneCellAnchor anchor = new(
            CreateFromMarker(row, column, pixelTop, pixelLeft),
            new Xdr.Extent { Cx = width, Cy = height },
            picture,
            new Xdr.ClientData());

        drawingsPart.WorksheetDrawing.Append(anchor);
        drawingsPart.WorksheetDrawing.Save();
        return anchor;
    }

    internal void MovePicture(Xdr.OneCellAnchor anchor, int row, int column, int pixelTop, int pixelLeft)
    {
        _pictures = null;
        anchor.FromMarker = CreateFromMarker(row, column, pixelTop, pixelLeft);
        _worksheetPart.DrawingsPart?.WorksheetDrawing?.Save();
    }

    private static Xdr.FromMarker CreateFromMarker(int row, int column, int pixelTop, int pixelLeft)
    {
        return new Xdr.FromMarker(
            new Xdr.ColumnId((column - 1).ToString(CultureInfo.InvariantCulture)),
            new Xdr.ColumnOffset((pixelLeft * 9525L).ToString(CultureInfo.InvariantCulture)),
            new Xdr.RowId((row - 1).ToString(CultureInfo.InvariantCulture)),
            new Xdr.RowOffset((pixelTop * 9525L).ToString(CultureInfo.InvariantCulture)));
    }

    internal IReadOnlyCollection<SpreadsheetPicture> ReadPictures()
    {
        if (_pictures is not null)
        {
            return _pictures;
        }

        DrawingsPart? drawingsPart = _worksheetPart.DrawingsPart;
        if (drawingsPart?.WorksheetDrawing is null)
        {
            _pictures = Array.Empty<SpreadsheetPicture>();
            return _pictures;
        }

        List<SpreadsheetPicture> pictures = new();
        foreach (Xdr.Picture picture in drawingsPart.WorksheetDrawing.Descendants<Xdr.Picture>())
        {
            A.Blip? blip = picture.BlipFill?.Blip;
            if (blip?.Embed?.Value is not string relationshipId)
            {
                continue;
            }

            OpenXmlPart part = drawingsPart.GetPartById(relationshipId);
            if (part is not ImagePart imagePart)
            {
                continue;
            }

            using Stream stream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            using MemoryStream memory = new();
            stream.CopyTo(memory);

            Xdr.NonVisualDrawingProperties? properties = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
            string name = properties?.Name?.Value ?? string.Empty;
            string title = properties?.Title?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(title))
            {
                title = properties?.Description?.Value ?? string.Empty;
            }
            if (string.IsNullOrWhiteSpace(title))
            {
                title = name;
            }

            int row = GetPictureRow(picture);
            pictures.Add(new SpreadsheetPicture(name, title, memory.ToArray(), GetImageFormat(imagePart.ContentType), row));
        }

        _pictures = pictures.ToArray();
        return _pictures;
    }

    internal void Commit(SpreadsheetStyleRepository styles)
    {
        foreach (int rowIndex in _rowStyles.Keys)
        {
            GetOrCreateRow(rowIndex);
        }

        foreach (long cellKey in _cellStyles.Keys)
        {
            GetOrCreateCell(GetCellRow(cellKey), GetCellColumn(cellKey));
        }

        foreach (KeyValuePair<long, Cell> entry in _cells)
        {
            long cellKey = entry.Key;
            Cell cell = entry.Value;
            int row = GetCellRow(cellKey);
            int column = GetCellColumn(cellKey);
            SpreadsheetStylePatch combined = _defaultStyle.Clone();
            if (_columnStyles.TryGetValue(column, out SpreadsheetStylePatch? columnStyle))
            {
                combined.Merge(columnStyle);
            }
            if (_rowStyles.TryGetValue(row, out SpreadsheetStylePatch? rowStyle))
            {
                combined.Merge(rowStyle);
            }
            if (_cellStyles.TryGetValue(cellKey, out SpreadsheetStylePatch? cellStyle))
            {
                combined.Merge(cellStyle);
            }
            if (_dateCells.Contains(cellKey) && combined.NumberFormat is null)
            {
                combined.NumberFormat = ImplicitDateNumberFormat;
            }

            cell.StyleIndex = styles.GetStyleIndex(combined);
        }

        CommitColumns(styles);
        CommitRows(styles);
        CommitMerges();
        DataValidations.Commit(OpenXmlWorksheet);
        CommitProtection();
        CommitTables();

        SpreadsheetDimension? dimension = Dimension;
        if (dimension is not null)
        {
            SheetDimension sheetDimension = new() { Reference = dimension.Reference };
            SheetDimension? existing = OpenXmlWorksheet.GetFirstChild<SheetDimension>();
            if (existing is null)
            {
                SheetProperties? sheetProperties = OpenXmlWorksheet.GetFirstChild<SheetProperties>();
                if (sheetProperties is null)
                {
                    OpenXmlWorksheet.InsertAt(sheetDimension, 0);
                }
                else
                {
                    OpenXmlWorksheet.InsertAfter(sheetDimension, sheetProperties);
                }
            }
            else
            {
                existing.Reference = dimension.Reference;
            }
        }

        OpenXmlWorksheet.Save();
    }

    private SheetData EnsureSheetData()
    {
        return OpenXmlWorksheet.GetFirstChild<SheetData>()
            ?? OpenXmlWorksheet.AppendChild(new SheetData());
    }

    private Cell GetOrCreateCell(int rowNumber, int columnNumber)
    {
        long cellKey = GetCellKey(rowNumber, columnNumber);
        if (_cells.TryGetValue(cellKey, out Cell? existing))
        {
            return existing;
        }

        uint rowIndex = (uint)rowNumber;
        Row row = GetOrCreateRow(rowNumber);

        Cell cell = new() { CellReference = SpreadsheetAddress.GetCellReference(rowNumber, columnNumber) };
        if (!_maximumColumnsByRow.TryGetValue(rowIndex, out int maximumColumn) || columnNumber > maximumColumn)
        {
            row.Append(cell);
        }
        else
        {
            Cell? nextCell = row.Elements<Cell>()
                .FirstOrDefault(x => SpreadsheetAddress.Parse(x.CellReference!.Value!).FromColumn > columnNumber);
            if (nextCell is null)
            {
                row.Append(cell);
            }
            else
            {
                row.InsertBefore(cell, nextCell);
            }
        }

        _cells[cellKey] = cell;
        RegisterCell(rowNumber, columnNumber);
        return cell;
    }

    private Row GetOrCreateRow(int rowNumber)
    {
        uint rowIndex = (uint)rowNumber;
        if (_rows.TryGetValue(rowIndex, out Row? existing))
        {
            return existing;
        }

        Row row = new() { RowIndex = rowIndex };
        if (_rows.Count == 0 || rowIndex > _maximumRowIndex)
        {
            EnsureSheetData().Append(row);
        }
        else
        {
            Row? nextRow = null;
            uint nextRowIndex = uint.MaxValue;
            foreach (KeyValuePair<uint, Row> candidate in _rows)
            {
                if (candidate.Key > rowIndex && candidate.Key < nextRowIndex)
                {
                    nextRow = candidate.Value;
                    nextRowIndex = candidate.Key;
                }
            }

            if (nextRow is null)
            {
                EnsureSheetData().Append(row);
            }
            else
            {
                EnsureSheetData().InsertBefore(row, nextRow);
            }
        }

        _rows[rowIndex] = row;
        _maximumRowIndex = Math.Max(_maximumRowIndex, rowIndex);
        return row;
    }

    private void ApplyCellStyle(int row, int column, SpreadsheetStylePatch patch)
    {
        long cellKey = GetCellKey(row, column);
        if (!_cellStyles.TryGetValue(cellKey, out SpreadsheetStylePatch? existing))
        {
            existing = new SpreadsheetStylePatch();
            _cellStyles[cellKey] = existing;
        }
        existing.Merge(patch);
    }

    private void ApplyRowStyle(int row, SpreadsheetStylePatch patch)
    {
        if (!_rowStyles.TryGetValue(row, out SpreadsheetStylePatch? existing))
        {
            existing = new SpreadsheetStylePatch();
            _rowStyles[row] = existing;
        }
        existing.Merge(patch);
    }

    private SpreadsheetColumnSettings GetColumnSettings(int column)
    {
        if (!_columnSettings.TryGetValue(column, out SpreadsheetColumnSettings? settings))
        {
            settings = new SpreadsheetColumnSettings();
            _columnSettings[column] = settings;
        }
        return settings;
    }

    private void CommitColumns(SpreadsheetStyleRepository styles)
    {
        if (_columnSettings.Count == 0 && _columnStyles.Count == 0)
        {
            return;
        }

        Columns columns = OpenXmlWorksheet.GetFirstChild<Columns>()
            ?? OpenXmlWorksheet.InsertBefore(new Columns(), EnsureSheetData());
        columns.RemoveAllChildren<Column>();

        foreach (int index in _columnSettings.Keys.Union(_columnStyles.Keys).OrderBy(x => x))
        {
            _columnSettings.TryGetValue(index, out SpreadsheetColumnSettings? settings);
            SpreadsheetStylePatch style = _defaultStyle.Clone();
            if (_columnStyles.TryGetValue(index, out SpreadsheetStylePatch? columnStyle))
            {
                style.Merge(columnStyle);
            }

            Column column = new()
            {
                Min = (uint)index,
                Max = (uint)index,
                CustomWidth = settings?.Width is not null,
                Width = settings?.Width,
                Hidden = settings?.Hidden,
                Style = styles.GetStyleIndex(style)
            };
            columns.Append(column);
        }
    }

    private void CommitRows(SpreadsheetStyleRepository styles)
    {
        foreach (KeyValuePair<int, SpreadsheetStylePatch> entry in _rowStyles)
        {
            SpreadsheetStylePatch style = _defaultStyle.Clone();
            style.Merge(entry.Value);
            Row row = GetOrCreateRow(entry.Key);
            row.StyleIndex = styles.GetStyleIndex(style);
            row.CustomFormat = true;
        }
    }

    private void CommitMerges()
    {
        if (_mergedRanges.Count == 0)
        {
            OpenXmlWorksheet.GetFirstChild<MergeCells>()?.Remove();
            return;
        }

        MergeCells mergeCells = OpenXmlWorksheet.GetFirstChild<MergeCells>()
            ?? OpenXmlWorksheet.InsertAfter(new MergeCells(), EnsureSheetData());
        mergeCells.RemoveAllChildren<MergeCell>();
        foreach (string range in _mergedRanges.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            mergeCells.Append(new MergeCell { Reference = range });
        }
        mergeCells.Count = (uint)mergeCells.ChildElements.Count;
    }

    private void CommitProtection()
    {
        if (_protection is null)
        {
            return;
        }

        OpenXmlWorksheet.RemoveAllChildren<SheetProtection>();
        SheetProtection protection = new()
        {
            Sheet = true,
            DeleteColumns = true,
            DeleteRows = true,
            InsertColumns = true,
            InsertRows = true,
            PivotTables = true,
            SelectLockedCells = true,
            Sort = _protection.LockEverything,
            AutoFilter = _protection.LockEverything,
            FormatCells = _protection.LockEverything,
            FormatColumns = _protection.LockEverything,
            FormatRows = _protection.LockEverything,
            InsertHyperlinks = _protection.LockEverything,
            SelectUnlockedCells = _protection.LockEverything
        };

        protection.Password = HashPassword(_protection.Password);

        OpenXmlWorksheet.InsertAfter(protection, EnsureSheetData());
    }

    private void CommitTables()
    {
        foreach (SpreadsheetTableDefinition definition in _tables)
        {
            TableDefinitionPart tablePart = _worksheetPart.AddNewPart<TableDefinitionPart>();
            SpreadsheetAddress address = SpreadsheetAddress.Parse(definition.Reference);
            Table table = new()
            {
                Id = _package.NextTableId(),
                Name = definition.Name,
                DisplayName = definition.Name,
                Reference = definition.Reference,
                TotalsRowShown = false,
                AutoFilter = new AutoFilter { Reference = definition.Reference }
            };

            TableColumns tableColumns = new() { Count = (uint)(address.ToColumn - address.FromColumn + 1) };
            HashSet<string> headings = new(StringComparer.OrdinalIgnoreCase);
            for (int column = address.FromColumn; column <= address.ToColumn; column++)
            {
                string existing = Convert.ToString(GetValue(address.FromRow, column), CultureInfo.InvariantCulture)
                    ?? string.Empty;
                string heading = GetUniqueTableHeading(
                    string.IsNullOrWhiteSpace(existing) ? SpreadsheetAddress.GetColumnName(column) : existing,
                    headings);

                // Excel checks a table's column names against the cells of its
                // header row and offers to repair the file when they disagree, so
                // a heading that had to be invented for an empty cell or made
                // unique against a duplicate has to go back into the cell. Doing
                // this here is safe because the style pass has already run and
                // the dimension has not: a header cell created now is styled with
                // the default and still counted in the sheet's extent.
                if (!string.Equals(existing, heading, StringComparison.Ordinal))
                {
                    SetValue(address.FromRow, column, heading);
                }

                tableColumns.Append(new TableColumn
                {
                    Id = (uint)(column - address.FromColumn + 1),
                    Name = heading
                });
            }
            table.Append(tableColumns);
            table.Append(new TableStyleInfo
            {
                Name = definition.StyleName,
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false
            });
            tablePart.Table = table;
            tablePart.Table.Save();
        }

        if (_tables.Count > 0)
        {
            TableParts tableParts = OpenXmlWorksheet.GetFirstChild<TableParts>() ?? new TableParts();
            tableParts.RemoveAllChildren<TablePart>();
            foreach (TableDefinitionPart part in _worksheetPart.TableDefinitionParts)
            {
                tableParts.Append(new TablePart { Id = _worksheetPart.GetIdOfPart(part) });
            }
            tableParts.Count = (uint)tableParts.ChildElements.Count;
            if (tableParts.Parent is null)
            {
                OpenXmlWorksheet.Append(tableParts);
            }
        }
    }

    /// <remarks>
    /// Excel's own password hash, and it is sixteen bits: collisions are trivial
    /// to find and the algorithm has been public for decades. Worth naming
    /// because the method is called <c>HashPassword</c> and looks like a security
    /// control. It is not one — it stops a colleague editing the total row by
    /// accident, and nothing else. Anything that must not be read belongs
    /// somewhere other than a worksheet.
    /// </remarks>
    private static string HashPassword(string password)
    {
        ushort hash = 0;
        if (password.Length > 0)
        {
            for (int index = password.Length - 1; index >= 0; index--)
            {
                hash = (ushort)(((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff));
                hash ^= password[index];
            }
            hash = (ushort)(((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff));
            hash ^= (ushort)password.Length;
            hash ^= 0xCE4B;
        }
        return hash.ToString("X4", CultureInfo.InvariantCulture);
    }

    private static string GetUniqueTableHeading(string heading, HashSet<string> headings)
    {
        if (headings.Add(heading))
        {
            return heading;
        }

        int suffix = 2;
        string candidate;
        do
        {
            candidate = $"{heading}{suffix.ToString(CultureInfo.InvariantCulture)}";
            suffix++;
        }
        while (!headings.Add(candidate));

        return candidate;
    }

    private static int GetPictureRow(Xdr.Picture picture)
    {
        OpenXmlElement? anchor = picture.Parent;
        string? row = anchor switch
        {
            Xdr.OneCellAnchor one => one.FromMarker?.RowId?.Text,
            Xdr.TwoCellAnchor two => two.FromMarker?.RowId?.Text,
            _ => null
        };
        return int.TryParse(row, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result) ? result : 0;
    }

    private void EnsureAutoFitWidths()
    {
        if (_autoFitWidths is not null)
        {
            return;
        }

        _autoFitWidths = new Dictionary<int, double>();
        foreach (KeyValuePair<long, Cell> entry in _cells)
        {
            int column = GetCellColumn(entry.Key);
            // Invariant so that a column width is a property of the workbook and
            // not of the machine that wrote it. It is a measurement rather than a
            // rendering, and a date whose invariant form is two characters longer
            // would otherwise make the same export differ between two developers.
            string value = Convert.ToString(_package.ReadCellValue(entry.Value), CultureInfo.InvariantCulture) ?? string.Empty;
            double width = value.Length + 2D;
            if (!_autoFitWidths.TryGetValue(column, out double existing) || width > existing)
            {
                _autoFitWidths[column] = width;
            }
        }
    }

    private void RegisterCell(int row, int column)
    {
        _minimumCellRow = Math.Min(_minimumCellRow, row);
        _minimumCellColumn = Math.Min(_minimumCellColumn, column);
        _maximumCellRow = Math.Max(_maximumCellRow, row);
        _maximumCellColumn = Math.Max(_maximumCellColumn, column);

        uint rowIndex = (uint)row;
        if (!_maximumColumnsByRow.TryGetValue(rowIndex, out int maximumColumn) || column > maximumColumn)
        {
            _maximumColumnsByRow[rowIndex] = column;
        }
    }

    private static long GetCellKey(int row, int column)
    {
        return (long)row << 32 | (uint)column;
    }

    private static int GetCellRow(long cellKey)
    {
        return (int)(cellKey >> 32);
    }

    private static int GetCellColumn(long cellKey)
    {
        return (int)cellKey;
    }

    private static (PartTypeInfo Type, long Width, long Height) GetImageInfo(Stream stream)
    {
        byte[] header = new byte[32];
        int read = ReadUpTo(stream, header, header.Length);
        stream.Position = 0;

        if (read >= 8 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
        {
            long width = read >= 24 ? ReadBigEndianInt32(header, 16) * 9525L : 180L * 9525L;
            long height = read >= 24 ? ReadBigEndianInt32(header, 20) * 9525L : 60L * 9525L;
            return (ImagePartType.Png, width, height);
        }
        if (read >= 3 && header[0] == 0xFF && header[1] == 0xD8 && header[2] == 0xFF)
        {
            (int width, int height) = ReadJpegDimensions(stream);
            return (ImagePartType.Jpeg, width * 9525L, height * 9525L);
        }
        if (read >= 10 && header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46)
        {
            int width = header[6] | header[7] << 8;
            int height = header[8] | header[9] << 8;
            return (ImagePartType.Gif, width * 9525L, height * 9525L);
        }
        if (read >= 4 && ((header[0] == 0x49 && header[1] == 0x49) || (header[0] == 0x4D && header[1] == 0x4D)))
        {
            (int width, int height) = ReadTiffDimensions(stream);
            return (ImagePartType.Tiff, width * 9525L, height * 9525L);
        }
        if (read >= 26 && header[0] == 0x42 && header[1] == 0x4D)
        {
            int width = Math.Abs(BitConverter.ToInt32(header, 18));
            int height = Math.Abs(BitConverter.ToInt32(header, 22));
            return (ImagePartType.Bmp, width * 9525L, height * 9525L);
        }
        throw new InvalidDataException("The image format is not supported.");
    }

    private static (int Width, int Height) ReadJpegDimensions(Stream stream)
    {
        stream.Position = 2;
        while (stream.Position < stream.Length)
        {
            int prefix = stream.ReadByte();
            if (prefix != 0xFF)
            {
                continue;
            }

            int marker;
            do
            {
                marker = stream.ReadByte();
            }
            while (marker == 0xFF);

            if (marker < 0 || marker is 0xD9 or 0xDA)
            {
                break;
            }
            if (marker is 0xD8 or >= 0xD0 and <= 0xD7)
            {
                continue;
            }

            int segmentLength = ReadBigEndianUInt16(stream);
            if (segmentLength < 2)
            {
                break;
            }

            bool isStartOfFrame = marker is 0xC0 or 0xC1 or 0xC2 or 0xC3
                or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;
            if (isStartOfFrame && segmentLength >= 7)
            {
                stream.ReadByte();
                int height = ReadBigEndianUInt16(stream);
                int width = ReadBigEndianUInt16(stream);
                stream.Position = 0;
                return (Math.Max(width, 1), Math.Max(height, 1));
            }

            stream.Seek(segmentLength - 2, SeekOrigin.Current);
        }

        stream.Position = 0;
        return (180, 60);
    }

    private static (int Width, int Height) ReadTiffDimensions(Stream stream)
    {
        stream.Position = 0;
        byte[] header = new byte[8];
        if (!TryRead(stream, header, header.Length))
        {
            return (180, 60);
        }

        bool littleEndian = header[0] == 0x49 && header[1] == 0x49;
        bool bigEndian = header[0] == 0x4D && header[1] == 0x4D;
        if ((!littleEndian && !bigEndian) || ReadUInt16(header, 2, littleEndian) != 42)
        {
            stream.Position = 0;
            return (180, 60);
        }

        uint directoryOffset = ReadUInt32(header, 4, littleEndian);
        if (directoryOffset > stream.Length - 2)
        {
            stream.Position = 0;
            return (180, 60);
        }

        stream.Position = directoryOffset;
        byte[] countBuffer = new byte[2];
        if (!TryRead(stream, countBuffer, countBuffer.Length))
        {
            stream.Position = 0;
            return (180, 60);
        }

        int entryCount = ReadUInt16(countBuffer, 0, littleEndian);
        int? width = null;
        int? height = null;
        byte[] entry = new byte[12];
        for (int index = 0; index < entryCount && stream.Position <= stream.Length - entry.Length; index++)
        {
            if (!TryRead(stream, entry, entry.Length))
            {
                break;
            }

            ushort tag = ReadUInt16(entry, 0, littleEndian);
            if (tag is not 256 and not 257)
            {
                continue;
            }

            ushort type = ReadUInt16(entry, 2, littleEndian);
            uint valueCount = ReadUInt32(entry, 4, littleEndian);
            if (valueCount != 1 || type is not 3 and not 4)
            {
                continue;
            }

            uint value = type == 3
                ? ReadUInt16(entry, 8, littleEndian)
                : ReadUInt32(entry, 8, littleEndian);
            if (value == 0 || value > int.MaxValue)
            {
                continue;
            }

            if (tag == 256)
            {
                width = (int)value;
            }
            else
            {
                height = (int)value;
            }
        }

        stream.Position = 0;
        return (width ?? 180, height ?? 60);
    }

    private static bool TryRead(Stream stream, byte[] buffer, int count)
    {
        int offset = 0;
        while (offset < count)
        {
            int read = stream.Read(buffer, offset, count - offset);
            if (read == 0)
            {
                return false;
            }
            offset += read;
        }
        return true;
    }

    private static int ReadUpTo(Stream stream, byte[] buffer, int count)
    {
        int offset = 0;
        while (offset < count)
        {
            int read = stream.Read(buffer, offset, count - offset);
            if (read == 0)
            {
                break;
            }
            offset += read;
        }
        return offset;
    }

    private static ushort ReadUInt16(byte[] value, int offset, bool littleEndian)
    {
        return littleEndian
            ? (ushort)(value[offset] | value[offset + 1] << 8)
            : (ushort)(value[offset] << 8 | value[offset + 1]);
    }

    private static uint ReadUInt32(byte[] value, int offset, bool littleEndian)
    {
        return littleEndian
            ? (uint)(value[offset] | value[offset + 1] << 8 | value[offset + 2] << 16 | value[offset + 3] << 24)
            : (uint)(value[offset] << 24 | value[offset + 1] << 16 | value[offset + 2] << 8 | value[offset + 3]);
    }

    private static int ReadBigEndianUInt16(Stream stream)
    {
        int high = stream.ReadByte();
        int low = stream.ReadByte();
        return high < 0 || low < 0 ? 0 : high << 8 | low;
    }

    private static int ReadBigEndianInt32(byte[] value, int offset)
    {
        return (value[offset] << 24)
            | (value[offset + 1] << 16)
            | (value[offset + 2] << 8)
            | value[offset + 3];
    }

    private static string GetImageFormat(string contentType)
    {
        return contentType switch
        {
            "image/png" => "Png",
            "image/jpeg" => "Jpeg",
            "image/gif" => "Gif",
            "image/tiff" => "Tiff",
            "image/bmp" => "Bmp",
            "image/x-emf" or "image/emf" => "Emf",
            "image/x-wmf" or "image/wmf" => "Wmf",
            _ => contentType
        };
    }
}
