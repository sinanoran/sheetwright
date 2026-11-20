using System.IO.Compression;
using System.Xml;

namespace Sheetwright;

/// <summary>
/// Reads the rows of a workbook someone else made, one at a time, within limits.
/// </summary>
/// <remarks>
/// <para>
/// <b>Not <see cref="SpreadsheetReader"/>.</b> That class reads back what Sheetwright wrote, into memory, by
/// cell address. This one reads any <c>.xlsx</c> a person can upload — Excel's, LibreOffice's, a script's — and
/// reads it as a stream: the package is opened without being loaded, a worksheet is advanced row by row, and
/// only the shared-string table is held. It is meant for files that cannot be trusted, which is why it is a
/// separate class with its own limits (<see cref="SpreadsheetReaderOptions"/>).
/// </para>
/// <para>
/// <b>What a cell holds.</b> Text, numbers (as <see cref="decimal"/>), booleans, errors and dates, where a date
/// is a number whose cell format is a date format. A formula is read as the value Excel stored with it and is
/// never evaluated; a formula with no stored value is empty. Formatting, merged cells, comments, drawings and
/// charts are ignored. Hidden sheets are listed (see <see cref="SpreadsheetSheetInfo.Visibility"/>) and
/// readable; whether to offer them is the caller's decision.
/// </para>
/// <para>
/// <b>What it refuses.</b> A file that is not a zip, or is not a workbook, is an
/// <see cref="InvalidDataException"/>. So is XML that is not well formed, one with a document type
/// declaration (no DTD is ever processed, so no entity is ever expanded) and a relationship that points outside
/// the package. A file past a limit is a <see cref="SpreadsheetReadLimitException"/>, named by the limit. The
/// size limits are checked against what the zip declares, to refuse an obvious bomb before reading it, and
/// again against the bytes that actually come out, because a header can lie.
/// </para>
/// <para>
/// The stream must be seekable: a zip's index is at its end. Not thread-safe. Rows are read lazily, so a
/// failure in a sheet's data surfaces while it is enumerated.
/// </para>
/// </remarks>
public sealed class SpreadsheetWorkbookReader : IDisposable
{
    private readonly ZipArchive _zip;
    private readonly Dictionary<string, ZipArchiveEntry> _entries;
    private readonly SpreadsheetReaderOptions _options;
    private readonly List<SpreadsheetSheetInfo> _sheets = [];
    private readonly string? _sharedStringsPart;
    private readonly string? _stylesPart;
    private readonly bool _date1904;
    private string[]? _sharedStrings;
    private bool[]? _dateStyles;
    private bool _disposed;

    private SpreadsheetWorkbookReader(
        ZipArchive zip,
        Dictionary<string, ZipArchiveEntry> entries,
        SpreadsheetReaderOptions options)
    {
        _zip = zip;
        _entries = entries;
        _options = options;

        string workbookPart = FindWorkbookPart();
        List<SpreadsheetOpc.Relationship> relationships = ReadRelationships(workbookPart);

        _sharedStringsPart = relationships
            .Where(r => !r.External && SpreadsheetOpc.IsType(r, SpreadsheetOpc.SharedStringsType))
            .Select(r => SpreadsheetOpc.Resolve(workbookPart, r.Target))
            .FirstOrDefault();
        _stylesPart = relationships
            .Where(r => !r.External && SpreadsheetOpc.IsType(r, SpreadsheetOpc.StylesType))
            .Select(r => SpreadsheetOpc.Resolve(workbookPart, r.Target))
            .FirstOrDefault();

        _date1904 = ReadSheets(workbookPart, relationships);
    }

    /// <summary>
    /// The worksheets of the workbook, in tab order. Chart sheets and other sheets with no cells are not listed.
    /// </summary>
    public IReadOnlyList<SpreadsheetSheetInfo> Sheets => _sheets;

    /// <summary>
    /// Opens a workbook. The package's index is read and checked against the limits, and the workbook part is
    /// read to list the sheets; no sheet's data is read yet.
    /// </summary>
    /// <param name="stream">The file, which must be readable and seekable.</param>
    /// <param name="options">The limits; <see cref="SpreadsheetReaderOptions.Default"/> when <see langword="null"/>.</param>
    /// <param name="leaveOpen">Whether disposing the reader leaves <paramref name="stream"/> open.</param>
    /// <exception cref="InvalidDataException">The file is not a workbook this reader can read.</exception>
    /// <exception cref="SpreadsheetReadLimitException">The package is past a limit.</exception>
    public static SpreadsheetWorkbookReader Open(Stream stream, SpreadsheetReaderOptions? options = null, bool leaveOpen = false)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanRead || !stream.CanSeek)
        {
            throw new ArgumentException(
                "A workbook is a zip file, which is read from its end: the stream must be readable and seekable. "
                + "Copy it to a file or a MemoryStream first.",
                nameof(stream));
        }

        options ??= SpreadsheetReaderOptions.Default;
        options.Validate();

        ZipArchive zip;
        try
        {
            zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen);
        }
        catch (InvalidDataException exception)
        {
            throw new InvalidDataException("The file is not an .xlsx workbook (it is not a zip file).", exception);
        }

        try
        {
            Dictionary<string, ZipArchiveEntry> entries = Index(zip, options);
            return new SpreadsheetWorkbookReader(zip, entries, options);
        }
        catch
        {
            zip.Dispose();
            throw;
        }
    }

    /// <summary>The rows of a worksheet, in order, one at a time.</summary>
    /// <param name="sheetIndex">The sheet's <see cref="SpreadsheetSheetInfo.Index"/>.</param>
    /// <exception cref="ArgumentOutOfRangeException">There is no such sheet.</exception>
    public IEnumerable<SpreadsheetRow> ReadRows(int sheetIndex)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (sheetIndex < 0 || sheetIndex >= _sheets.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(sheetIndex), "The workbook has no such worksheet.");
        }

        return ReadSheet(_sheets[sheetIndex]);
    }

    /// <summary>The rows of a worksheet, found by its name (letter case does not matter), one at a time.</summary>
    /// <exception cref="KeyNotFoundException">There is no such sheet.</exception>
    public IEnumerable<SpreadsheetRow> ReadRows(string sheetName)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentNullException.ThrowIfNull(sheetName);

        SpreadsheetSheetInfo sheet = _sheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException("The workbook has no worksheet with that name.");

        return ReadSheet(sheet);
    }

    /// <summary>Releases the package and, unless it was opened with <c>leaveOpen</c>, the stream.</summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _zip.Dispose();
    }

    private static Dictionary<string, ZipArchiveEntry> Index(ZipArchive zip, SpreadsheetReaderOptions options)
    {
        if (zip.Entries.Count > options.MaxParts)
        {
            throw new SpreadsheetReadLimitException(
                SpreadsheetReadLimit.MaxParts,
                options.MaxParts,
                "the package has more parts than that.");
        }

        Dictionary<string, ZipArchiveEntry> entries = new(StringComparer.OrdinalIgnoreCase);
        long declared = 0;

        foreach (ZipArchiveEntry entry in zip.Entries)
        {
            declared += entry.Length;
            if (declared > options.MaxUncompressedBytes)
            {
                throw new SpreadsheetReadLimitException(
                    SpreadsheetReadLimit.MaxUncompressedBytes,
                    options.MaxUncompressedBytes,
                    "the parts of the package declare more than that when decompressed.");
            }

            if (entry.Length > SpreadsheetBoundedPartStream.RatioFloorBytes
                && entry.Length > (long)options.MaxCompressionRatio * Math.Max(entry.CompressedLength, 1))
            {
                throw new SpreadsheetReadLimitException(
                    SpreadsheetReadLimit.MaxCompressionRatio,
                    options.MaxCompressionRatio,
                    $"part '{entry.FullName}' declares more than {options.MaxCompressionRatio} times its compressed size.");
            }

            string name = SpreadsheetOpc.Normalise(entry.FullName);
            if (name.Length == 0 || name.EndsWith('/'))
            {
                continue;
            }

            if (!entries.TryAdd(name, entry))
            {
                throw new InvalidDataException($"The package has two parts named '{name}'.");
            }
        }

        return entries;
    }

    private string FindWorkbookPart()
    {
        if (!_entries.ContainsKey("_rels/.rels"))
        {
            throw new InvalidDataException("The file is a zip but not an .xlsx workbook: it has no package relationships.");
        }

        List<SpreadsheetOpc.Relationship> relationships = ReadRelationshipsOf("_rels/.rels");
        SpreadsheetOpc.Relationship? document = relationships.FirstOrDefault(
            r => !r.External && SpreadsheetOpc.IsType(r, SpreadsheetOpc.OfficeDocumentType));

        string part = document is null
            ? throw new InvalidDataException("The file is a zip but not an .xlsx workbook: it names no office document.")
            : SpreadsheetOpc.Resolve(string.Empty, document.Target);

        if (!_entries.ContainsKey(part))
        {
            throw new InvalidDataException("The workbook part the package names is not in it.");
        }

        return part;
    }

    private List<SpreadsheetOpc.Relationship> ReadRelationships(string part)
    {
        string relationshipsPart = SpreadsheetOpc.RelationshipsPartOf(part);
        return _entries.ContainsKey(relationshipsPart) ? ReadRelationshipsOf(relationshipsPart) : [];
    }

    private List<SpreadsheetOpc.Relationship> ReadRelationshipsOf(string relationshipsPart)
    {
        using Stream stream = OpenPart(relationshipsPart);
        return SpreadsheetOpc.Read(stream, relationshipsPart);
    }

    private bool ReadSheets(string workbookPart, List<SpreadsheetOpc.Relationship> relationships)
    {
        bool date1904 = false;
        using Stream stream = OpenPart(workbookPart);
        using XmlReader reader = SpreadsheetXml.Create(stream);

        while (SpreadsheetXml.Read(reader, workbookPart))
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            if (reader.LocalName == "workbookPr")
            {
                date1904 = SpreadsheetXml.Flag(reader.GetAttribute("date1904"));
            }
            else if (reader.LocalName == "sheet" && reader.Depth == 2)
            {
                string? name = reader.GetAttribute("name");
                string? id = SpreadsheetXml.RelationshipId(reader);
                if (name is null || id is null)
                {
                    continue;
                }

                SpreadsheetOpc.Relationship? relationship = relationships.FirstOrDefault(r => r.Id == id);
                if (relationship is null || relationship.External || !SpreadsheetOpc.IsType(relationship, SpreadsheetOpc.WorksheetType))
                {
                    continue;
                }

                string part = SpreadsheetOpc.Resolve(workbookPart, relationship.Target);
                if (!_entries.ContainsKey(part))
                {
                    throw new InvalidDataException($"The workbook names a worksheet part ({part}) that is not in the package.");
                }

                SpreadsheetSheetVisibility visibility = reader.GetAttribute("state") switch
                {
                    "hidden" => SpreadsheetSheetVisibility.Hidden,
                    "veryHidden" => SpreadsheetSheetVisibility.VeryHidden,
                    _ => SpreadsheetSheetVisibility.Visible,
                };

                _sheets.Add(new SpreadsheetSheetInfo(name, _sheets.Count, visibility, part));
            }
        }

        return date1904;
    }

    private IEnumerable<SpreadsheetRow> ReadSheet(SpreadsheetSheetInfo sheet)
    {
        // The tables are loaded now, so a limit they break is raised by the call and not by the first row.
        string[] sharedStrings = LoadSharedStrings();
        bool[] dateStyles = LoadDateStyles();

        return ReadSheetRows(sheet, sharedStrings, dateStyles);
    }

    private IEnumerable<SpreadsheetRow> ReadSheetRows(SpreadsheetSheetInfo sheet, string[] sharedStrings, bool[] dateStyles)
    {
        using Stream stream = OpenPart(sheet.PartName);
        using XmlReader reader = SpreadsheetXml.Create(stream);

        SpreadsheetWorksheetRows rows = new(reader, sheet.PartName, _options, sharedStrings, dateStyles, _date1904);
        foreach (SpreadsheetRow row in rows.Read())
        {
            yield return row;
        }
    }

    private string[] LoadSharedStrings()
    {
        if (_sharedStrings is not null)
        {
            return _sharedStrings;
        }

        if (_sharedStringsPart is null || !_entries.ContainsKey(_sharedStringsPart))
        {
            return _sharedStrings = [];
        }

        using Stream stream = OpenPart(_sharedStringsPart);
        return _sharedStrings = SpreadsheetSharedStrings.Load(stream, _sharedStringsPart, _options);
    }

    private bool[] LoadDateStyles()
    {
        if (_dateStyles is not null)
        {
            return _dateStyles;
        }

        if (_stylesPart is null || !_entries.ContainsKey(_stylesPart))
        {
            return _dateStyles = [];
        }

        using Stream stream = OpenPart(_stylesPart);
        return _dateStyles = SpreadsheetStyleDates.Load(stream, _stylesPart);
    }

    private SpreadsheetBoundedPartStream OpenPart(string part)
    {
        ZipArchiveEntry entry = _entries[part];
        return new SpreadsheetBoundedPartStream(
            entry.Open(),
            part,
            entry.Length,
            entry.CompressedLength,
            _options);
    }
}
