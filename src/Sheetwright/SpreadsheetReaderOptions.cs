using System.Text;

namespace Sheetwright;

/// <summary>
/// What <see cref="SpreadsheetWorkbookReader"/> and <see cref="SpreadsheetCsvReader"/> refuse to read, and how
/// the CSV reader is told to decode and split a file.
/// </summary>
/// <remarks>
/// <para>
/// <b>These readers take files from strangers.</b> A workbook is a zip, and a small one can expand into
/// gigabytes or name millions of strings; a CSV can be one field a gigabyte long. Every limit below is checked
/// <em>while</em> the file is read and before the memory for the content is allocated, and a file past one is
/// refused with a <see cref="SpreadsheetReadLimitException"/> that names it. A limit is never a truncation: a
/// reader that silently stopped would import half a file.
/// </para>
/// <para>
/// The defaults are the file format's own ceilings (Excel's 1,048,576 rows, 16,384 columns and 32,767
/// characters in a cell) and a 200 MB, 100:1, one-million-string workbook. A caller that knows its own ceiling
/// sets it lower. Instances are immutable; <see cref="Default"/> is shared.
/// </para>
/// </remarks>
public sealed class SpreadsheetReaderOptions
{
    private const int ExcelMaxRows = 1_048_576;
    private const int ExcelMaxColumns = 16_384;
    private const int ExcelMaxCellCharacters = 32_767;

    private int _maxRows = ExcelMaxRows;
    private int _maxColumns = ExcelMaxColumns;
    private int _maxCellCharacters = ExcelMaxCellCharacters;
    private long _maxUncompressedBytes = 200L * 1024 * 1024;
    private int _maxCompressionRatio = 100;
    private int _maxSharedStrings = 1_000_000;
    private long _maxSharedStringCharacters = 50_000_000;
    private int _maxParts = 1_000;

    /// <summary>The options a reader uses when it is given none.</summary>
    public static SpreadsheetReaderOptions Default { get; } = new();

    /// <summary>
    /// The most rows a reader returns (<see cref="SpreadsheetCsvReader"/> and
    /// <see cref="SpreadsheetWorkbookReader"/>). Rows that are skipped because they are empty do not count.
    /// Default 1,048,576.
    /// </summary>
    public int MaxRows
    {
        get => _maxRows;
        init => _maxRows = Positive(value, nameof(MaxRows));
    }

    /// <summary>The most columns in one row. Default 16,384.</summary>
    public int MaxColumns
    {
        get => _maxColumns;
        init => _maxColumns = Positive(value, nameof(MaxColumns));
    }

    /// <summary>
    /// The most characters in one cell. A longer cell is refused while it is being read, never truncated.
    /// Default 32,767, the most Excel itself holds.
    /// </summary>
    public int MaxCellCharacters
    {
        get => _maxCellCharacters;
        init => _maxCellCharacters = Positive(value, nameof(MaxCellCharacters));
    }

    /// <summary>
    /// Whether rows with no non-empty cell are returned. Default <see langword="false"/>: a spreadsheet often
    /// has thousands of formatted but empty rows below its data, and a CSV a blank line at the end. A skipped
    /// row still has its place in the numbering: <see cref="SpreadsheetRow.RowNumber"/> is the row's own
    /// number in the file.
    /// </summary>
    public bool IncludeEmptyRows { get; init; }

    /// <summary>
    /// Workbooks: the most bytes all the parts of the package may decompress to. The sizes the zip declares are
    /// added up when the package is opened, and no part can be read past the size it declares, so what is
    /// decompressed never exceeds this however many times a sheet is read. Default 200 MB.
    /// </summary>
    public long MaxUncompressedBytes
    {
        get => _maxUncompressedBytes;
        init => _maxUncompressedBytes = Positive(value, nameof(MaxUncompressedBytes));
    }

    /// <summary>
    /// Workbooks: the most any one part may expand to, as a multiple of its compressed size (uncompressed ÷
    /// compressed). Checked from the first mebibyte on, as the part is decompressed, so a bomb stops at the
    /// ratio and not at the total size. Default 100; spreadsheet XML compresses 5 to 20 times.
    /// </summary>
    public int MaxCompressionRatio
    {
        get => _maxCompressionRatio;
        init => _maxCompressionRatio = Positive(value, nameof(MaxCompressionRatio));
    }

    /// <summary>Workbooks: the most entries in the shared-string table. Default 1,000,000.</summary>
    public int MaxSharedStrings
    {
        get => _maxSharedStrings;
        init => _maxSharedStrings = Positive(value, nameof(MaxSharedStrings));
    }

    /// <summary>
    /// Workbooks: the most characters in the whole shared-string table. Default 50,000,000, which is what
    /// the table costs in memory once it is loaded (two bytes each).
    /// </summary>
    public long MaxSharedStringCharacters
    {
        get => _maxSharedStringCharacters;
        init => _maxSharedStringCharacters = Positive(value, nameof(MaxSharedStringCharacters));
    }

    /// <summary>Workbooks: the most entries (parts) in the package. Default 1,000.</summary>
    public int MaxParts
    {
        get => _maxParts;
        init => _maxParts = Positive(value, nameof(MaxParts));
    }

    /// <summary>
    /// CSV: the encoding of the file. <see langword="null"/> (the default) means: the byte-order mark if there
    /// is one (UTF-8, UTF-16 little or big endian), otherwise UTF-8. A file with bytes that are not valid in
    /// the encoding is refused with an <see cref="InvalidDataException"/> rather than read with replacement
    /// characters; the caller may then offer the person another encoding. A legacy code page needs
    /// <see cref="CodePagesEncodingProvider"/> registered by the caller.
    /// </summary>
    public Encoding? Encoding { get; init; }

    /// <summary>
    /// CSV: the field separator. <see langword="null"/> (the default) means a leading <c>sep=;</c> line if
    /// <see cref="HonourSepDirective"/> is on, otherwise the one of comma, semicolon, tab and pipe that splits
    /// the first rows into the same number of fields most consistently.
    /// </summary>
    public char? Delimiter { get; init; }

    /// <summary>
    /// CSV: whether a first line of exactly <c>sep=</c> and one character sets the separator and is not
    /// returned as a row (Excel's own convention). Default <see langword="true"/>.
    /// </summary>
    public bool HonourSepDirective { get; init; } = true;

    /// <summary>
    /// CSV: whether one leading apostrophe is removed from a field when the character after it is one of
    /// <c>= + - @</c>, tab, carriage return or the full-width forms of the first four. That is exactly the
    /// prefix <see cref="SpreadsheetGridCsv"/> adds to defuse a spreadsheet formula, so a file this library
    /// wrote reads back as the values that were written. Any other apostrophe is data and is kept. Default
    /// <see langword="true"/>. See <see cref="SpreadsheetGridCsv"/> for the one text this cannot tell apart
    /// from a guarded one.
    /// </summary>
    public bool UnescapeFormulaGuard { get; init; } = true;

    internal void Validate()
    {
        if (Delimiter is '"' or '\r' or '\n')
        {
            throw new ArgumentException(
                "The delimiter cannot be a quote or a line break.", nameof(Delimiter));
        }
    }

    private static int Positive(int value, string name)
    {
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(value, 0, name);
        return value;
    }

    private static long Positive(long value, string name)
    {
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(value, 0L, name);
        return value;
    }
}
