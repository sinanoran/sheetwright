using System.Text;

namespace Sheetwright;

/// <summary>
/// Reads a CSV file one row at a time, whatever machine wrote it.
/// </summary>
/// <remarks>
/// <para>
/// <b>What it accepts.</b> RFC 4180 quoting (a quoted field may hold the separator, a doubled quote and line
/// breaks), records ended by CRLF, LF or a lone CR, a last record with no line break, and ragged rows. It is
/// lenient where Excel is: a quote in the middle of an unquoted field is a character, and text after a closing
/// quote is kept. It refuses what cannot be read faithfully: a quoted field still open at the end of the file,
/// and bytes that are not valid in the encoding.
/// </para>
/// <para>
/// <b>Encoding and separator.</b> The start of the file (64 KiB) is looked at once, to choose the encoding
/// (a UTF-8, UTF-16 LE or UTF-16 BE byte-order mark; otherwise UTF-8) and the separator (a <c>sep=</c> line,
/// otherwise the most consistent of comma, semicolon, tab and pipe); both can be set in
/// <see cref="SpreadsheetReaderOptions"/>, and are readable here after opening so that a screen can show what
/// was chosen and let the person change it. The file is read forward once and never held whole, and the
/// stream need not be seekable.
/// </para>
/// <para>
/// <b>Every cell is text</b> (or <see cref="SpreadsheetCellKind.Empty"/> when it has no characters): a CSV
/// has no types, and what <c>12.480,00</c> means depends on the culture of whoever wrote it. The one change
/// made to a value is the apostrophe rule of <see cref="SpreadsheetReaderOptions.UnescapeFormulaGuard"/>.
/// </para>
/// <para>Not thread-safe. Dispose it to release the stream.</para>
/// </remarks>
public sealed class SpreadsheetCsvReader : IDisposable
{
    private const int SniffBytes = 64 * 1024;
    private const int BufferChars = 8 * 1024;

    private readonly StreamReader _reader;
    private readonly SpreadsheetReaderOptions _options;
    private readonly bool _skipDirectiveLine;
    private readonly char[] _buffer = new char[BufferChars];
    private int _bufferPosition;
    private int _bufferLength;
    private bool _consumed;
    private bool _disposed;

    private SpreadsheetCsvReader(
        StreamReader reader,
        SpreadsheetReaderOptions options,
        Encoding encoding,
        char delimiter,
        bool skipDirectiveLine)
    {
        _reader = reader;
        _options = options;
        Encoding = encoding;
        Delimiter = delimiter;
        _skipDirectiveLine = skipDirectiveLine;
    }

    /// <summary>The encoding the file is read in.</summary>
    public Encoding Encoding { get; }

    /// <summary>The separator the file is read with.</summary>
    public char Delimiter { get; }

    /// <summary>
    /// Opens a CSV file. Only the first 64 KiB is read here; nothing else is read until
    /// <see cref="ReadRows"/> is enumerated.
    /// </summary>
    /// <param name="stream">The file. It need not be seekable.</param>
    /// <param name="options">The limits and the settings; <see cref="SpreadsheetReaderOptions.Default"/> when <see langword="null"/>.</param>
    /// <param name="leaveOpen">Whether disposing the reader leaves <paramref name="stream"/> open.</param>
    /// <exception cref="InvalidDataException">The file starts with a UTF-32 byte-order mark, which is not supported.</exception>
    public static SpreadsheetCsvReader Open(Stream stream, SpreadsheetReaderOptions? options = null, bool leaveOpen = false)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanRead)
        {
            throw new ArgumentException("The stream cannot be read.", nameof(stream));
        }

        options ??= SpreadsheetReaderOptions.Default;
        options.Validate();

        byte[] sniff = new byte[SniffBytes];
        int count = ReadFully(stream, sniff);
        bool endOfFile = count < SniffBytes;

        (Encoding encoding, int preamble) = ChooseEncoding(sniff, count, options.Encoding);

        string sample = DecodeLeniently(encoding, sniff, preamble, count - preamble);

        bool hasDirective = false;
        char directive = ',';
        if (options.HonourSepDirective)
        {
            hasDirective = SpreadsheetCsvDelimiter.TryReadDirective(sample, out directive);
        }

        char delimiter = options.Delimiter
            ?? (hasDirective ? directive : SpreadsheetCsvDelimiter.Detect(sample, truncated: !endOfFile));

        SpreadsheetPrefixedStream joined = new(sniff, preamble, count - preamble, stream, leaveOpen);
        StreamReader reader = new(joined, encoding, detectEncodingFromByteOrderMarks: false, bufferSize: 16 * 1024);

        return new SpreadsheetCsvReader(reader, options, encoding, delimiter, hasDirective);
    }

    /// <summary>
    /// The rows of the file, in order, one at a time. A file can be enumerated once.
    /// </summary>
    /// <remarks>
    /// Rows with no content are skipped unless <see cref="SpreadsheetReaderOptions.IncludeEmptyRows"/> is on;
    /// they still use up a <see cref="SpreadsheetRow.RowNumber"/>. A <c>sep=</c> line is not a row and has no
    /// number.
    /// </remarks>
    /// <exception cref="SpreadsheetReadLimitException">A row, a column or a cell is past its limit.</exception>
    /// <exception cref="InvalidDataException">The file cannot be read: bytes invalid in the encoding, or a quote left open.</exception>
    public IEnumerable<SpreadsheetRow> ReadRows()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (_consumed)
        {
            throw new InvalidOperationException("A CSV file is read once; open it again to read it again.");
        }

        _consumed = true;
        return Enumerate();
    }

    /// <summary>Releases the stream, unless it was opened with <c>leaveOpen</c>.</summary>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _reader.Dispose();
    }

    private IEnumerable<SpreadsheetRow> Enumerate()
    {
        List<string> fields = [];
        StringBuilder field = new();
        int recordNumber = 0;
        int returned = 0;

        if (_skipDirectiveLine)
        {
            SkipLine();
        }

        while (true)
        {
            recordNumber++;
            if (!TryReadRecord(recordNumber, fields, field))
            {
                yield break;
            }

            SpreadsheetCellValue[] cells = new SpreadsheetCellValue[fields.Count];
            bool empty = true;
            for (int index = 0; index < cells.Length; index++)
            {
                if (fields[index].Length > 0)
                {
                    cells[index] = SpreadsheetCellValue.FromText(fields[index]);
                    empty = false;
                }
            }

            if (empty && !_options.IncludeEmptyRows)
            {
                continue;
            }

            if (returned >= _options.MaxRows)
            {
                throw new SpreadsheetReadLimitException(
                    SpreadsheetReadLimit.MaxRows,
                    _options.MaxRows,
                    $"there are more rows than that (row {recordNumber} is the first one past it).");
            }

            returned++;
            yield return new SpreadsheetRow(recordNumber, cells);
        }
    }

    private bool TryReadRecord(int recordNumber, List<string> fields, StringBuilder field)
    {
        try
        {
            return ReadRecord(recordNumber, fields, field);
        }
        catch (DecoderFallbackException exception)
        {
            throw new InvalidDataException(
                $"The file is not valid {Encoding.WebName} text (a byte sequence it cannot decode, near row {recordNumber}). "
                + "Read it in the encoding it was saved in.",
                exception);
        }
    }

    private bool ReadRecord(int recordNumber, List<string> fields, StringBuilder field)
    {
        fields.Clear();
        field.Clear();

        int next = NextChar();
        if (next < 0)
        {
            return false;
        }

        bool inQuotes = false;
        bool fieldStarted = false;

        while (true)
        {
            if (next < 0)
            {
                if (inQuotes)
                {
                    throw new InvalidDataException(
                        $"A quoted field that starts in row {recordNumber} is never closed.");
                }

                EndField(recordNumber, fields, field);
                return true;
            }

            char character = (char)next;

            if (inQuotes)
            {
                if (character == '"')
                {
                    if (PeekChar() == '"')
                    {
                        NextChar();
                        Append(recordNumber, field, '"');
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    Append(recordNumber, field, character);
                }
            }
            else if (character == '"' && !fieldStarted)
            {
                inQuotes = true;
                fieldStarted = true;
            }
            else if (character == Delimiter)
            {
                EndField(recordNumber, fields, field);
                fieldStarted = false;
            }
            else if (character is '\r' or '\n')
            {
                if (character == '\r' && PeekChar() == '\n')
                {
                    NextChar();
                }

                EndField(recordNumber, fields, field);
                return true;
            }
            else
            {
                fieldStarted = true;
                Append(recordNumber, field, character);
            }

            next = NextChar();
        }
    }

    private void Append(int recordNumber, StringBuilder field, char character)
    {
        if (field.Length >= _options.MaxCellCharacters)
        {
            throw new SpreadsheetReadLimitException(
                SpreadsheetReadLimit.MaxCellCharacters,
                _options.MaxCellCharacters,
                $"a cell in row {recordNumber} has more characters than that.");
        }

        field.Append(character);
    }

    private void EndField(int recordNumber, List<string> fields, StringBuilder field)
    {
        if (fields.Count >= _options.MaxColumns)
        {
            throw new SpreadsheetReadLimitException(
                SpreadsheetReadLimit.MaxColumns,
                _options.MaxColumns,
                $"row {recordNumber} has more columns than that.");
        }

        string text = field.ToString();
        field.Clear();

        if (_options.UnescapeFormulaGuard
            && text.Length >= 2
            && text[0] == '\''
            && SpreadsheetFormulaGuard.IsLeader(text[1]))
        {
            text = text[1..];
        }

        fields.Add(text);
    }

    private void SkipLine()
    {
        try
        {
            int next;
            while ((next = NextChar()) >= 0)
            {
                if (next == '\n')
                {
                    return;
                }

                if (next == '\r')
                {
                    if (PeekChar() == '\n')
                    {
                        NextChar();
                    }

                    return;
                }
            }
        }
        catch (DecoderFallbackException exception)
        {
            throw new InvalidDataException(
                $"The file is not valid {Encoding.WebName} text. Read it in the encoding it was saved in.",
                exception);
        }
    }

    // The reader is not asked to Peek: on a stream that cannot seek, StreamReader.Peek answers -1 whenever its
    // buffer happens to be empty, which would end the file early. This keeps a buffer of its own.
    private int NextChar()
    {
        if (_bufferPosition == _bufferLength && !Fill())
        {
            return -1;
        }

        return _buffer[_bufferPosition++];
    }

    private int PeekChar()
    {
        if (_bufferPosition == _bufferLength && !Fill())
        {
            return -1;
        }

        return _buffer[_bufferPosition];
    }

    private bool Fill()
    {
        _bufferLength = _reader.Read(_buffer, 0, _buffer.Length);
        _bufferPosition = 0;
        return _bufferLength > 0;
    }

    private static int ReadFully(Stream stream, byte[] buffer)
    {
        int total = 0;
        while (total < buffer.Length)
        {
            int read = stream.Read(buffer, total, buffer.Length - total);
            if (read == 0)
            {
                break;
            }

            total += read;
        }

        return total;
    }

    private static (Encoding Encoding, int Preamble) ChooseEncoding(byte[] bytes, int count, Encoding? requested)
    {
        if (requested is not null)
        {
            // Whether or not the encoding writes a mark of its own, a file in its family may start with one,
            // and a mark is not part of the first cell.
            int mark = requested.CodePage switch
            {
                65001 when count >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF => 3,
                1200 when count >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE => 2,
                1201 when count >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF => 2,
                _ => 0,
            };

            ReadOnlySpan<byte> preamble = requested.Preamble;
            if (mark == 0 && preamble.Length > 0 && count >= preamble.Length && bytes.AsSpan(0, preamble.Length).SequenceEqual(preamble))
            {
                mark = preamble.Length;
            }

            return (requested, mark);
        }

        if (count >= 4
            && ((bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0 && bytes[3] == 0)
                || (bytes[0] == 0 && bytes[1] == 0 && bytes[2] == 0xFE && bytes[3] == 0xFF)))
        {
            throw new InvalidDataException(
                "The file is UTF-32, which is not supported. Save it as UTF-8 or UTF-16.");
        }

        if (count >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            return (new UTF8Encoding(false, true), 3);
        }

        if (count >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
        {
            return (new UnicodeEncoding(false, false, true), 2);
        }

        if (count >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
        {
            return (new UnicodeEncoding(true, false, true), 2);
        }

        return (new UTF8Encoding(false, true), 0);
    }

    // For choosing a separator only: the sample may end half way through a character, and a replacement
    // character there changes nothing. The real read is strict.
    private static string DecodeLeniently(Encoding encoding, byte[] bytes, int offset, int count)
    {
        Encoding lenient = (Encoding)encoding.Clone();
        lenient.DecoderFallback = DecoderFallback.ReplacementFallback;
        return lenient.GetString(bytes, offset, count);
    }
}
