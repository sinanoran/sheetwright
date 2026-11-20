using System.Globalization;

namespace Sheetwright;

/// <summary>
/// A part of the package, read through a counter that stops it.
/// </summary>
/// <remarks>
/// <para>
/// A zip header says how big a part is, and a hostile file can lie. The declared sizes are added up when the
/// package is opened, to refuse an obvious bomb before reading any of it, and <em>this</em> counts the bytes that
/// actually come out and stops the read the moment the part has grown past what its own header declared or has
/// expanded more than <see cref="SpreadsheetReaderOptions.MaxCompressionRatio"/> times its compressed size
/// (checked from <see cref="RatioFloorBytes"/> on, so a tiny part is not refused for a ratio that means
/// nothing).
/// </para>
/// <para>
/// It also bounds the one thing an XML reader cannot bound for itself: the size of a single token. A start tag
/// with millions of attributes, a comment, a CDATA section or a run of text that never ends is read whole into
/// memory by an XML reader, and none of the limits above sees it. So the bytes are followed through a small
/// state machine (text, tag, comment, CDATA, processing instruction), and a token that runs longer than any
/// cell can be (<see cref="SpreadsheetReaderOptions.MaxCellCharacters"/>, with room for the markup around it)
/// is refused. It follows quotes inside a tag, so a <c>&gt;</c> in an attribute value does not end one.
/// </para>
/// <para>
/// The state machine reads bytes, which is only sound where the characters that make up markup are single bytes:
/// UTF-8 and the single-byte encodings. A part that starts with a UTF-16 or UTF-32 byte-order mark, or is
/// written as UTF-16 without one, is refused. Every program that writes a workbook writes UTF-8.
/// </para>
/// </remarks>
internal sealed class SpreadsheetBoundedPartStream : Stream
{
    /// <summary>The ratio is not checked below this many bytes.</summary>
    internal const long RatioFloorBytes = 1024 * 1024;

    private readonly Stream _inner;
    private readonly string _partName;
    private readonly long _declaredLength;
    private readonly long _compressedLength;
    private readonly SpreadsheetReaderOptions _options;
    private readonly long _tokenLimit;
    private long _read;
    private TokenState _state = TokenState.Text;
    private char _quote;
    private long _token;
    private byte _previous;
    private byte _beforePrevious;
    private bool _checkedStart;

    internal SpreadsheetBoundedPartStream(
        Stream inner,
        string partName,
        long declaredLength,
        long compressedLength,
        SpreadsheetReaderOptions options)
    {
        _inner = inner;
        _partName = partName;
        _declaredLength = declaredLength;
        _compressedLength = compressedLength;
        _options = options;

        // Eight bytes a character covers a surrogate pair in UTF-8 and an entity for it, with room for the tag.
        _tokenLimit = Math.Max(1024L * 1024, (8L * options.MaxCellCharacters) + (64 * 1024));
    }

    private enum TokenState
    {
        Text,
        Open,
        Bang,
        Tag,
        Comment,
        CData,
        Instruction,
    }

    public override bool CanRead => true;

    public override bool CanSeek => false;

    public override bool CanWrite => false;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override int Read(byte[] buffer, int offset, int count) => Read(buffer.AsSpan(offset, count));

    public override int Read(Span<byte> buffer)
    {
        int read = _inner.Read(buffer);
        if (read > 0)
        {
            Account(buffer[..read]);
        }

        return read;
    }

    public override void Flush()
    {
    }

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _inner.Dispose();
        }

        base.Dispose(disposing);
    }

    private void Account(ReadOnlySpan<byte> bytes)
    {
        if (!_checkedStart)
        {
            _checkedStart = true;
            RefuseUtf16(bytes);
        }

        _read += bytes.Length;

        if (_read > _declaredLength)
        {
            throw new SpreadsheetReadLimitException(
                SpreadsheetReadLimit.MaxUncompressedBytes,
                _options.MaxUncompressedBytes,
                $"part '{_partName}' decompresses to more than the {_declaredLength.ToString(CultureInfo.InvariantCulture)} bytes its header declares.");
        }

        if (_read > RatioFloorBytes && _read > (long)_options.MaxCompressionRatio * Math.Max(_compressedLength, 1))
        {
            throw new SpreadsheetReadLimitException(
                SpreadsheetReadLimit.MaxCompressionRatio,
                _options.MaxCompressionRatio,
                $"part '{_partName}' expands to more than {_options.MaxCompressionRatio.ToString(CultureInfo.InvariantCulture)} times its compressed size.");
        }

        foreach (byte value in bytes)
        {
            Step(value);
            _beforePrevious = _previous;
            _previous = value;
        }
    }

    private void RefuseUtf16(ReadOnlySpan<byte> bytes)
    {
        bool bom = bytes.Length >= 2
            && ((bytes[0] == 0xFF && bytes[1] == 0xFE) || (bytes[0] == 0xFE && bytes[1] == 0xFF));
        bool nulAfterMarkup = bytes.Length >= 2
            && ((bytes[0] == 0x3C && bytes[1] == 0x00) || (bytes[0] == 0x00 && bytes[1] == 0x3C));

        if (bom || nulAfterMarkup)
        {
            throw new InvalidDataException(
                $"Part '{_partName}' is UTF-16 or UTF-32 XML. A workbook's parts must be UTF-8 to be read within limits.");
        }
    }

    // One byte of the part. The state is what the last '<' started; the counter is the size of that token so far.
    private void Step(byte value)
    {
        _token++;
        if (_token > _tokenLimit)
        {
            throw new SpreadsheetReadLimitException(
                SpreadsheetReadLimit.MaxCellCharacters,
                _options.MaxCellCharacters,
                $"part '{_partName}' has a tag, a comment or a run of text that is longer than any cell can be.");
        }

        switch (_state)
        {
            case TokenState.Text:
                if (value == (byte)'<')
                {
                    _state = TokenState.Open;
                    _token = 1;
                }

                break;

            case TokenState.Open:
                _state = value switch
                {
                    (byte)'!' => TokenState.Bang,
                    (byte)'?' => TokenState.Instruction,
                    _ => TokenState.Tag,
                };
                _quote = '\0';
                break;

            case TokenState.Bang:
                _state = value switch
                {
                    (byte)'-' => TokenState.Comment,
                    (byte)'[' => TokenState.CData,
                    _ => TokenState.Tag,
                };
                _quote = '\0';
                break;

            case TokenState.Tag:
                if (_quote != '\0')
                {
                    if (value == _quote)
                    {
                        _quote = '\0';
                    }
                }
                else if (value is (byte)'"' or (byte)'\'')
                {
                    _quote = (char)value;
                }
                else if (value == (byte)'>')
                {
                    EndToken();
                }

                break;

            case TokenState.Comment:
                if (value == (byte)'>' && _previous == (byte)'-' && _beforePrevious == (byte)'-')
                {
                    EndToken();
                }

                break;

            case TokenState.CData:
                if (value == (byte)'>' && _previous == (byte)']' && _beforePrevious == (byte)']')
                {
                    EndToken();
                }

                break;

            case TokenState.Instruction:
                if (value == (byte)'>' && _previous == (byte)'?')
                {
                    EndToken();
                }

                break;
        }
    }

    private void EndToken()
    {
        _state = TokenState.Text;
        _token = 0;
    }
}
