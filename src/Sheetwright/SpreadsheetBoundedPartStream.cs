using System.Globalization;

namespace Sheetwright;

/// <summary>
/// A part of the package, read through a counter that stops it.
/// </summary>
/// <remarks>
/// <para>
/// A zip header says how big a part is, and a hostile file can lie: it can declare a kilobyte and deflate to a
/// gigabyte. So the declared sizes are used to refuse an obvious bomb early, and <em>this</em> counts the bytes
/// that actually come out, and stops the read the moment one of two things is true: the part has grown past
/// what its own header declared (the headers together were already checked against
/// <see cref="SpreadsheetReaderOptions.MaxUncompressedBytes"/> when the package was opened, so no part can be read
/// past its share of it), or it has expanded more than <see cref="SpreadsheetReaderOptions.MaxCompressionRatio"/>
/// times its compressed size (checked from <see cref="RatioFloorBytes"/> on, so a tiny part is not refused for a
/// ratio that means nothing).
/// </para>
/// <para>
/// It also watches for the one thing an XML reader cannot bound for itself: a single token with no end. Every
/// start tag, end tag and comment ends with <c>&gt;</c>, and text between them is a cell's, which is at most
/// <see cref="SpreadsheetReaderOptions.MaxCellCharacters"/> long. A run of bytes with no <c>&gt;</c> in it longer
/// than that is a tag with millions of attributes, or a comment or text that never ends, and reading it would
/// cost memory the other limits do not see.
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
    private long _sinceTagEnd;

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

        // Eight bytes a character covers UTF-16 and a surrogate pair in UTF-8, with room for the tag around it.
        _tokenLimit = Math.Max(1024L * 1024, 8L * options.MaxCellCharacters + 64 * 1024);
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

        int lastTagEnd = bytes.LastIndexOf((byte)'>');
        _sinceTagEnd = lastTagEnd >= 0 ? bytes.Length - lastTagEnd - 1 : _sinceTagEnd + bytes.Length;
        if (_sinceTagEnd > _tokenLimit)
        {
            throw new SpreadsheetReadLimitException(
                SpreadsheetReadLimit.MaxCellCharacters,
                _options.MaxCellCharacters,
                $"part '{_partName}' has a tag, a comment or a run of text that is longer than any cell can be.");
        }
    }
}
