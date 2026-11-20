namespace Sheetwright;

/// <summary>
/// A read-only stream that returns bytes already taken from a stream, then the rest of it: what lets the CSV
/// reader look at the start of a file to choose an encoding and a separator without needing to seek back.
/// </summary>
internal sealed class SpreadsheetPrefixedStream : Stream
{
    private readonly byte[] _prefix;
    private readonly Stream _inner;
    private readonly bool _leaveOpen;
    private int _position;
    private readonly int _end;

    internal SpreadsheetPrefixedStream(byte[] prefix, int offset, int count, Stream inner, bool leaveOpen)
    {
        _prefix = prefix;
        _position = offset;
        _end = offset + count;
        _inner = inner;
        _leaveOpen = leaveOpen;
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
        if (_position < _end)
        {
            int take = Math.Min(buffer.Length, _end - _position);
            _prefix.AsSpan(_position, take).CopyTo(buffer);
            _position += take;
            return take;
        }

        return _inner.Read(buffer);
    }

    public override void Flush()
    {
    }

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing && !_leaveOpen)
        {
            _inner.Dispose();
        }

        base.Dispose(disposing);
    }
}
