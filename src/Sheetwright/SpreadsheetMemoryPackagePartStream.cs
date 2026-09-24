namespace Sheetwright;

internal sealed class SpreadsheetMemoryPackagePartStream : Stream
{
    private readonly Stream _stream;
    private readonly FileAccess _access;
    private bool _disposed;

    internal SpreadsheetMemoryPackagePartStream(Stream stream, FileAccess access)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _access = access;
    }

    public override bool CanRead => !_disposed && _access != FileAccess.Write && _stream.CanRead;

    public override bool CanSeek => !_disposed && _stream.CanSeek;

    public override bool CanWrite => !_disposed && _access != FileAccess.Read && _stream.CanWrite;

    public override long Length
    {
        get
        {
            ThrowIfDisposed();
            return _stream.Length;
        }
    }

    public override long Position
    {
        get
        {
            ThrowIfDisposed();
            return _stream.Position;
        }
        set
        {
            ThrowIfDisposed();
            _stream.Position = value;
        }
    }

    public override void Flush()
    {
        ThrowIfDisposed();
        _stream.Flush();
    }

    public override int Read(byte[] buffer, int offset, int count)
    {
        ThrowIfDisposed();
        return _stream.Read(buffer, offset, count);
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        ThrowIfDisposed();
        return _stream.Seek(offset, origin);
    }

    public override void SetLength(long value)
    {
        ThrowIfDisposed();
        _stream.SetLength(value);
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        ThrowIfDisposed();
        _stream.Write(buffer, offset, count);
    }

    protected override void Dispose(bool disposing)
    {
        _disposed = true;
        base.Dispose(disposing);
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(SpreadsheetMemoryPackagePartStream));
        }
    }
}
