using System.IO.Packaging;

namespace Sheetwright;

internal sealed class SpreadsheetMemoryPackagePart : PackagePart
{
    private readonly MemoryStream _content = new();

    internal SpreadsheetMemoryPackagePart(
        Package package,
        Uri partUri,
        string contentType,
        CompressionOption compressionOption)
        : base(package, partUri, contentType, compressionOption)
    {
    }

    internal void CopyContentTo(Stream destination)
    {
        _content.Position = 0;
        _content.CopyTo(destination);
    }

    internal void DisposeContent()
    {
        _content.Dispose();
    }

    protected override Stream GetStreamCore(FileMode mode, FileAccess access)
    {
        switch (mode)
        {
            case FileMode.Create:
            case FileMode.Truncate:
                _content.SetLength(0);
                _content.Position = 0;
                break;
            case FileMode.Append:
                _content.Position = _content.Length;
                break;
            default:
                _content.Position = 0;
                break;
        }

        return new SpreadsheetMemoryPackagePartStream(_content, access);
    }
}
