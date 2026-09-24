using System.IO.Packaging;

namespace Sheetwright;

internal sealed class SpreadsheetMemoryPackage : Package
{
    private readonly Dictionary<string, SpreadsheetMemoryPackagePart> _parts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Stream _destination;

    internal SpreadsheetMemoryPackage(Stream destination)
        : base(FileAccess.ReadWrite)
    {
        _destination = destination ?? throw new ArgumentNullException(nameof(destination));
    }

    protected override PackagePart CreatePartCore(
        Uri partUri,
        string contentType,
        CompressionOption compressionOption)
    {
        SpreadsheetMemoryPackagePart part = new(this, partUri, contentType, compressionOption);
        _parts.Add(GetPartKey(partUri), part);
        return part;
    }

    protected override void DeletePartCore(Uri partUri)
    {
        string key = GetPartKey(partUri);
        if (_parts.TryGetValue(key, out SpreadsheetMemoryPackagePart? part))
        {
            part.DisposeContent();
            _parts.Remove(key);
        }
    }

    protected override PackagePart GetPartCore(Uri partUri)
    {
        // Package.PartExists expects the core lookup to return null when a part is absent.
        return _parts.TryGetValue(GetPartKey(partUri), out SpreadsheetMemoryPackagePart? part)
            ? part
            : null!;
    }

    protected override PackagePart[] GetPartsCore()
    {
        return _parts.Values.Cast<PackagePart>().ToArray();
    }

    protected override void FlushCore()
    {
        SpreadsheetPackageArchiveWriter.Write(_destination, _parts.Values);
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);

        if (disposing)
        {
            foreach (SpreadsheetMemoryPackagePart part in _parts.Values)
            {
                part.DisposeContent();
            }
        }
    }

    private static string GetPartKey(Uri partUri)
    {
        return partUri.OriginalString;
    }
}
